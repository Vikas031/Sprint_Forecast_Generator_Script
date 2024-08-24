import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import os
import warnings
import getpass
import sys
import time
import threading

# Suppress all warnings
warnings.filterwarnings("ignore")

jira_url = "https://jira.iongroup.com"

import pandas as pd

def read_excel_file(file_path):
    try:
        # Load the Excel file
        df = pd.read_excel(file_path, engine='openpyxl')
        return df
    except FileNotFoundError:
        print(f"Error: The file at path '{file_path}' was not found.")
        return None
    except Exception as e:
        print(f"Error loading file: {e}")
        return None


def read_excel_data(df, column_letter, start_row, end_row):
    try:
        # Convert column letter to column index
        column_index = ord(column_letter.upper()) - ord('A')
        if column_index < 0 or column_index >= len(df.columns):
            raise ValueError(f"Invalid column letter '{column_letter}'. Column out of range.")
        
        # Slice the DataFrame to get the specific rows and columns
        data = df.iloc[start_row-1:end_row, column_index:column_index+2]
        
        # Check if the row indices are within the bounds of the DataFrame
        if start_row < 1 or end_row > len(df):
            raise ValueError("Row indices are out of range.")

        return data

    except IndexError as e:
        print(f"Error processing data: {e}")
        return None
    except ValueError as e:
        print(f"ValueError: {e}")
        return None
    except Exception as e:
        print(f"Unexpected error: {e}")
        return None


def fetch_jira_summary(jira_id, auth):
    # Construct the API URL for the specific Jira issue
    url = f"{jira_url}/rest/api/2/issue/{jira_id}"
    
    # Make the API request
    response = requests.get(url, auth=auth)
    
    # Check if the request was successful
    if response.status_code == 200:
        # Parse the JSON response to get the summary
        issue_data = response.json()
        return issue_data['fields']['summary']
    else:
        return None

def generate_output(data,auth,file_path):
    with open(file_path, 'a') as file:
        # Iterate over the values in the first column
        for index, row in data.iterrows():
            jira_id = row[0]
            
            if isinstance(jira_id,float):
                continue
            
            column_next_data = row[1]
            summary = fetch_jira_summary(jira_id,auth)
                
            if summary:
                    line = f"[{jira_id}]({jira_url}/browse/{jira_id}) - {summary}  ~  **{column_next_data}hrs**  "
                    file.write(line + '\n')
            else:
                    file.write(f"[{jira_id}] ~ **{column_next_data}hrs**  \n")
        
        file.write('\n\n')  # Add two new lines before appending new data
                
def authenticate_jira(username, password):
    # Make a test request to authenticate the credentials
    url = f"{jira_url}/rest/api/2/myself"
    response = requests.get(url, auth=HTTPBasicAuth(username, password))
    
    # Check if authentication is successful
    if response.status_code == 200:
        print("Authentication Successful .... Arre bhai tum to sher ho !")
        return True
    else:
        print("Authentication failed. Please check your username or password. Try Again !")
        return False

def get_unique_file_name(file_path):
    base, extension = os.path.splitext(file_path)
    counter = 1
    new_file_path = file_path
    
    # Generate a new file name if the file already exists
    while os.path.exists(new_file_path):
        new_file_path = f"{base}({counter}){extension}"
        counter += 1
        
    return new_file_path   

def loading_animation():
    animation = "|/-\\"
    idx = 0
    while not stop_animation:
        sys.stdout.write(f"\rGenerating output... {animation[idx % len(animation)]}")
        sys.stdout.flush()
        idx += 1
        time.sleep(0.1) 

if __name__ == "__main__":
    # User input for file path and parameters
    print("Ar ke haal he mittar !!")
    
    auth = HTTPBasicAuth("vikas.bhandari", "W3lcome&&991155")
    input_file_path = "sample.xlsx"
    data = read_excel_file(input_file_path)
    
    output_file_path = get_unique_file_name("output.md")

        
    while True:
        # column = 'I'
        # start_row = 31
        # end_row = 37
        column = input("Enter the column (for ex. 'A','B'...): ")
        start_row = int(input("Enter the starting row number: "))
        end_row = int(input("Enter the ending row number: "))
        
        # Start the animation in a separate thread
        stop_animation = False
        animation_thread = threading.Thread(target=loading_animation)
        animation_thread.start()
        
        result = read_excel_data(data,column,start_row,end_row)
        
        
        # Save the results to a text file
        generate_output(result,auth,output_file_path)
        
        # Stop the animation
        stop_animation = True
        animation_thread.join()
        print("\n")
    
        print(f"Hurray Successful! Output Generated/Appended in {output_file_path}")
        # Ask the user if they want to generate output with different parameters
        another_run = input("Do you want to append more info to same file with different data? (yes/no): ").strip().lower()
        print("\n\n")
        
        if another_run != 'yes':
            print("Exiting the program. Thank you ! Have a great day :) \n\n")
            time.sleep(2)
            break
