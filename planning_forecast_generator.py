import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import csv
import os
import warnings
import getpass
import sys
import time
import threading
import math

# Suppress all warnings
warnings.filterwarnings("ignore")

jira_url = "https://jira.iongroup.com"
output_file_path = ""
auth = ""
stop_animation = True

def read_csv_file(file_path):
    data = []
    with open(file_path, mode='r') as file:
            csv_reader = csv.reader(file, delimiter=',')
            for row in csv_reader:
                data.append(row)
    
    return data

def read_excel_file(file_path):
        df = pd.read_excel(file_path, engine='openpyxl')
        return df


def read_excel_data(df, column_letter, start_row, end_row):
    try:
        # Convert column letter to column index
        column_index = ord(column_letter.upper()) - ord('A')
        if column_index < 0 or column_index >= len(df.columns):
            raise ValueError(f"Invalid column letter '{column_letter}'. Column out of range.")
        
        # Slice the DataFrame to get the specific rows and columns
        excel_data = df.iloc[start_row-1:end_row-1, column_index:column_index+2]
        
        # Check if the row indices are within the bounds of the DataFrame
        if start_row < 1 or end_row > len(df):
            raise ValueError("Row indices are out of range.")

        return excel_data

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

def generate_output(data,auth,file):
        # Iterate over the values in the first column
        for index, row in data.iterrows():
            jira_id = row[0]
            
            if isinstance(jira_id,float):
                continue
            
            column_next_data = row[1]
            summary = fetch_jira_summary(jira_id,auth)
                
            if summary:
                    line = f"- [{jira_id}]({jira_url}/browse/{jira_id}) - {summary}  ~  {column_next_data}hrs  "
                    file.write(line + '\n')
            else:
                    file.write(f"- [{jira_id}] ~ {column_next_data}hrs  \n")
        
        file.write('\n\n')  # Add two new lines before appending new data
                
def authenticate_jira(_username,_password):
    # Make a test request to authenticate the credentials
    url = f"{jira_url}/rest/api/2/myself"
    response = requests.get(url, auth=HTTPBasicAuth(_username, _password))
    
    # Check if authentication is successful
    if response.status_code == 200:
        print("Jira Authentication Successful !\n")
        return True
    else:
        print("Authentication failed. Please check your username or password. Try Again ! \n")
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
        sys.stdout.write(f"\rFetching jira tickets... {animation[idx % len(animation)]}")
        sys.stdout.flush()
        idx += 1
        time.sleep(0.1) 

def add_jira_tickets(row,file,section_name):
    print(f"Extracting jira data from give column info ${row} ...")
     # column = 'I'
     # start_row = 31
    # end_row = 37
    column = row[0]
    start_row = int(row[1])-1
    end_row = int(row[2])
            
    # Start the animation in a separate thread
    global stop_animation 
    stop_animation = False
    animation_thread = threading.Thread(target=loading_animation)
    animation_thread.start()
            
    result = read_excel_data(excel_data,column,start_row,end_row)
            
    # Save the results to a text file
    generate_output(result,auth,file)
            
    # Stop the animation
    stop_animation = True
    animation_thread.join()
    print(f"\n{section_name} section added successfully !\n")
        
def calculate_total_time(row_data):
    column = row_data[0]
    start_row = int(row_data[1])-1
    end_row = int(row_data[2])
    
    data = read_excel_data(excel_data,column,start_row,end_row)
    total = 0
    for index, row in data.iterrows():
        if not math.isnan(row[1]):
            total += int(row[1])
    
    return total
    
        
if __name__ == "__main__":
    # User input for file path and parameters
    print("Hello, Welcome to ION.WEB Sprint Forecast Generator !!\n")
    
    print("Reading input files : planning.xlsx and input.txt ....\n ")
    try:  
        input_file_path="input.txt"
        input_data = read_csv_file(input_file_path)
        forecast_file_path = "planning.xlsx"
        excel_data = None
        
        # Call the function and print the result
        excel_data = read_excel_file(forecast_file_path)    
    except Exception as e :
        print(e)
        time.sleep(2)
        exit()
    
    username = ''
    password='' 
    # credential check kar lete he 
    while True:
        username = input("Please enter your jira username  : ")
        password = getpass.getpass("Please enter you jira password :")
        value = authenticate_jira(username,password)
        if value:
            break
    
    # Setup basic authentication
    auth = HTTPBasicAuth(username, password)
      
    output_file_path = get_unique_file_name("forecast.md")
    
    sprint_no = input_data[0][0]
    
    with open(output_file_path,'w') as file:
         print("Generating Forecast ...\n")
         
         file.write("Hi  \n")
         file.write(f"Please find below forecast for **Sprint {sprint_no}**  \n\n")
         
         
         if(input_data[1][0]!="!"):
             # 1. add product backlog section 
            print("Adding Product Backlog Section !")
            product_backlog_time  = calculate_total_time(input_data[1])
            file.write(f"**Product backlog items (Estimate - {product_backlog_time} Hours) . Following stories will be completed in this sprint:**  \n\n")
            add_jira_tickets(input_data[1],file,"Product Backlog")
         
        #  Adding other sections
         for i in range(4, len(input_data)):
             if(input_data[i][0]=="!"):
                 continue
             file.write(f"**{input_data[i][0]} (Estimate - {input_data[i][1]} Hours)**  \n\n")
             print(f"Added {input_data[i][0]} Estimate  !\n")
             
         # 2. Refinement section 
         if(input_data[2][0] != '!'):
            print("Adding Refinement Section  !")
            refinement_time= calculate_total_time(input_data[2])
            file.write(f"**Refinement Items (Estimate - {refinement_time} Hours):**  \n\n")
            add_jira_tickets(input_data[2],file,"Refinement")
             
         #  3. Issue section 
         if(input_data[3][0] != '!'):
            print("Adding Issues Section !")
            issues_time= calculate_total_time(input_data[3])
            file.write(f"**Issues (Estimate - {issues_time} Hours):**  \n\n")
            add_jira_tickets(input_data[3],file, "Issue")
         
         
         file.write("Thanks")
             
         print("")
         print(f"The forecast has been generated successfully in {output_file_path} file.") 
         print("Exiting the program. Thank you ! Have a great day :) \n\n")
         time.sleep(2)

                