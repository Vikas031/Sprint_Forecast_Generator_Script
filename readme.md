This is a utility helper tool to parse jira tickets from capacity planning excel sheet and helps in preparing ION.WEB forecast quickly.

To build exec : pyinstaller  --onefile forecast_generator.py

**Note.** The tool require two files as input at same directory level as the executable. Also the name of the file should be same as mentioned here - 
- **forecast.xlsx file** - The capacity planning excel sheet you want the tool to parse. *( Tip - Generate a excel file with this name and copy the content of current sprint planning sheet here)*.
5
- **input.txt** The input data which will help generator fetch the information from excel. The content should be in CSV format in which each row denotes a section and as per the guidelines/format given below. 
    
    Ex. input.txt file 

        24.6 => Current sprint info.  
        I,31,37 => Product Backlog item row - cl,sr,er
        D,22,34 => Refinement row - cl,sr,er
        G,23,33 => Issues row - cl,sr,er
        DBX Project Involvement,44  => Other miscellaneous streams, estimate
        Accessiblity,33  => Other miscellaneous streams, estimate

    * hrs = Total Estimate in hours 
    * cl = column denoting the jira tickets for the corresponding section
    * sr = start row of the section containing jira tickets 
    * er = end row 

    Basically cl,sr and er combination given to fetch the jira tickets from the cl and between the sr to er section.

    Note : The first 4 rows are compulsory to be added in file and the rest other miscellaneous section can be added any number for times. If you want to skip any section just pass the first param as !;
    For ex. Suppose we want to skip the issue section : 
        
        24.6 => Current sprint info. 
        135,I,31,37 => Product Backlog item row - hrs,cl,sr,er
        !,D,22,34 => Refinement row - hrs,cl,sr,er 
        12,G,23,33 => Issues row - *This will be skipped*

**Output** : The output will be generated in forecast.md file( will be generate in same directory as executable). 

**For Testing** :
The Repo has sample_input.txt file and sample_planning.xlsx file. Remove sample_ from the file names and test the script with these file. 

**How to use this forecast file :**  

- How to share it via team : Open forecast.md file in preview format and copy content from preview and directly paste it in teams message box. In this way you will get format as in md file here. Now verify and edit the final forecast. Voila you're done !  

