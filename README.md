#Query Builder for SAP BO4 

Description

The SAP BO Query Builder is a tool used within the SAP Business Objects environment. Its main task is to retrieve information from the SAP BO repository using SQL queries. 

The Query Builder allows you to retrieve meta information about many SAP BO objects such as reports, schedules, user information, etc.
The information cannot be exported and is only available in HTML format, which makes further analysis very difficult. The attributes and values are displayed in separate containers for each object. 
Some attributes also contain substructures with further attributes and values.

For this reason, a Python tool was developed to export and evaluate the result set in the SAP BO Query Builder. The Python tool Export_SAPQueryBuilder is able to log on to the SAP BO Query Builder as a user, start a predefined query and save the result as an Excel file. The data is transposed, i.e. the attributes are written in columns and the values in rows.

The following document contains instructions for installing and using the Python tool Export_SAPQueryBuilder


Prerequisite

- SAP BO4 Admin account
- Python 3.9 or higher on the machine
- Package: Selenium and Webdriver, BeautifulSoup4, Python Decoupler, Openpyxl (for more information, see Installation)


Preparation

- Copy all files to your computer, e.g. C:\Export_SAPQueryBuilder
- Execute the file "install_requirements.bat" once on your computer (Win). This script installs the libraries.
- Open the .env file and replace the wildcards with your user name and password.
- Execute the script: > python main.py
- This will create an Excel file SAPBOQueryResult.xlsx in the same folder.
- To modify the SAP query, open the script Export_SAPQueryBuilder.py with an editor and edit the variable sqlst with your select statement. 


Hint

- This tool should also handle the Query Builder for SAP BO 3. The main difference between BO3 and BO4 is in the naming of the HTML tag elements. If you are using BO 3, please feel free to modify the configuration variables to suit your needs.
