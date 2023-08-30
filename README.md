Preparation

- Copy all files to your computer, e.g. C:\Export_SAPQueryBuilder
- Execute the file "install_requirements.bat" once on your computer. This script installs the libraries.
- Open the .env file and replace the wildcards with your user name and password.
- Execute the script: > python Export_SAPQueryBuilder.py
- This will create an Excel file SAPBOQueryResult.xlsx in the same folder.
- To modify the SAP query, open the script Export_SAPQueryBuilder.py with an editor and edit the variable sqlst with your select statement. 
