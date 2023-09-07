import os
from datetime import datetime
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import pandas as pd
from decouple import config

now = datetime.now()
date_time = now.strftime("%Y.%m.%d")
export_date_time = now.strftime("%Y_%m_%d")

_number_of_objects = None

sqlst = ("SELECT Top 100 " 
              " SI_NEXTRUNTIME, SI_ANCESTOR, SI_KIND, SI_OWNER, SI_NAME, SI_CREATION_TIME,SI_LOCAL_FILEPATH, SI_UNIVERSE, SI_AUTHOR, SI_PARENT_FOLDER, SI_ID, SI_WEBI_DOC_PROPERTIES, SI_PROCESSINFO.SI_FULLCLIENTDATAPROVIDERS, "
              " SI_NEXTRUNTIME, SI_UNIVERSE, SI_CHILDREN, SI_UPDATE_TS, SI_RECURRING, SI_SCHEDULEINFO.SI_SCHEDULE_TYPE, SI_SCHEDULEINFO.SI_STARTTIME, SI_SCHEDULEINFO.SI_ENDTIME, SI_SCHEDULEINFO.SI_DESTINATIONS "
              " FROM CI_INFOOBJECTS "
              " WHERE SI_KIND in ('FullClient','WebI','Excel') AND SI_RECURRING = 1"
              " AND SI_NEXTRUNTIME > '%s'" % date_time )

BASE_DIR = os.path.realpath(__file__)

CONFIG = {
        'USERNAME' : config('QB_USERNAME'),
        'PASSWORD' : config('QB_PASSWORD'),
        'EXPORT_PATH' : BASE_DIR ,
        'EXPORT_FILENAME' : 'SAPBOQuery_Export',
        'EXPORT_FILEEXTENSION' : '.xlsx',
        'URL' : 'http://XXXXSAPBO4:8080/AdminTools/querybuilder/logonform.jsp',
        'TITEL' : 'SAP BusinessObjects Business Intelligence platform - Query Builder',
        'SYS' : 'ncvboxip4:6400',
        'LOGON_CLICK' : "//input[@value='Log On']",
        'SEND_CLICK' : "//input[@value='Submit Query']",
        'RESULT' : "http://XXXXSAPBO4:8080/AdminTools/querybuilder/query.jsp",
        'EXIT' : 'Exit'
}


def get_number_of_objects():
    return _number_of_objects

def set_number_of_objects(value):
    global _number_of_objects
    if value > 0:
        _number_of_objects = value


#Normalises the dictionary data. Remove outlines
def prepareData(dict_data):
     list_to_remove_keys = []

     for i in dict_data:
          if len(dict_data[i]) != get_number_of_objects():
               list_to_remove_keys.append(i)
     
     for rem in list_to_remove_keys:
          dict_data.pop(rem)     

     df = pd.DataFrame(dict_data)   
     print(df)
     return df

#Process the source data into a value pair dictionary
def readData(in_data):

     soup = BeautifulSoup(in_data, 'html.parser')
     data_dict = {}

     for rows in soup.find_all('tr'):
          cell = rows.find_all('td')
          
          column = rows.find_all('td')[0].get_text()

          #Get from results page the number of items found
          if  "Number of InfoObject" in column:
               number_of_objects = column.split(':')
               set_number_of_objects(int(number_of_objects[1]))

          if len(cell) == 1:
              cell = [column,'NA']
          else:
               cell = [column, rows.find_all('td')[1].get_text()]

          if cell[0] in data_dict:
               data_dict[cell[0]].append(cell[1])
          else:
               data_dict[cell[0]] = [cell[1]]
     return data_dict

#Logon to SAP BO Query Builder to retrieve result set
def getSrouceDataFromQBO(sqlst):
     driver = webdriver.Chrome()
     
     driver.get(CONFIG['URL'])
     assert CONFIG['TITEL'] in driver.title

     elem = driver.find_element(By.NAME, 'aps')
     elem.clear()
     elem.send_keys(CONFIG['SYS'])

     elem = driver.find_element(By.NAME,"usr")
     elem.clear()
     elem.send_keys(CONFIG['USERNAME']) #Enters uername

     elem = driver.find_element(By.NAME, "pwd") #Enter password
     elem.clear()
     elem.send_keys(CONFIG['PASSWORD'])
     elem = driver.find_element(By.XPATH, CONFIG['LOGON_CLICK']).click();

     elem = driver.find_element(By.NAME,"sqlStmt")
     elem.clear()
     elem.send_keys(sqlst)
     elem = driver.find_element(By.XPATH, CONFIG['SEND_CLICK']).click();

     driver.get(CONFIG['RESULT'])
     time.sleep(2)
     sroucecode = driver.page_source

     driver.back()
     #Logout from Site
     elem = driver.find_element(By.PARTIAL_LINK_TEXT,CONFIG['EXIT']).click();

     driver.close()
     return sroucecode
     

def main():
     global sqlst
     run_app = True
     i = 0

     while run_app:
          fileexport = f"{CONFIG['EXPORT_FILENAME']}_{i}_{export_date_time}{CONFIG['EXPORT_FILEEXTENSION']}"
          query_input = input("Enter your query, else leave empty default query will be executed: ")
          
          if query_input:
               print(query_input)
               sqlst = query_input
          else:
               print('Default query executed')
          
          prepareData(readData(getSrouceDataFromQBO(sqlst))).to_excel(str(fileexport)) 
          print(f'File created: {fileexport}')
          
          end_input = input("Continue (y/n)?: ")

          if end_input.lower() == 'n':
               run_app = False
          else:
               i += 1
          
if __name__ == "__main__":
    main()


