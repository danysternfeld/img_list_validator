from tkinter import NO
import pyodbc
import sys
import glob
import os


def getTablePath():    
    access_files = glob.glob("*.accdb")
    if(len(access_files)==0):
        print("No access file found.")
        sys.exit()
    if(len(access_files)>1):
        print(f"More than 1 access file in this dir {os.getcwd()}- exiting.")
        sys.exit()
    acc = access_files[0]
    return acc



def tableConnect(tablepath):
    try:
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\\' + tablepath + ';')
    except pyodbc.InterfaceError:
        print("Can't connect to data. Is MS-ACCESS installed ?")
        sys.exit()    
    cursor = conn.cursor()  
    return cursor 

def runSQL(query,path):
    cursor = tableConnect(path)
    cursor.execute(query)
    return cursor.fetchall()  

def getAccessMetaData(table_path):
    schoolIndex = 1
    lrcatIndex = 2
    SheetID = 3
    SheetName = 4
    rows = runSQL('select * from metadata',table_path)
    path = ""
    pathlist = []
    # allow multiple metadata rows to support usage on multiple computers
    if(len(rows) > 0):
        for row in rows:
            if(row[lrcatIndex] != None):
                path = os.path.expandvars(row[lrcatIndex])
            else:
                path = os.getcwd()
            pathlist.append(path)
            if os.path.exists(path):
                result = []
                for item in row[1:]:
                    result.append(item)
                if(len(result) < 4):
                    l = len(result)
                    for item in [""]*l:
                        result.append(item)   
                return (result)             
    return [None,None,None,None]
