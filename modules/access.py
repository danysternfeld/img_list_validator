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
        print("More than 1 access file in this dir - exiting.")
        sys.exit()
    acc = access_files[0]
    return acc

TABLE = getTablePath()

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

def getAccessMetaData():
    schoolIndex = 1
    lrcatIndex = 2
    SheetID = 3
    SheetName = 4
    rows = runSQL('select * from metadata',TABLE)
    # allow multiple metadata rows to support usage on multiple computers
    for row in rows:
        path = os.path.expandvars(row[lrcatIndex])
        if os.path.exists(path):
            result = []
            for item in row[1:]:
                result.append(item)
            if(len(result) < 4):
                l = len(result)
                for item in [""]*l:
                    result.append(item)   
            return (result)        
    raise Exception(f"No LR catalogue was found in {path} - check access metadata") 