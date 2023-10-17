import pyodbc
import re

tablename = 'data'
serialIndex = 0
codeIndex = 1
imgnumIndex = 5
lastnameIndex = 3
firstnameIndex = 4
timestampIndex = 7
tablepath = r'C:\Users\danys\OneDrive\Documents\TP_Shotlist1_with_timestamp.accdb'

def tableConnect():
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + tablepath + ';')
    cursor = conn.cursor()  
    return cursor  

def runSQL(query):
    cursor = tableConnect()
    cursor.execute(query)
    return cursor.fetchall()  

def getEmpty():
    return runSQL('select * from ' + tablename + ' where פספורט is NULL')

def getAccessData():
    Result = []
    query = 'select * from ' + tablename
    #(786, 800265.0, 'כיתה 14', 'רתם', 'קציר', '4222', None, None)   
    for row in runSQL(query):
        i = 0
        while i < len(row):
            if row[i] == None:
                row[i] = ""
            i += 1
        #print(row)
        print(row[imgnumIndex] + " " + row[firstnameIndex] + " " + row[lastnameIndex])
        Result.append([row[imgnumIndex] , row[firstnameIndex] , row[lastnameIndex],row[timestampIndex]])
    return Result

def ParseImgList():
    imglistfile = open(r'TestData\Summary.txt')
    imgList = []
    data = imglistfile.readlines()
    pattern = re.compile("(\d+)")
    matches = pattern.findall(data[0])
    return matches

 

#################
### main
#################
Accessdata = getAccessData()
print(Accessdata)
imgList = ParseImgList()
print(imgList)
print(getEmpty())
