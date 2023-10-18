import pyodbc
import re
import sys
import os



tablename = 'data'
serialIndex = 0
codeIndex = 1
imgnumIndex = 5
lastnameIndex = 3
firstnameIndex = 4
timestampIndex = 7
tablepath = r'.\TP_Shotlist1_with_timestamp.accdb'

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
    imglistfile = open(r'Summary.txt')
    imgList = []
    data = imglistfile.readlines()
    pattern = re.compile("(\d+)")
    matches = pattern.findall(data[0])
    return matches

def getNonEmpty():
    return runSQL('select * from ' + tablename + ' where פספורט is NOT NULL')

def doDND():
    # this supports dropping a folder onto this script
    isDND = False    
    if(len(sys.argv) > 1 ):
        dir = sys.argv[1]
        os.chdir(sys.argv[1])
        os.chdir(dir)
        isDND = True
    return isDND

def printRowHeader():
    print("serial\t\timg Number\t\tname")
    print("-------------------------------")

def printRow(row):
    print(str(row[serialIndex]) + "\t\t" + str(row[imgnumIndex]) + "\t\t" + row[firstnameIndex] + " " + row[lastnameIndex])

def checkImgExists(nonEmptyRows,imgList):
    print("Registered images that do not exists:")
    print("=====================================")
    print()
    printRowHeader()
    count = 0
    for row in nonEmptyRows:
        imgNum = row[imgnumIndex]
        if not imgNum in imgList:
            printRow(row)
            count += 1
    print("Total :" + str(count))
    print()
    print()


def checkImagesNotInDB(nonEmptyRows,imgList):
    print("Images that are not registered:")
    print("-------------------------------")
    print()
    print("image number")
    print("------------")
    registered = []
    count = 0
    for row in nonEmptyRows:
        registered.append(row[imgnumIndex])
    for img in imgList:
        if not img in registered:
            print(img)
            count += 1
    print("Total :" + str(count))
    print()
    print()

def printEmpty(empty):
    print("No registered images (" + str(len(empty))+ "):")
    print("=====================")
    print()
    printRowHeader()
    for row in empty:
        printRow(row)
    print()
    print()

def FindDuplicates(imagesInDB):
    dupList = dict()
    dupes = 0 
    for row in imagesInDB:
        imgnum = int(row[imgnumIndex])
        if not imgnum in dupList.keys():
            dupList[imgnum] = [row]
        else:
            dupes += 1
            dupList[imgnum].append(row)
    print("Found " + str(dupes) + " duplicates.")
    for rowlist in dupList.values():
        if len(rowlist) > 1 :
            print( str(len(rowlist)) + " instances of " + str(rowlist[0][imgnumIndex]) + ":")
            printRowHeader()
            for row in rowlist:
                printRow(row)
            print()

def printSectSeperator():
    print("==================================================================================================")
    print("==================================================================================================")


#################
### main
#################
isDND = doDND()
sys.stdout =  open('imgListValidatorOut.txt', 'w', encoding='utf-8')
#Accessdata = getAccessData()
#print(Accessdata)
imgList = ParseImgList()
#print(imgList)
empty = getEmpty()
nonEmpty = getNonEmpty()

checkImgExists(nonEmpty,imgList)
printSectSeperator()
checkImagesNotInDB(nonEmpty,imgList)
printSectSeperator()
FindDuplicates(nonEmpty)
printSectSeperator()
printEmpty(empty)

