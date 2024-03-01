# Copyright 2023 by Dany Sternfeld.
# All rights reserved.
#
# This script validates the headshot image database for:
# 1. people with a missing photo
# 2. photos not linked to anyone
# 3. duplicate photos linked to multiple people
# 4. people without any photo
#
# Usage:
# place the image DB in a directory along with the photo list Summary.txt.
# the photo list is created with the lightroom plugin "lightroom transporter"
# and is formatted as a single line like this:
#   ;;;IMGR1152;;;  ;;;IMGR1167;;;  ;;;IMGR4051;;; ...etc
# OR
# Get data directly from lightroom :
# 1. populate the metadata table in access with:
#    a. Name of the school
#    b. Path to the LR catalog
# 2. Place all school images in LR in a collection named the same as in 1a.
#
# Drag and drop the directory onto this script.
# it will create a report in a file named imgListValidatorOut.txt
# if Summary.txt is missing the script will only perform some of the checks.
#################################################################################

from sqlite3 import Timestamp
from cv2 import log
import pyodbc
import re
import sys
import os
import glob
import traceback
from lrtools.lrcat import LRCatDB, LRCatException
from lrtools.lrselectgeneric import LRSelectException



tablename = 'data'
serialIndex = 0
codeIndex = 1
imgnumIndex = 5
lastnameIndex = 3
firstnameIndex = 4
timestampIndex = 7

def getTablePath():
    acc = (glob.glob("*.accdb"))[0]
    return acc

def tableConnect():
    tablepath = getTablePath()
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\\' + tablepath + ';')
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
    matches = []
    file = r'Summary.txt'
    if os.path.exists(file):
        imglistfile = open(file)
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
        dir = (sys.argv[1])
        os.chdir(dir)
        isDND = True
    return isDND

def printRowHeader():
    print("serial\t\timg Number\t\tname")
    print("-------------------------------")

def printRow(row):
    print(str(row[serialIndex]) + "\t\t" + str(row[imgnumIndex]) + "\t\t" + str(row[firstnameIndex]) + " " + str(row[lastnameIndex]))

def checkImgExists(nonEmptyRows,imgList):
    print("Registered images that do not exist:")
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

def getAccessMetaData():
    schoolIndex = 1
    lrcatIndex = 2
    rows = runSQL('select * from metadata')
    # allow multiple metadata rows to support usage on multiple computers
    for row in rows:
        if os.path.exists(row[lrcatIndex]):
            return (row[schoolIndex],row[lrcatIndex])        
    raise Exception("No LR catalogue was found - check access metadata") 

def getImgListFromLR():
    school,lrcat = getAccessMetaData()
    lrdb = LRCatDB(lrcat)

    # select photos
    columns = "name=base"
    criteria = f"collection={school}, rating=>4"
    rows = lrdb.lrphoto.select_generic(columns, criteria).fetchall() # type: ignore
    print("Got " + str(len(rows)) + " images from LR")
    if(len(rows) == 0):
        print(r"/!\ /!\ /!\  !!! ")
        print(r" T   T   T")
        print("WARNING !!!!  NO IMAGES FROM LIGHTROOM !!!!")
        print(r"/!\ /!\ /!\  !!! ")
        print(r" T   T   T")        
        return []
    pattern = re.compile("(\d+)")
    matches = []
    for row in rows:
        matches.append((pattern.findall(row[0]))[0])
    return matches

def RemoveLeadingZerosFromList(Ilist):
    cleanList = []
    for item in Ilist:
        cleanList.append(item.lstrip("0"))
    return cleanList


def RemoveLeadingZerosFromDB(DBRows):
    for item in DBRows:
        item[imgnumIndex] = item[imgnumIndex].lstrip("0")



#################
### main
#################
try:
    isDND = doDND()
    logFile = 'imgListValidatorOut.txt'
    sys.stdout =  open(logFile, 'w', encoding='utf-8')
    imgList = ParseImgList()
    empty = getEmpty()
    nonEmpty = getNonEmpty()
    if len(imgList) == 0:
        imgList = getImgListFromLR()
    if len(imgList) > 0:
        imgList = RemoveLeadingZerosFromList(imgList)
        RemoveLeadingZerosFromDB(nonEmpty)
        checkImgExists(nonEmpty,imgList)
        printSectSeperator()
        checkImagesNotInDB(nonEmpty,imgList)
        printSectSeperator()
    else:
        print("could not open Summary.txt - operation will be limited")

    FindDuplicates(nonEmpty)
    printSectSeperator()
    printEmpty(empty)
    currdir = os.getcwd()
    sys.stdout.close()
    os.system(f"notepad {currdir}\\{logFile}")
except Exception as err:
    print(f"Unexpected ERROR:\n {err=}, {type(err)=}")
    print("error is: "+ traceback.format_exc())
    input("press any key")
