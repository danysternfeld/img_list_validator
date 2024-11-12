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

import re
import sys
import os
import glob
import traceback
from lrtools.lrcat import LRCatDB, LRCatException
from lrtools.lrselectgeneric import LRSelectException
from modules.access import *

TABLE = ""

tablename = 'data'
serialIndex = 0
codeIndex = 1
imgnumIndex = 5
lastnameIndex = 3
firstnameIndex = 4
timestampIndex = 7

def getEmpty():
    return runSQL('select * from ' + tablename + ' where פספורט is NULL',TABLE)

def getAccessData():
    Result = []
    query = 'select * from ' + tablename
    #(786, 800265.0, 'כיתה 14', 'רתם', 'קציר', '4222', None, None)   
    for row in runSQL(query,TABLE):
        i = 0
        while i < len(row):
            if row[i] == None:
                row[i] = ""
            i += 1
        #print(row)
        print(row[imgnumIndex] + " " + row[firstnameIndex] + " " + row[lastnameIndex])
        Result.append([row[imgnumIndex] , row[firstnameIndex] , row[lastnameIndex],row[timestampIndex]])
    return Result

def getNonEmpty():
    return runSQL('select * from ' + tablename + ' where פספורט is NOT NULL',TABLE)

def doDND():
    # this supports dropping a folder onto this script
    isDND = False    
    if(len(sys.argv) > 1 ):
        dir = (sys.argv[1])
        os.chdir(dir)
        isDND = True
    return isDND

def printRowHeader():
    print("code\t\timg Number\t\tname")
    print("-------------------------------")

def printRow(row):
    code = row[codeIndex]
    if(code != None):
        code = str(int(code))
    else:
        code = "       "
    print(code + "\t\t" + str(row[imgnumIndex]) + "\t\t" + str(row[firstnameIndex]) + " " + str(row[lastnameIndex]))

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



def getImgListFromLR(school,lrcat):
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
    pattern = re.compile(r"(\d+)$")
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

def isLR(lrcat_field):
    lr_postfix = ".lrcat"
    res = lrcat_field.find(lr_postfix,len(lrcat_field)-len(lr_postfix))
    if( res != -1):
        return True
    return False

def getImgListFromFolder(lrcat_field):
    folder = os.getcwd()
    imglist = []
    if(lrcat_field != ""):
        folder = lrcat_field
    if(not os.path.isdir(folder)):
        print(f"ERROR: No such folder {folder}")
    else:
        files = glob.glob("*.jpg")
        for file in files:
            match = re.findall(r"(\d+)\.jpg",file)
            if(len(match)>0):
                imglist.append(match[0])
    print("Got " + str(len(imglist)) + f" images from folder {folder}")
    if(len(imglist) == 0):
        print(r"/!\ /!\ /!\  !!! ")
        print(r" T   T   T")
        print("WARNING !!!!  NO IMAGES FOUND IN FOLDER!!!!")
        print(r"/!\ /!\ /!\  !!! ")
        print(r" T   T   T") 
    return imglist

#################
### main
#################
if __name__ == "__main__":
    try:
        isDND = doDND()
        TABLE = getTablePath()
        school,lrcat,SheetID,SheetName = getAccessMetaData(TABLE)
        logFile = 'imgListValidatorOut.txt'
        sys.stdout =  open(logFile, 'w', encoding='utf-8')
        empty = getEmpty()
        nonEmpty = getNonEmpty()
        if(lrcat == None):
            # empty lrcat field - use CWD
            imgList = getImgListFromFolder("")
        else:
            if isLR(lrcat):
                imgList = getImgListFromLR(school,lrcat)
            else:
                # non empty and not LR cat: assume folder
                imgList = getImgListFromFolder(lrcat)
        if len(imgList) > 0:
            imgList = RemoveLeadingZerosFromList(imgList)
            RemoveLeadingZerosFromDB(nonEmpty)
            checkImgExists(nonEmpty,imgList)
            printSectSeperator()
            checkImagesNotInDB(nonEmpty,imgList)
            printSectSeperator()

        FindDuplicates(nonEmpty)
        printSectSeperator()
        printEmpty(empty)
        currdir = os.getcwd()
        #sys.stdout.close()
        #os.system(f"notepad {currdir}\\{logFile}")
    except Exception as err:
        print(f"Unexpected ERROR:\n {err=}, {type(err)=}")
        print("error is: "+ traceback.format_exc())
        #input("press any key")
