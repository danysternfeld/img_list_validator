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
# Get data directly from lightroom :
# 1. populate the metadata table in access with:
#    a. Name of the school
#    b. Path to the LR catalog
# 2. Place all school images in LR in a collection named the same as in 1a.
# OR
# 1. Use an image folder instead of lightroom
# 2. Populate the LRCat field in the access metadata table with either 
#    the path to the folder or leave it blank if the images and access are in
#    the same folder.
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


def Print2File(txt=''):
    OUTFILE.writelines(txt+'\n')

def getEmpty():
    return runSQL('select * from ' + tablename + ' where פספורט is NULL',TABLE)

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

def PrintRowHeader():
    Print2File("code\t\timg Number\t\tname")
    Print2File("-------------------------------")

def Print2FileRow(row):
    code = row[codeIndex]
    if(code != None):
        code = str(int(code))
    else:
        code = "       "
    Print2File(code + "\t\t" + str(row[imgnumIndex]) + "\t\t" + str(row[firstnameIndex]) + " " + str(row[lastnameIndex]))

def checkImgExists(nonEmptyRows,imgList):
    Print2File("Registered images that do not exist:")
    Print2File("=====================================")
    Print2File()
    PrintRowHeader()
    count = 0
    for row in nonEmptyRows:
        imgNum = row[imgnumIndex]
        if not imgNum in imgList:
            Print2FileRow(row)
            count += 1
    Print2File("Total :" + str(count))
    Print2File()
    Print2File()


def checkImagesNotInDB(nonEmptyRows,imgList):
    Print2File("Images that are not registered:")
    Print2File("-------------------------------")
    Print2File()
    Print2File("image number")
    Print2File("------------")
    registered = []
    count = 0
    for row in nonEmptyRows:
        registered.append(row[imgnumIndex])
    for img in imgList:
        if not img in registered:
            Print2File(img)
            count += 1
    Print2File("Total :" + str(count))
    Print2File()
    Print2File()

def Print2FileEmpty(empty):
    Print2File("No registered images (" + str(len(empty))+ "):")
    Print2File("=====================")
    Print2File()
    PrintRowHeader()
    for row in empty:
        Print2FileRow(row)
    Print2File()
    Print2File()

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
    Print2File("Found " + str(dupes) + " duplicates.")
    for rowlist in dupList.values():
        if len(rowlist) > 1 :
            Print2File( str(len(rowlist)) + " instances of " + str(rowlist[0][imgnumIndex]) + ":")
            PrintRowHeader()
            for row in rowlist:
                Print2FileRow(row)
            Print2File()

def Print2FileSectSeperator():
    Print2File("==================================================================================================")
    Print2File("==================================================================================================")



def getImgListFromLR(school,lrcat):
    lrdb = LRCatDB(lrcat)

    # select photos
    columns = "name=base"
    criteria = f"collection={school}, rating=>4"
    rows = lrdb.lrphoto.select_generic(columns, criteria).fetchall() # type: ignore
    Print2File("Got " + str(len(rows)) + " images from LR")
    if(len(rows) == 0):
        Print2File(r"/!\ /!\ /!\  !!! ")
        Print2File(r" T   T   T")
        Print2File("WARNING !!!!  NO IMAGES FROM LIGHTROOM !!!!")
        Print2File(r"/!\ /!\ /!\  !!! ")
        Print2File(r" T   T   T")        
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
        Print2File(f"ERROR: No such folder {folder}")
    else:
        files = glob.glob("*.jpg")
        for file in files:
            match = ""
            match = re.findall(r"(\d{1,4})\.jpg",file)
            if(len(match)>0):
                imglist.append(match[0])
    Print2File("Got " + str(len(imglist)) + f" images from folder {folder}")
    if(len(imglist) == 0):
        Print2File(r"/!\ /!\ /!\  !!! ")
        Print2File(r" T   T   T")
        Print2File("WARNING !!!!  NO IMAGES FOUND IN FOLDER!!!!")
        Print2File(r"/!\ /!\ /!\  !!! ")
        Print2File(r" T   T   T") 
    return imglist

#################
### main
#################
if __name__ == "__main__":
    try:
        isDND = doDND()
        logFile = 'imgListValidatorOut.txt'
        try:
            OUTFILE = open(logFile, 'w', encoding='utf-8')
        except:
            print(f"Failed to open {logFile}")
            input("Press any key...")
            sys.exit()
        TABLE = getTablePath()
        school,lrcat,SheetID,SheetName = getAccessMetaData(TABLE)
        empty = getEmpty()
        nonEmpty = getNonEmpty()
        if(lrcat == None):
            # empty lrcat field - use CWD
            imgList = getImgListFromFolder("")
        else:
            if isLR(lrcat):
                lrcat = os.path.expandvars(lrcat)
                imgList = getImgListFromLR(school,lrcat)
            else:
                # non empty and not LR cat: assume folder
                imgList = getImgListFromFolder(lrcat)
        if len(imgList) > 0:
            imgList = RemoveLeadingZerosFromList(imgList)
            RemoveLeadingZerosFromDB(nonEmpty)
            checkImgExists(nonEmpty,imgList)
            Print2FileSectSeperator()
            checkImagesNotInDB(nonEmpty,imgList)
            Print2FileSectSeperator()

        FindDuplicates(nonEmpty)
        Print2FileSectSeperator()
        Print2FileEmpty(empty)
        currdir = os.getcwd()
        OUTFILE.close()
    except Exception as err:
        Print2File(f"Unexpected ERROR:\n {err=}, {type(err)=}")
        Print2File("error is: "+ traceback.format_exc())
        OUTFILE.close()
