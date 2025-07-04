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

from itertools import count
import re
import sys
import os
import glob
import traceback
from lrtools.lrcat import LRCatDB, LRCatException
from lrtools.lrselectgeneric import LRSelectException
from traitlets import All
from modules.access import *

TABLE = ""

tablename = 'data'
serialIndex = 0
codeIndex = 1
imgnumIndex = 5
lastnameIndex = 3
firstnameIndex = 4
prevImgIndex = 7
timestampIndex = 8


def Print2File(txt=''):
    OUTFILE.writelines(txt+'\n')

def GetTableVersion():
    rows = runSQL('select * from ' + tablename, TABLE)
    for row in rows:
        if len(row) == 8:
            return 1.0
        else:
            return 2.0

def getEmpty():
    return runSQL('select * from ' + tablename + ' where פספורט is NULL',TABLE)

def getNonEmpty():
    return runSQL('select * from ' + tablename + ' where פספורט is NOT NULL',TABLE)

def getRowsWithPrevImage():
    return runSQL('select * from ' + tablename + ' where פספורט is NOT NULL AND Prev_Img is NOT NULL',TABLE)

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

def Print2FileRow(row,index=imgnumIndex,existsButUnrated=False):
    code = row[codeIndex]
    ExistsComment = ""
    if existsButUnrated:
        ExistsComment = "\t\t*** (Image exists but not rated) ***"
    if(code != None):
        code = str(int(code))
    else:
        code = "       "
    Print2File(code + "\t\t" + str(row[index]) + "\t\t" + str(row[firstnameIndex]) + " " + str(row[lastnameIndex]) + ExistsComment)

def checkImgExists(nonEmptyRows,imgList,AllImagesList=[]):
    Print2File("Registered images that do not exist:")
    Print2File("=====================================")
    Print2File()
    PrintRowHeader()
    count = 0
    for row in nonEmptyRows:
        imgNum = row[imgnumIndex]
        ExistsButUnrated = False
        if not imgNum in imgList:
            if imgNum in AllImagesList:
                ExistsButUnrated = True
            Print2FileRow(row,imgnumIndex,ExistsButUnrated)
            count += 1
    Print2File("Total :" + str(count))
    Print2File()
    Print2File()

def OutdatedImages(doubleImages):
    if(len(doubleImages) == 0):
        return
    Print2File("Outdated Images (Newer images registered) :")
    Print2File("-------------------------------")
    Print2File()
    PrintRowHeader()
    count = 0
    for row in doubleImages:
        Print2FileRow(row,prevImgIndex)
        count += 1
    Print2File("Total :" + str(count))
    Print2File()
    Print2File()

def checkImagesNotInDB(nonEmptyRows,imgList,doubleImages):
    Print2File("Images that are not registered:")
    Print2File("-------------------------------")
    Print2File()
    Print2File("image number")
    Print2File("------------")
    registered = []
    doubles = []
    count = 0
    for row in nonEmptyRows:
        registered.append(row[imgnumIndex])
    if(TableVersion == 2.0):
        for row in doubleImages:
            doubles.append(row[prevImgIndex])
    else:
        doubles = []
    for img in imgList:
        if (not img in registered) and (not img in doubles):
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



def getImgListFromLR(school,lrcat,rating=">=1"):
    lrdb = LRCatDB(lrcat)
    # select photos
    columns = "name=base"
    criteria = f"collection={school}, rating={rating}"
    rows = lrdb.lrphoto.select_generic(columns, criteria).fetchall() # type: ignore
    Print2File("Got " + str(len(rows)) + " images from LR with rating " + rating)
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
        TableVersion = GetTableVersion()
        empty = getEmpty()
        nonEmpty = getNonEmpty()
        imgList = []
        AllImagesList = []
        if(TableVersion == 2.0):
            doubleImages = getRowsWithPrevImage()
        else:
            doubleImages = []
        if(lrcat == None):
            # empty lrcat field - use CWD
            imgList = getImgListFromFolder("")
        else:
            if isLR(lrcat):
                lrcat = os.path.expandvars(lrcat)
                imgList = getImgListFromLR(school,lrcat)
                AllImagesList = getImgListFromLR(school,lrcat,rating="=0")
            else:
                # non empty and not LR cat: assume folder
                imgList = getImgListFromFolder(lrcat)
        if len(imgList) > 0:
            imgList = RemoveLeadingZerosFromList(imgList)
            if len(AllImagesList) > 0:
                AllImagesList = RemoveLeadingZerosFromList(AllImagesList)            
            RemoveLeadingZerosFromDB(nonEmpty)  
            checkImgExists(nonEmpty,imgList,AllImagesList)
            Print2FileSectSeperator()
            if(TableVersion == 2.0):
                OutdatedImages(doubleImages)
                Print2FileSectSeperator()
            checkImagesNotInDB(nonEmpty,imgList,doubleImages)
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
