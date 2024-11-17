from encodings import utf_8
import os
from sys import exception
import sys
import tempfile
from threading import local
from cv2 import filter2D
from numpy import astype
import undetected_chromedriver as uc
#from selenium import webdriver
#from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import csv
import pandas as pd
from bidi.algorithm import get_display
from modules.access import *

def flip(x):
  if(type(x) is list):
    new_list = []
    for i in x : 
      new_list.append(get_display(i))
    return new_list
  return get_display(x)

def make_selenium_data_dir():
    user = os.environ["USERPROFILE"]
    datadir = fr'{user}\selenium_data'
    datadir_created_now = False
    if(not os.path.isdir(datadir)):
        os.mkdir(datadir)
        datadir_created_now = True
    return (datadir,datadir_created_now)


def download_sheet(sheet_ID):

    tmpdir = tempfile.gettempdir()
    (datadir,created_now) = make_selenium_data_dir()
    options = uc.ChromeOptions()
    waittime = 600
    if(not created_now):
        options.add_argument('--headless')
        waittime = 10
    options.add_argument(rf'--user-data-dir={datadir}')
    options.add_experimental_option( "prefs", { "download.default_directory": tmpdir })
    url = f"https://docs.google.com/spreadsheets/d/{sheet_ID}/export?format=csv"
    print(f"downloading {url}...")
    driver = uc.Chrome(options = options,use_subprocess=True)
    driver.get(url)
    try:
        WebDriverWait(driver, waittime).until(EC.title_is("xxx"))
    except:
        print(f"Done.")
        return
    
    
def clean(filename):
    if(os.path.exists(filename)):
        os.remove(filename)

def read_csv(filename):
    with open(filename, newline='',encoding='utf8') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=' ', quotechar='|')
        lines = []
        for row in csvreader:
            lines.append(row[0].split(","))   
        return lines[5:]

def transform_data(lines):
    col_headers = ["DB Code",	"auto number", 	"First Name",	"Last Name",	"Class number",	
                   "InBook",	"InPicture",	"Content",	"Content2",	"Content2",	"Job Title",
                   "Day 2",	"Day 3"]
    ColIndices = {}
    for header,index in zip(col_headers,range(len(col_headers))):
        ColIndices[header] = index
    df = pd.DataFrame(lines[5:])  
    if df.empty:
        return df
    filteredDf = pd.DataFrame(df[[ColIndices["auto number"],ColIndices["Class number"],ColIndices["First Name"],ColIndices["Last Name"],
                        ColIndices["Job Title"]]])
    filteredDf.insert(loc=4, column='פספורט', value=['' for i in range(filteredDf.shape[0])])
    filteredDf.columns = ["קוד", "פרק", "פרטי",	"משפחה",	"פספורט",	"תפקיד"]  
    filteredDf["קוד"]  = filteredDf["קוד"].astype(int) + 900111
    filteredDf["פרק"]  =  "כיתה " + filteredDf["פרק"] 
    return filteredDf

def gen_new_excel(lines):
    filename = 'data.xlsx'
    try:
        lines.to_excel(filename, sheet_name='sheet1', index=False)
        currentDir = os.getcwd()
        print(flip(fr"Created {currentDir}\{filename}"))
    except PermissionError:
        print("Error: Failed to write data.xlsx. If opened in excel please close it.")
        input()
        sys.exit() 



TABLE = ""


if __name__ == "__main__":
    needCleanup = False
    localfile = glob.glob(r".\*.csv")
    if len(localfile) > 1 :
        print(f"Error: More than 1 csv file foud in {os.getcwd()} : " + ','.join(localfile))
        input("press any key...")
        sys.exit()
    if len(localfile) == 1 :
        filename = localfile[0]
    else:
        TABLE = getTablePath()
        school,lrcat,sheet_id,sheet_name = getAccessMetaData(TABLE)    
        print("Sheet name: ", flip(sheet_name))
        #sheet_name = u"יסודי רבין פרדס חנה - 3127280 - 22497 - קובץ צלם"
        #sheet_id = "1GjIfbHUEdywOmaCLYQqSBEN0gqRLKqnj-cPc6KGuAgk"
        #sheet_id = "19Bi24J01aKROwjhoYbK4Zgn3E5h2sa41wfHiGCBO8FM"
        #sheet_name = "אורט אמירים בית שאן - 2401270 - 22864 - קובץ צלם"
        basefilename = f"{sheet_name} - Student_items.csv"
        filename = fr"{tempfile.gettempdir()}\{basefilename}" 
        clean(filename)
        needCleanup = True
        download_sheet(sheet_id)  
        if(not os.path.exists(filename)):
            raise Exception(f"Could not download file {filename}")
    
    lines = read_csv(filename)
    data = transform_data(lines)
    if not data.empty:
        gen_new_excel(data)
    if needCleanup:
        clean(filename)
    





