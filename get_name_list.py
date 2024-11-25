#  Dany Sternfeld - Oct 2024
# Get data from TP's new site and transform it to the older format.
# Generated an excel file data.xlsx that can be imported to TP_shotlist.accdb
#
# Can either download the data from TP or use a .csv file that was manually downloaded.
# to use, place a manualy downloaded csv in the same dir as the accdb file.
#
####################################################################################
import os
import sys
import csv
import glob
import pandas as pd
from bidi.algorithm import get_display


def flip(x):
  if(type(x) is list):
    new_list = []
    for i in x : 
      new_list.append(get_display(i))
    return new_list
  return get_display(x)

    
def clean(filename):
    if(os.path.exists(filename)):
        os.remove(filename)

def read_csv(filename):
    with open(filename, newline='',encoding='utf8') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=' ', quotechar='|')
        lines = []
        for row in csvreader:
            lines.append(row[0].split(",")) 
        # drop the first 5 lines. they contain junk we don't need.      
        return lines[5:]

def transform_data(lines):
    col_headers = ["DB Code",	"auto number", 	"First Name",	"Last Name",	"Class number",	
                   "InBook",	"InPicture",	"Content",	"Content2",	"Content2",	"Job Title",
                   "Day 2",	"Day 3"]
    ColIndices = {}
    for header,index in zip(col_headers,range(len(col_headers))):
        ColIndices[header] = index
    df = pd.DataFrame(lines)  
    if df.empty:
        return df
    filteredDf = pd.DataFrame(df[[ColIndices["auto number"],ColIndices["Class number"],ColIndices["First Name"],ColIndices["Last Name"],
                        ColIndices["Content"],ColIndices["Job Title"]]])
    filteredDf.columns = ["קוד", "פרק", "פרטי",	"משפחה",	"פספורט",	"תפקיד"] 
    # doing this because Added entries may have empty code values
    val = 0
    for i in  range(filteredDf.shape[0]):
        filteredDf.iloc[i,0] = val+1
        val+=1

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
        input("Press any key...")
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
        print(f"Error: No csv file foud in {os.getcwd()} : " + ','.join(localfile))
        input("press any key...")
        sys.exit()

    
    lines = read_csv(filename)
    data = transform_data(lines)
    if not data.empty:
        gen_new_excel(data)
    if needCleanup:
        clean(filename)
    





