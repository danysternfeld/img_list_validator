import os
import subprocess
from bidi.algorithm import get_display
import re


def flip(x):
  if(type(x) is list):
    new_list = []
    for i in x : 
      new_list.append(get_display(i))
    return new_list
  return get_display(x)


def checkDiff() -> bool:
  gitOut = str(subprocess.check_output(f"git diff {outfile}"),encoding='utf-8').split('\n')
  ignores = [
    'diff --git',
    'index',
    '--- a',
    '+++ b'
    '@@',
    'Got'
    'Registered images that do not exist',
    '===',
  ]
  for line in gitOut:
    if(line.strip()==''):     
      continue
    Matched = False
    for ignore in ignores:
      if(line.find(ignore) > -1):
        Matched = True
        break  
  if(not Matched):
    return False
  return True

   


if __name__ == '__main__':
    cwd = os.getcwd()
    Totest = r"..\..\img_list_validator.py"
    outfile = 'imgListValidatorOut.txt'
    testcaseRoot = '..\\'
    testcases = ["With_folder","With_LR_DATA","כלי עזר של דני_משירלי","With_folder_TABLE_VERSION_1.0"]
    passnum = 0
    failnum = 0
    for testcase in testcases:
        dir = testcaseRoot+testcase
        os.chdir(dir)
        try:
            os.remove(outfile)
        except:
            # ok if file is not there.
            pass    
        print(f"Running in {flip(dir)}")
        subprocess.run(f"python {Totest} .")
        gitOutLen = len(subprocess.check_output(f"git status -s {outfile}"))
        if(gitOutLen == 0):
            print("Passed")
            passnum += 1
        else:
            if(checkDiff()):
              print("Passed")
              passnum += 1
            else:
              print("Failed")
              failnum += 1
     


                        
    print(f"Passed: {passnum}")
    print(f"failed: {failnum}")
