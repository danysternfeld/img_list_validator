import pyodbc



def getAccessData():
    tablepath = r'C:\Users\danys\OneDrive\Documents\TP_Shotlist1_with_timestamp.accdb'
    tablename = 'data'
    serialIndex = 0
    codeIndex = 1
    imgnumIndex = 5
    lastnameIndex = 3
    firstnameIndex = 4
    timestampIndex = 7
    Result = []
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + tablepath + ';')
    cursor = conn.cursor()
    cursor.execute('select * from ' + tablename)
    #(786, 800265.0, 'כיתה 14', 'רתם', 'קציר', '4222', None, None)   
    for row in cursor.fetchall():
        i = 0
        while i < len(row):
            if row[i] == None:
                row[i] = ""
            i += 1
        #print(row)
        print(row[imgnumIndex] + " " + row[firstnameIndex] + " " + row[lastnameIndex])
        Result.append([row[imgnumIndex] , row[firstnameIndex] , row[lastnameIndex],row[timestampIndex]])
    return Result

#################
### main
#################
Accessdata = getAccessData()
print(Accessdata)
