VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ImportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Public Sub Import_from_excel()
FileName = CurrentProject.path & ".\data.xlsx"
On Error GoTo ErrorHandler
DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "Data", FileName, True
Exit Sub

ErrorHandler:
    res = MsgBox("Import failed. Check that data.xslx exists.", vbOKOnly, "Error !")
End Sub

Private Sub init_timestamps()
Dim db As Database
Dim qdf As QueryDef
Set db = CurrentDb
sqlstr = "UPDATE data SET data.[timeStamp] = #1/1/1980#;"
Set qdf = db.CreateQueryDef("deleteme", sqlstr)
qdf.Execute dbFailOnError
qdf.Close
CurrentDb.QueryDefs.Delete "deleteme"
End Sub

Private Sub Command0_Click()
If (MsgBox("All Data will be deleted !!!" & (Chr(13) & Chr(10)) & " Are you Sure ?", vbOKCancel, "WARNING !!!!")) = vbCancel Then
    Exit Sub
End If
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM Data"
DoCmd.SetWarnings True
Me.Import_from_excel
init_timestamps
res = MsgBox("Import done.", vbOKOnly, "Success !")

End Sub

Public Function FileExists(ByVal path_ As String) As Boolean
    FileExists = (Len(Dir(path_)) > 0)
End Function

Private Function read_reg()
Dim RegObj, RegKey
RegKey = ""
Set RegObj = CreateObject("WScript.Shell")
On Error GoTo errorH
RegKey = RegObj.RegRead("HKCU\Software\DanySternfeld\img_list_validator\Path")
If RegKey = "" Then
    MsgBox "Could not find img_list_validator - check your installation !"
Else
    Set RegObj = Nothing
End If
read_reg = RegKey
Exit Function
errorH:
    MsgBox "Could not find img_list_validator - check your installation !"
    read_reg = RegKey
End Function



Private Function find_names_downloader() As String
    Dim locations(4) As String
    install_location = read_reg()
    
    locations(0) = install_location & "\get_name_list.exe"
    locations(1) = "C:\Users\danys\OneDrive\Documents\scripts\img_list_validator\get_name_list.py"
    arrlen = UBound(locations) - LBound(locations) + 1
    For i = 0 To arrlen - 1
        If FileExists(locations(i)) Then
            find_names_downloader = locations(i)
            Exit Function
        End If
    Next i
End Function

Sub ShellAndWait(pathFile As String)
    With CreateObject("WScript.Shell")
        .Run pathFile, 1, True
    End With
End Sub


Private Sub Command3_Click()
    If (MsgBox("All Data will be deleted !!!" & (Chr(13) & Chr(10)) & " Are you Sure ?", vbOKCancel, "WARNING !!!!")) = vbCancel Then
        Exit Sub
    End If
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM Data"
    DoCmd.SetWarnings True
    path = find_names_downloader
    If path = "" Then
        res = MsgBox("Error. Could not find img_list_validator. Is it installed ?", vbOKOnly, "Error !")
    End If
    If (InStr(path, ".py") > 0) Then
        path = "python " & path
    End If
    ChDir (CurrentProject.path)
    On Error GoTo errorH
    Me.ShellAndWait (path)
    Me.Import_from_excel
    res = MsgBox("Import done.", vbOKOnly, "Success !")
    Exit Sub
errorH:
    res = MsgBox("Error. Failed to run importer." & (Chr(13) & Chr(10)) & "Path was:    " & path, vbOKOnly, "Error !")
    
End Sub

Private Sub Command5_Click()
    If (MsgBox("This will reset the download browser - " & (Chr(13) & Chr(10)) & " Are you Sure ?", vbOKCancel, "WARNING !!!!")) = vbCancel Then
        Exit Sub
    End If
    path = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%USERPROFILE%\selenium_data")
    On Error GoTo errorH:
    CreateObject("Scripting.FileSystemObject").DeleteFolder path
    If (Not FileExists(path)) Then
        res = MsgBox("Cleared browser data.", vbOKOnly, "Success !")
    Else
        res = MsgBox("Failed to clear browser data!", vbOKOnly, "Failed !")
    End If
errorH:
    
End Sub
