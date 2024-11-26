VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ExportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()
ChDir (CurrentProject.path)
qname = CurrentDb.QueryDefs("exportQ").Name
FileName = "data_out"
On Error GoTo errorH
Kill FileName & ".xlsx"
errorH:
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, qname, FileName, True

Shell "excel.exe " & FileName & ".xlsx", 1

End Sub
