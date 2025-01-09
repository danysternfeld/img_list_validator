VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub additionals_Enter()
update_shotCount
End Sub

Private Sub additionals_Exit(Cancel As Integer)
update_shotCount
End Sub

Private Sub Command35_Click()
Dim strPath As String
strPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
nowYear = Year(Now())
nowDay = Day(Now())
nowMonth = Month(Now())
xclPath = strPath & "\noimages_" & nowDay & "_" & nowMonth & "_" & nowYear & ".xlsx"
DoCmd.TransferSpreadsheet _
    TransferType:=acExport, _
    SpreadsheetType:=acSpreadsheetTypeExcel12Xml, _
    TableName:="NoImages", _
    FileName:=xclPath, _
    HasFieldNames:=True

 Shell "excel.exe " & xclPath, 1
End Sub

Private Sub Command36_Click()
Me.srch.SetFocus
Me.srch.Text = ""
End Sub

Private Sub Command7_Click()
Me.srch.SetFocus
Me.srch.Text = ""
End Sub

Private Sub data_subform_Enter()
update_shotCount
End Sub

Private Sub data_subform_Exit(Cancel As Integer)
update_shotCount
End Sub

Private Sub Form_Load()
Me.Width = Me.data_subform.Width + 0.2
End Sub

Private Sub Form_Open(Cancel As Integer)
Me.Width = Me.data_subform.Width + 0.2
End Sub

Private Sub srch_Change()
Dim srchval As String
Me.ghost.value = Me.srch.Text
'Me.srch.Value = srchval
'SELECT data.чег, *
'FROM Data
'WHERE (((data.чег) Like "*" & [Forms]![Form1].[Form]![srch] & "*"));

Me.data_subform.Requery


End Sub

Private Sub srch_GotFocus()
update_shotCount
End Sub

Private Sub update_shotCount()
Me.shotCount.Requery
Me.numAdditionals.Requery
Me.additionals.Requery
Me.noShotsByClass.Requery
Me.showLast.Requery
Dim recs As Integer
Dim shots As Integer
Dim Adds As Integer
shots = Me.shotCount.value
Adds = Me.numAdditionals.value
recs = DCount("*", "[Data]")
On Error GoTo errorH
Me.shotCountLabel.Caption = shots & " (" & Int(shots / recs * 100) & "%)  Added: " & Adds
doProgress
errorH:
End Sub

Private Sub srch_LostFocus()
update_shotCount
End Sub

Private Sub doProgress()
Dim recs As Integer
Dim shots As Integer
Dim percent As Double
fullW = Me.Back.Width
shots = Me.shotCount.value
recs = DCount("*", "[Data]")
percent = shots / recs
Me.progress.Width = fullW * percent
Me.PCT.Caption = Int(percent * 100) & "%"
End Sub
