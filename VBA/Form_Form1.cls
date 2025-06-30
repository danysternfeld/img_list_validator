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
'Me.Width = Me.data_subform.Width + 0.2
DoCmd.Maximize
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

Function queryEmpty(sqlstr) As Boolean
   Set RS = CurrentDb.OpenRecordset(sqlstr)
    If (RS.EOF And RS.BOF) Then
    ' empty
    queryEmpty = True
    Else
    ' not empty
    queryEmpty = False
    End If
End Function

Private Sub show_Light()
    strsqlDup = "SELECT First(dataNoHeb.[IMG]) AS [IMG Field], Count(dataNoHeb.[IMG]) AS NumberOfDups FROM dataNoHeb GROUP BY dataNoHeb.[IMG] HAVING (((Count(dataNoHeb.[IMG]))>1));"
    strsqllegal = "SELECT datanoheb.[img] FROM Datanoheb  WHERE (((datanoheb.img) Not Like '????')) OR ((( datanoheb.img) Like '*[!0123456789]*'))"
    Dim dup As Boolean
    Dim illegal As Boolean
    dup = Not queryEmpty(strsqlDup)
    illegal = Not queryEmpty(strsqllegal)
If dup Or illegal Then
        Me.RedLight.Visible = True
        Me.Greenlight.Visible = False
        If dup Then
            Me.dupreg.Visible = True
            Me.dupregLabel.Visible = True
        End If
        If illegal Then
            Me.illegal.Visible = True
            Me.illegalLabel.Visible = True
        End If
Else
        Me.RedLight.Visible = False
        Me.Greenlight.Visible = True
        Me.dupreg.Visible = False
        Me.dupregLabel.Visible = False
        Me.illegal.Visible = False
        Me.illegalLabel.Visible = False
End If
End Sub

Private Sub update_shotCount()
Me.shotCount.Requery
Me.numAdditionals.Requery
Me.additionals.Requery
Me.noShotsByClass.Requery
Me.Rates.Requery
Me.showLast.Requery
Me.dupreg.Requery
Me.illegal.Requery
show_Light
Dim recs As Integer
Dim shots As Integer
Dim Adds As Integer
shots = Me.shotCount.value
Adds = Me.numAdditionals.value
recs = DCount("*", "[Data]")
On Error GoTo errorH
Me.shotCountLabel.Caption = shots & " (" & Int(shots / recs * 100) & "%)  Added: " & Adds & " Remaining: " & recs - shots
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
doTimeToFinish
End Sub

Private Sub doTimeToFinish()
    strSQL = "SELECT Last(getRate.SHootDate) AS LastOfSHootDate, Last(getRate.PersonPerHour) AS LastOfPersonPerHour FROM getRate"
    Set RS = CurrentDb.OpenRecordset(strSQL)
    If Not (RS.EOF And RS.BOF) Then
        RS.MoveFirst
        PPH = RS.Fields(1)
        If Nz(PPH) = 0 Then
            PPH = "0"
        End If
        TTF = Round((DCount("*", "[Data]") - Me.shotCount.value) / PPH, 2)
        If TTF < 1 Then
            Me.LTimeToFinish.Caption = "Time to finish: " & Round(TTF * 60, 0) & " Minutes"
        Else
            Me.LTimeToFinish.Caption = "Time to finish: " & TTF & " Hours"
        End If
    Else
        Me.LTimeToFinish.Caption = "Inf"
    End If
    

End Sub
