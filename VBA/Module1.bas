Attribute VB_Name = "Module1"
Option Compare Database

' ----------------------------------------------------------------------'
' Return True if the given string value matches the given Regex pattern '
' ----------------------------------------------------------------------'
Public Function RegexMatch(value As Variant, pattern As String) As Boolean
    If IsNull(value) Then Exit Function
    ' Using a static, we avoid re-creating the same regex object for every call '
    Static regex As Object
    ' Initialise the Regex object '
    If regex Is Nothing Then
        Set regex = CreateObject("vbscript.regexp")
        With regex
            .Global = True
            .IgnoreCase = True
            .MultiLine = True
        End With
    End If
    ' Update the regex pattern if it has changed since last time we were called '
    If regex.pattern <> pattern Then regex.pattern = pattern
    ' Test the value against the pattern '
    RegexMatch = regex.test(value)
End Function

Public Function ExtractNumeric(strInput) As Integer
'http://www.utteraccess.com/forum/Creating-Unique-Property-t1964859.html

' Returns the numeric characters within a string in
' sequence in which they are found within the string

Dim strResult As String, strCh As String
Dim intI As Integer
If Not IsNull(strInput) Then
For intI = 1 To Len(strInput)
strCh = Mid(strInput, intI, 1)
Select Case strCh
Case "0" To "9"
strResult = strResult & strCh
Case Else
End Select
Next intI
End If
ExtractNumeric = Val(strResult)

End Function



Sub ExportAllCode()

  For Each C In Application.VBE.VBProjects(1).VBComponents
      Select Case C.Type
          Case vbext_ct_ClassModule, vbext_ct_Document
              Sfx = ".cls"
          Case vbext_ct_MSForm
              Sfx = ".frm"
          Case vbext_ct_StdModule
              Sfx = ".bas"
          Case Else
              Sfx = ""
      End Select
      If Sfx <> "" Then
          C.Export _
              FileName:=CurrentProject.path & "\" & _
              C.Name & Sfx
      End If
  Next C

  End Sub




