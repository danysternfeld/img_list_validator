VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SaveCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()
vbext_ct_Document = 100
vbext_ct_ClassModule = 2
vbext_ct_StdModule = 1
vbext_ct_MSForm = 3
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
              FileName:=CurrentProject.path & "\VBA\" & _
              C.Name & Sfx
      End If
  Next C
End Sub
