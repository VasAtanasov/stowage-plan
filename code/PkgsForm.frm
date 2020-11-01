VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PkgsForm 
   Caption         =   "Packages Form"
   ClientHeight    =   1665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1995
   OleObjectBlob   =   "PkgsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PkgsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "StowagePlan.view.forms"
Option Explicit

Public Event OnAddPkgs()

Private Sub AddPkgsButton_Click()
    RaiseEvent OnAddPkgs
    Unload Me
End Sub

Private Sub PkgsCountTextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With CreateObject("VBScript.RegExp")
        .Pattern = "^[0-9]+$"
        .IgnoreCase = True
        If Not .test(PkgsCountTextBox.value & Chr$(KeyAscii.value)) Then KeyAscii.value = 0
    End With
End Sub

Private Sub PkgsCountTextBox_AfterUpdate()
    PkgsCountTextBox.value = Format$(PkgsCountTextBox, "0")
End Sub

Private Sub PkgsWeightTextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With CreateObject("VBScript.RegExp")
        .Pattern = "^[+]?([0-9]+([.][0-9]*)?|[.][0-9]+)$"
        .IgnoreCase = True
        If Not .test(PkgsWeightTextBox.value & Chr$(KeyAscii.value)) Then KeyAscii.value = 0
    End With
End Sub

Private Sub PkgsWeightTextBox_AfterUpdate()
    PkgsWeightTextBox.value = Format$(PkgsWeightTextBox, "0.0")
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Unload Me
    End If
End Sub
