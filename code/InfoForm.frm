VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InfoForm 
   Caption         =   "Info Box"
   ClientHeight    =   1785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3075
   OleObjectBlob   =   "InfoForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InfoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "StowagePlan.view.forms"
Option Explicit

Public Event OnAddInfoBox()

Private Sub AddInfoButton_Click()
    RaiseEvent OnAddInfoBox
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Unload Me
    End If
End Sub

