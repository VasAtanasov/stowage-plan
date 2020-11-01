VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CargoForm 
   Caption         =   "Cargo Details Form"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3765
   OleObjectBlob   =   "CargoForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CargoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "StowagePlan.view.forms"
Option Explicit

Private Sub CommandButton1_Click()
   
    'If Trim(TextBox1.Value) = "" Then
    '    MsgBox "Destination port cannot be empty"
    '    Exit Sub
    'End If
    
    'If Trim(TextBox2.Value) = "" Then
    '    MsgBox "Loading port cannot be empty"
    '    Exit Sub
    'End If
    
    
    AddStaticCargoShape TextBox1.value, TextBox2.value, ComboBox1.value
    Unload Me
End Sub


Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With CreateObject("VBScript.RegExp")
        .Pattern = "^[A-Z]{0,5}$"
        .IgnoreCase = True
        If Not .test(TextBox1.value & Chr$(KeyAscii)) Then KeyAscii = 0
    End With
End Sub

Private Sub TextBox1_AfterUpdate()
    TextBox1.value = UCase$(TextBox1.value)
End Sub

Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With CreateObject("VBScript.RegExp")
        .Pattern = "^[A-Z]{0,5}$"
        .IgnoreCase = True
        If Not .test(TextBox2.value & Chr$(KeyAscii)) Then KeyAscii = 0
    End With
End Sub

Private Sub TextBox2_AfterUpdate()
    TextBox2.value = UCase$(TextBox2.value)
End Sub

Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With CreateObject("VBScript.RegExp")
        .Pattern = "^[0-9]+$"
        .IgnoreCase = True
        If Not .test(TextBox3.value & Chr$(KeyAscii)) Then KeyAscii = 0
    End With
End Sub

Private Sub TextBox3_AfterUpdate()
    TextBox3 = Format$(TextBox3, "0")
End Sub

Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    With CreateObject("VBScript.RegExp")
        .Pattern = "^[+]?([0-9]+([.][0-9]*)?|[.][0-9]+)$"
        .IgnoreCase = True
        If Not .test(TextBox4.value & Chr$(KeyAscii)) Then KeyAscii = 0
    End With
End Sub

Private Sub TextBox4_AfterUpdate()
    TextBox4 = Format$(TextBox4, "0.000")
End Sub

Private Sub UserForm_Initialize()
    With ComboBox1
        .AddItem PACKING_UNITS
        .AddItem PACKING_PKGS
    End With
End Sub

