VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SetDestinationPortForm 
   Caption         =   "Set Destination Port"
   ClientHeight    =   1845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3390
   OleObjectBlob   =   "SetDestinationPortForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SetDestinationPortForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "StowagePlan.view.forms"
Option Explicit

Dim portsList           As Range

Private Sub PortsListDropDown_Change()
    Dim cell As Range
    With PortsListDropDown
        For Each cell In portsList
            If cell.Value2 = .value Then
                .BackColor = cell.Interior.color
                Exit Sub
            End If
        Next cell
        .BackColor = -2147483643
    End With
End Sub

Private Sub SetDestinationButton_Click()
    SetDestinationPort
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Set portsList = PORTS_LIST_RANGE
    Dim cell As Range
    With PortsListDropDown
        .AddItem "NONE"
        For Each cell In portsList
            If cell.Value2 <> vbNullString Then
                .AddItem cell.Value2
            End If
        Next cell
    End With
End Sub

