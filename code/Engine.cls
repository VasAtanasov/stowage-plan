VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "StowagePlan.core"
Option Explicit

Private m_manager As IManager

Private Sub Class_Initialize()
    Set m_manager = Nothing
End Sub

Public Sub InitiateProperties(ByRef manager As IManager)
    Set m_manager = manager
End Sub

Public Sub Run(ByVal control As IRibbonControl)
    If m_manager Is Nothing Then
        Exit Sub
    End If

    Dim command As ICommand
    Set command = CreateCommand(control, m_manager)
    If command Is Nothing Then
        MsgBox "Unknow action", vbCritical
        Exit Sub
    End If
    command.Execute
End Sub

