VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandUpdateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "StowagePlan.feature.update_cargo_table"
Option Explicit

Implements ICommand

Private m_manager As IManager
Private m_control As IRibbonControl

Private Sub ICommand_InitiateProperties(ByRef manager As IManager, ByVal control As IRibbonControl)
    Set m_manager = manager
    Set m_control = control
End Sub

Private Sub ICommand_Execute(ParamArray args() As Variant)
    With m_control.Context.ActiveSheet
        If Not .Name = STOWPLAN_SHEET_NAME Then
            MsgBox "Swicht to '" & STOWPLAN_SHEET_NAME & "' sheet to performe calculations.", vbCritical
            Exit Sub
        End If
    End With
    
    OnStart
    m_manager.InitiateProperties
    m_manager.UpdateTableCommand
    OnEnd
End Sub

