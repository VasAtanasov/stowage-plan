VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandMoveRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "StowagePlan.feature.table"
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
            MsgBox "Cargo Table not found.", vbCritical
            Exit Sub
        End If
    End With
    m_manager.MoveStowageTableRow IIf(m_control.Tag = "Up", -1, 1)
End Sub

