VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandPreArrivalPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "StowagePlan.feature.pre_arrival_plan"
Option Explicit

Implements ICommand

Private m_manager As IManager
Private m_control As IRibbonControl

Private Sub ICommand_InitiateProperties(ByRef manager As IManager, ByVal control As IRibbonControl)
    Set m_manager = manager
    Set m_control = control
End Sub

Private Sub ICommand_Execute(ParamArray args() As Variant)
    m_manager.CreatePreArivalPlan
End Sub

