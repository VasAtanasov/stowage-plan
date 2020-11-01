Attribute VB_Name = "Global_Tools_Optimization"
'@Folder "StowagePlan.utils"
Option Private Module
Option Explicit

Public Sub OptimizeVBA(ByVal isOn As Boolean)
    With Application
        .DisplayAlerts = Not (isOn)
        .ScreenUpdating = Not (isOn)
        .EnableEvents = Not (isOn)
        '.DisplayStatusBar = Not (isOn)
        .Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
    End With
End Sub

Public Sub Formulas_Off()
    Application.Calculation = xlManual
End Sub

Public Sub Formulas_On()
    Application.Calculation = xlAutomatic
End Sub

Public Sub ScreenUpdating_Off()
    Application.ScreenUpdating = False
End Sub

Public Sub ScreenUpdating_On()
    Application.ScreenUpdating = True
End Sub

Public Sub OnStart()
    OptimizeVBA True
End Sub

Public Sub OnEnd()
    OptimizeVBA False
End Sub

