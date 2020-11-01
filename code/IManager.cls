VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "StowagePlan.api"
Option Explicit

Public Sub InitiateProperties()
End Sub

Public Sub UpdateTableCommand()
End Sub

Public Sub CreatePreArivalPlan()
End Sub

Public Sub ExportDischargingPlanToFile()
End Sub

Public Sub ExportDeparturePlanToFile()
End Sub

Public Sub CreateStowagePlanBackUp()
End Sub

Public Sub MoveStowageTableRow(ByVal offset As Long)
End Sub

Public Sub ToggleTableEmptyRows()
End Sub

Public Sub CreateStowageDirectionShape(ByVal direction As Variant)
End Sub

Public Sub CreateCargoBoxShape(ByVal boxId As String, ByVal boxTag As String)
End Sub
