Attribute VB_Name = "Factories_BoxFactory"
'@Folder("StowagePlan.factories")
Option Explicit

Public Function CreateCargoBox(ByVal boxTag As String) As CargoBox
    Set CreateCargoBox = New CargoBox
    CreateCargoBox.InitiateProperties boxTag
End Function

Public Function CreateInfoBox() As InfoBox
    Set CreateInfoBox = New InfoBox
    CreateInfoBox.InitiateProperties
End Function


