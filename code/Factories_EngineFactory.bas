Attribute VB_Name = "Factories_EngineFactory"
'@Folder "StowagePlan.factories"
Option Explicit

Public Function CreateEngine(ByRef manager As IManager) As Engine
    Set CreateEngine = New Engine
    CreateEngine.InitiateProperties manager
End Function

