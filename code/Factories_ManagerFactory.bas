Attribute VB_Name = "Factories_ManagerFactory"
'@Folder "StowagePlan.factories"
Option Explicit

Public Function CreateManager() As IManager
    Set CreateManager = New ManagerImpl
    'OnEnd
    'CreateManager.InitiateProperties
    'OnEnd
End Function

