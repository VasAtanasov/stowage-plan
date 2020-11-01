Attribute VB_Name = "App"
'@Folder "StowagePlan"
Option Explicit

Global selectedShape As Object

Public Sub Main(ByVal control As IRibbonControl)
    With control.Context
        If TypeName(control.Context.Application.Selection) <> "Range" Then
            Set selectedShape = .Application.Selection
            .ActiveSheet.[A1].Select
        End If
        
        Dim manager     As IManager:    Set manager = CreateManager
        Dim eng         As Engine:      Set eng = CreateEngine(manager)
        eng.Run control
    End With
End Sub
