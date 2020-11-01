Attribute VB_Name = "Feature_DeparturePlan_Create"
'@Folder "StowagePlan.feature.departure_plan"
Option Explicit
Option Private Module

Public Sub ExportDeparturePlan()
    OnStart
    
    Dim newFileName As String
    
    CreateCurrentVoyageFolder (CURRENT_VOY)
    
    newFileName = ThisWorkbook.Path & _
                  Application.PathSeparator & _
                  CURRENT_VOY & _
                  Application.PathSeparator & _
                  VESSEL_CODE & _
                  "_Stowage Plan Dep. " & _
                  CURRENT_PORT & _
                  " Voy. " & _
                  CURRENT_VOY & _
                  ".xlsx"
    
    Dim departurePlan As Workbook
    Set departurePlan = Workbooks.Add
    
    With departurePlan
        .SaveAs newFileName, 51
        Dim stowagePlanWorksheet As Worksheet
        For Each stowagePlanWorksheet In STOWAG_PLAN_BOOK.Worksheets
            If stowagePlanWorksheet.Name <> DISCHARGE_PLAN_SHEET_NAME And stowagePlanWorksheet.Name <> DISCHARGE_PLAN_MAIN_DECK_SHEET_NAME Then
                stowagePlanWorksheet.Copy After:=.Worksheets.Item(.Worksheets.Count)
                .Worksheets.Item(stowagePlanWorksheet.Name).UsedRange.Value2 = stowagePlanWorksheet.UsedRange.Value2
            End If
        Next stowagePlanWorksheet
        .Worksheets.Item(1).Delete
        .Worksheets.Item(1).Select
        .Save
        .Close
    End With
    
    OnEnd
End Sub


