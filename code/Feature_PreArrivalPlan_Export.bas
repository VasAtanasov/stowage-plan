Attribute VB_Name = "Feature_PreArrivalPlan_Export"
'@Folder "StowagePlan.feature.pre_arrival_plan"
Option Explicit
Option Private Module

Public Sub ExportDischargePlan()
    If Not WorksheetExists(DISCHARGE_PLAN_SHEET_NAME) Then
        MsgBox "There is no discharging plan to export. Create one first.", vbInformation
        Exit Sub
    End If
    
    OnStart
    
    Dim newFileName     As String
    Dim voyageNum       As String: voyageNum = DISCHARGE_PLAN_SHEET.Range("BU2").Value2
    Dim arrivalPort     As String: arrivalPort = DISCHARGE_PLAN_SHEET.Range("BU3").Value2
    
    CreateCurrentVoyageFolder (voyageNum)
    
    newFileName = ThisWorkbook.Path & _
                  Application.PathSeparator & _
                  voyageNum & _
                  Application.PathSeparator & _
                  VESSEL_CODE & _
                  "_Discharging Plan_" & _
                  arrivalPort & _
                  "_" & _
                  voyageNum
    
    
    Dim dischargingPlan As Workbook
    Set dischargingPlan = Workbooks.Add
    With dischargingPlan
        .SaveAs newFileName, 51
        
        DISCHARGE_PLAN_SHEET.Copy After:=.Worksheets.Item(.Worksheets.Count)
        .Worksheets.Item(DISCHARGE_PLAN_SHEET_NAME).UsedRange.Value2 = DISCHARGE_PLAN_SHEET.UsedRange.Value2
        
        STOWAG_PLAN_BOOK.Worksheets.Item(DISCHARGE_PLAN_MAIN_DECK_SHEET_NAME).Copy After:=.Worksheets.Item(.Worksheets.Count)
        
        .Worksheets.Item(1).Delete
        .Worksheets.Item(1).Select
        DISCHARGE_PLAN_SHEET.Delete
        STOWAG_PLAN_BOOK.Worksheets.Item(DISCHARGE_PLAN_MAIN_DECK_SHEET_NAME).Delete
        .Save
        .Close
    End With

    STOWAGE_PLAN_SHEEET.Select
    OnEnd
End Sub


