Attribute VB_Name = "Feature_Table_MoveRows"
'@Folder "StowagePlan.feature.table"
Option Explicit
Option Private Module

Public Sub MoveRow(ByVal lowerBound As Long, ByVal upperBound As Long, ByVal offset As Long)
    Dim temp As Variant
    Dim tempColor As Long
    With ActiveCell
        If .row + offset < lowerBound Or .row + offset > upperBound Then
            Exit Sub
        End If
        OnStart
        temp = STOWAGE_PLAN_CARGO_TABLE_ROW(.row).Value2
        tempColor = STOWAGE_PLAN_CARGO_TABLE_ROW(.row).Interior.color
        
        STOWAGE_PLAN_CARGO_TABLE_ROW(.row).Value2 = STOWAGE_PLAN_CARGO_TABLE_ROW(.row + offset).Value2
        STOWAGE_PLAN_CARGO_TABLE_ROW(.row).Interior.color = STOWAGE_PLAN_CARGO_TABLE_ROW(.row + offset).Interior.color
        
        STOWAGE_PLAN_CARGO_TABLE_ROW(.row + offset).Value2 = temp
        STOWAGE_PLAN_CARGO_TABLE_ROW(.row + offset).Interior.color = tempColor
        
        ShiftHatchTable (.row - 9), offset
    
        STOWAGE_PLAN_CARGO_TABLE_ROW(.row + offset).Select
        OnEnd
    End With
End Sub

Private Sub ShiftHatchTable(ByVal portIndex As Long, ByVal offset As Long)
    
    Dim leftCol             As Long
    Dim rightCol            As Long
    Dim offsetLeftCol       As Long
    Dim offsetRithgtCol     As Long
    
    leftCol = 8 + portIndex * 4
    rightCol = 8 + portIndex * 4 + 3
    
    offsetLeftCol = leftCol + (offset * 4)
    offsetRithgtCol = rightCol + (offset * 4)
    
    Dim temp As Variant
    temp = HATCH_TABLE_SECTION(leftCol, rightCol).Value2
    HATCH_TABLE_SECTION(leftCol, rightCol).Value2 = HATCH_TABLE_SECTION(offsetLeftCol, offsetRithgtCol).Value2
    HATCH_TABLE_SECTION(offsetLeftCol, offsetRithgtCol).Value2 = temp
End Sub

Private Function HATCH_TABLE_SECTION(ByVal leftCol As Long, ByVal rightCol As Long) As Range
    Set HATCH_TABLE_SECTION = HATCH_SUMMARY_SHEET.Range( _
                              HATCH_SUMMARY_SHEET.Cells.Item(LOADING_SUMMARY_TOP_ROW, leftCol).Address, _
                              HATCH_SUMMARY_SHEET.Cells.Item(LOADING_SUMMARY_BOTTOM_ROW, rightCol).Address)
End Function




