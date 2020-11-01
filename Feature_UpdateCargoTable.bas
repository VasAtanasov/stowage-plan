Attribute VB_Name = "Feature_UpdateCargoTable"
'@Folder "StowagePlan.feature.update_cargo_table"
Option Explicit
Option Private Module

Public Sub UpdateTable()
    Dim rowIndexByColor As Object
    Set rowIndexByColor = GetRowIndexByColorMap(PORTS_LIST_RANGE)
    
    Dim cell        As Range
    Dim Hold        As Long
    Dim colorCode   As Long
    Dim cellRow     As Long
    
    Dim unitsCells As Range
    Dim weightCells As Range
    
    For Hold = 1 To HOLDS
        If SheetAndRangeExists(STOWPLAN_SHEET_NAME, "HOLD" & Hold) Then
            For Each cell In HOLD_RANGE(Hold)
                colorCode = CLng(cell.Interior.Color)
                If cell.Interior.colorIndex <> xlColorIndexNone And rowIndexByColor.Exists(colorCode) And WorksheetFunction.IsNumber(cell.Value2) = True Then
                    cellRow = rowIndexByColor.Item(colorCode)
                    If Int(cell.Value2) = cell.Value2 Then
                        If unitsCells Is Nothing Then
                            Set unitsCells = cell
                        Else
                            Set unitsCells = Union(unitsCells, cell)
                        End If
                        PORT_TOTAL_UNITS_FOR_HOLD_CELL(Hold, cellRow).Value2 = PORT_TOTAL_UNITS_FOR_HOLD_CELL(Hold, cellRow).Value2 + cell.Value2
                        PORT_TOTAL_UNITS_CELL(cellRow).Value2 = PORT_TOTAL_UNITS_CELL(cellRow).Value2 + cell.Value2
                    Else
                        If weightCells Is Nothing Then
                            Set weightCells = cell
                        Else
                            Set weightCells = Union(weightCells, cell)
                        End If
                        PORT_TOTAL_WEIGHTS_FOR_HOLD_CELL(Hold, cellRow).Value2 = PORT_TOTAL_WEIGHTS_FOR_HOLD_CELL(Hold, cellRow).Value2 + cell.Value2
                        PORT_TOTAL_WEIGHTS_CELL(cellRow).Value2 = PORT_TOTAL_WEIGHTS_CELL(cellRow).Value2 + cell.Value2
                    End If
                End If
            Next
        End If
    Next Hold
    
    If Not unitsCells Is Nothing Then
        unitsCells.NumberFormatLocal = UNITS_FORMAT
    End If
    If Not weightCells Is Nothing Then
        weightCells.NumberFormatLocal = WEIGHT_FORMAT
    End If
    
    On Error GoTo NotWellFormed
    
    Dim shp                     As Shape
    Dim textArray()             As String
    Dim portRow                 As Long
    Dim first                   As Long
    Dim last                    As Long
    Dim lengthOfArray           As Long
    
    For Each shp In STOWAGE_PLAN_SHAPES
        If Right$(shp.Name, Len(PACKAGE_TAG)) = PACKAGE_TAG Then
            colorCode = CLng(shp.Fill.ForeColor)
            If rowIndexByColor.Exists(colorCode) Then
                textArray = MultiSplit(Trim$(shp.TextFrame2.TextRange), Chr$(10), " ")
                portRow = rowIndexByColor.Item(colorCode)
                first = LBound(textArray)
                last = UBound(textArray)
                lengthOfArray = last - first + 1
            
                If lengthOfArray < 8 Then
                    Err.Raise ERROR_INVALID_DATA, "UnitsAndWeigths", _
                              "There seems to be a problem with the format of package text boxes for " & STOWAGE_PLAN_SHEEET.Cells.Item(portRow, 2).Value2 & _
                              " with value: " & vbNewLine & shp.TextFrame2.TextRange.Text
                End If
                
                If Not IsNumeric(textArray(3)) Or Not IsNumeric(textArray(6)) Then
                    Err.Raise ERROR_INVALID_DATA, "UnitsAndWeigths", _
                              "Expected values are not numbers. Please check packages for " & STOWAGE_PLAN_SHEEET.Cells.Item(portRow, 2).Value2 & _
                              " with value: " & vbNewLine & shp.TextFrame2.TextRange.Text
                End If
                
                PORT_TOTAL_PACKAGES_COUNT_CELL(portRow).Value2 = PORT_TOTAL_PACKAGES_COUNT_CELL(portRow).Value2 + CInt(textArray(3))
                PORT_TOTAL_PACKAGES_WEIGHT_CELL(portRow).Value2 = PORT_TOTAL_PACKAGES_WEIGHT_CELL(portRow).Value2 + CDbl(textArray(6))
                
IfNotWellFormed:

            End If
        End If
    Next shp
    On Error GoTo 0
    
    PopulateLoadingTableSummary
    
    Exit Sub
    
NotWellFormed:
    MsgBox Err.Description, vbCritical
    Err.Clear
    Resume IfNotWellFormed
End Sub

Private Sub PopulateLoadingTableSummary()
    If Not WorksheetExists(HATCH_SUMMARY_SHEET_NAME) Then
        Exit Sub
    End If
    
    Dim hatchSummary    As Variant
    Dim tableSummary    As Variant
    hatchSummary = HATCH_SUMMARY_TABLE_RANGE.Value2
    tableSummary = LOADING_SUMMARY_RANGE.Value2
    
    Dim row         As Long
    Dim colSP       As Long
    Dim colHS       As Long
    Dim i           As Long
    Dim totalCount  As Long
    Dim sum         As Long
    
    For row = LBound(hatchSummary, 1) To UBound(hatchSummary, 1)
        If hatchSummary(row, 1) <> vbNullString Then
            colSP = 1 + (row - 1) * 3
            For i = LBound(tableSummary, 1) To UBound(tableSummary, 1)
                colHS = 7 + (i - 1) * 4
                sum = hatchSummary(row, colHS) + hatchSummary(row, colHS + 2)
                If sum <> 0 Then
                    tableSummary(i, colSP) = sum
                    totalCount = totalCount + sum
                End If
            Next i
        End If
    Next row
    
    Dim isTotalEqual As Boolean
    
    isTotalEqual = STOWAGE_PLAN_LOADING_PORTS_TOTAL.Value2 = totalCount
    
    LOADING_SUMMARY_RANGE.Value2 = tableSummary
    
End Sub

Public Sub ResetStowageSummaryTable()
    ResetUnitsAndWeightsHoldTable
    ResetUnitsAndWightsPackagesTable
    ResetLoadingPortsSummaryTable
    ResetTotalUnitsAndWeightsTable
End Sub

Public Sub ResetUnitsAndWeightsHoldTable()
    HOLD_SUMMARY_RANGE.Value2 = vbNullString
End Sub

Public Sub ResetUnitsAndWightsPackagesTable()
    PACKAGES_SUMMARY_RANGE.Value2 = vbNullString
End Sub

Public Sub ResetLoadingPortsSummaryTable()
    LOADING_SUMMARY_RANGE.Value2 = vbNullString
End Sub

Public Sub ResetTotalUnitsAndWeightsTable()
    TOTAL_UNITS_SUMMARY_RANGE.Value2 = vbNullString
End Sub

Public Sub ResetFormat()
    UPPER_DECK_RANGE.NumberFormatLocal = "General"
    LOWER_DECK_RANGE.NumberFormatLocal = "General"
End Sub


