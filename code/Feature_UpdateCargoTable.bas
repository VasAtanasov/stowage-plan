Attribute VB_Name = "Feature_UpdateCargoTable"
'@Folder "StowagePlan.feature.update_cargo_table"
Option Explicit
Option Private Module

Public Sub UpdateTable(ByVal dischPortsData As Scripting.Dictionary)
    Dim cargoTableValues    As Variant:     cargoTableValues = CARGO_SUMMARY_TABLE_RANGE.Value2
    Dim row                 As Long
    Dim color               As Long
    Dim currentPort         As Port
    Dim portNameCell        As Range
    
    Dim colorKey            As Variant
    For Each colorKey In dischPortsData.Keys
        Set currentPort = dischPortsData.Item(colorKey)
        PopulateHoldSummaryRange currentPort
        PopulateTotalUnitsSummaryRange currentPort
        PopulateTotalPkgsSummaryRange currentPort
    Next colorKey
    
    PopulateLoadingTableSummary
End Sub

Private Sub PopulateHoldSummaryRange(ByVal currentPort As Port)
    Dim holdsTableValues    As Variant:     holdsTableValues = HOLD_SUMMARY_RANGE.Value2
    Dim holdNumber          As Long
    Dim currentHold         As Hold
    Dim row                 As Long:        row = currentPort.RowNumber - TABLE_TOP_ROW + 1
    Dim countCol            As Long
    Dim weightsCol          As Long
    Dim index               As Long:        index = 1
    
    For holdNumber = HOLDS To 1 Step -1
        countCol = 1 + ((index - 1) * HOLD_COL_SPREAD)
        weightsCol = countCol + (HOLD_COL_SPREAD / 2)
        Set currentHold = currentPort.HoldData.Item(holdNumber)
        holdsTableValues(row, countCol) = IIf(currentHold.Count = 0, vbNullString, currentHold.Count)
        holdsTableValues(row, weightsCol) = IIf(currentHold.Weight = 0, vbNullString, currentHold.Weight)
        index = index + 1
    Next holdNumber
    
    HOLD_SUMMARY_RANGE.Value2 = holdsTableValues
End Sub

Private Sub PopulateTotalUnitsSummaryRange(ByVal currentPort As Port)
    Dim totalUnitsTableValues       As Variant:     totalUnitsTableValues = TOTAL_UNITS_SUMMARY_RANGE.Value2
    Dim row                         As Long:        row = currentPort.RowNumber - TABLE_TOP_ROW + 1
    Dim countCol                    As Long
    Dim weightsCol                  As Long
    
    countCol = 1
    weightsCol = countCol + (HOLD_COL_SPREAD / 2)
    totalUnitsTableValues(row, countCol) = IIf(currentPort.TotalUnits = 0, vbNullString, currentPort.TotalUnits)
    totalUnitsTableValues(row, weightsCol) = IIf(currentPort.TotalUnitsWeight = 0, vbNullString, currentPort.TotalUnitsWeight)
    
    TOTAL_UNITS_SUMMARY_RANGE.Value2 = totalUnitsTableValues
End Sub

Private Sub PopulateTotalPkgsSummaryRange(ByVal currentPort As Port)
    Dim totalPackagesTableValues    As Variant:     totalPackagesTableValues = PACKAGES_SUMMARY_RANGE.Value2
    Dim row                         As Long:        row = currentPort.RowNumber - TABLE_TOP_ROW + 1
    Dim countCol                    As Long
    Dim weightsCol                  As Long
    
    countCol = 1
    weightsCol = countCol + (HOLD_COL_SPREAD / 2)
    totalPackagesTableValues(row, countCol) = IIf(currentPort.TotalPkgs = 0, vbNullString, currentPort.TotalPkgs)
    totalPackagesTableValues(row, weightsCol) = IIf(currentPort.TotalPkgsWeight = 0, vbNullString, currentPort.TotalPkgsWeight)
    
    PACKAGES_SUMMARY_RANGE.Value2 = totalPackagesTableValues
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
