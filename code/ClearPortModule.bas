Attribute VB_Name = "ClearPortModule"
'@Folder "StowagePlan"
Option Explicit

Public Sub ClearSelectedPorts(control As IRibbonControl)
    With ActiveSheet
        If .Name <> STOWPLAN_SHEET_NAME Then
            Exit Sub
        End If
    End With
    ClearPort
End Sub

Public Sub CleareSelectionAction(control As IRibbonControl)
    With ActiveSheet
        If .Name <> STOWPLAN_SHEET_NAME Then
            Exit Sub
        End If
    End With
    CleareSelection
End Sub

Private Sub ClearPort()
    Dim selectedPorts As Range
    Set selectedPorts = Selection
    
    If Not ProperSubSet(PORTS_LIST_RANGE, selectedPorts) Then
        MsgBox "Please select port(s) from the ports list"
        Exit Sub
    End If

    Dim porstByColorIndex As Object
    Set porstByColorIndex = GetDischargingPortsMap(PORTS_LIST_RANGE)
    
    Dim porstByColorIndexForUnload As Object
    Set porstByColorIndexForUnload = GetDischargingPortsMap(selectedPorts)
    
    Dim colorIndexByColorCode As Object
    Set colorIndexByColorCode = GetStowageTableColorsMap(PORTS_LIST_RANGE)
    
    If porstByColorIndexForUnload.Count = 0 Then
        MsgBox "Please select at least one port"
        Exit Sub
    End If
    
    If MsgBox("All units and weights for selected ports will be discarded. Do you wish to continue?", vbYesNo, "Clear ports") = vbNo Then
        Exit Sub
    End If
    
    OnStart
    BackUpStowagePlan

    Dim cell As Range
    Dim Hold As Integer
    Dim colorIndex As Integer
    Dim holdRange As Range
    Dim mCells As Range
    
    Dim key As Variant
    For Each key In porstByColorIndexForUnload.Keys
        CleareLoadedUnitsSectionForPort porstByColorIndexForUnload(key)
        Dim shp As Shape
        Dim shpRng As Range
        Dim colorCode As Long
        Dim portRow As Integer
        For Each shp In Sheets(STOWPLAN_SHEET_NAME).Shapes
            If Left$(shp.Name, Len(STOW_DORECTION_TAG)) = STOW_DORECTION_TAG Then
                Set shpRng = Range(shp.TopLeftCell.Address, Range(shp.BottomRightCell.Address))
                If shpRng.Interior.colorIndex <> xlNone And shpRng.Interior.colorIndex = key Then
                    shp.Delete
                End If
            ElseIf Right$(shp.Name, Len(PACKAGE_TAG)) = PACKAGE_TAG Or Right$(shp.Name, Len(INFO_BOX_TAG)) = INFO_BOX_TAG Then
                colorCode = CLng(shp.Fill.ForeColor)
                If colorIndexByColorCode.Exists(colorCode) And colorIndexByColorCode(colorCode) = key Then
                    shp.Delete
                End If
            End If
        Next shp
    Next key
    
    For Hold = 1 To HOLDS
        Set mCells = Nothing
        If SheetAndRangeExists(STOWPLAN_SHEET_NAME, "HOLD" & Hold) Then
            Set holdRange = Sheets(STOWPLAN_SHEET_NAME).Range("HOLD" & Hold)
            For Each cell In holdRange
                colorIndex = cell.Interior.colorIndex
                If colorIndex <> xlColorIndexNone And porstByColorIndexForUnload.Exists(colorIndex) Then
                    If mCells Is Nothing Then
                        Set mCells = cell
                    Else
                        Set mCells = Union(mCells, cell)
                    End If
                    If cell.MergeCells Then
                        cell.MergeArea.UnMerge
                    End If
                End If
            Next
            If Not mCells Is Nothing Then
                mCells.Value2 = vbNullString
                mCells.Interior.colorIndex = xlColorIndexNone
            End If
        End If
    Next Hold
    
    selectedPorts.Value2 = vbNullString
    selectedPorts.offset(0, 1).Value2 = vbNullString
    selectedPorts.EntireRow.Hidden = True
    
    UnitsAndWeigths
    
    OnEnd
End Sub

Private Sub CleareSelection()
    Dim cell As Range
    Dim colorIndex As Integer
    Dim selectedCells As Range
    Dim mCells As Range
    
    If TypeName(Selection) <> "Range" Then
        Exit Sub
    End If
    
    If MsgBox("Selected data will be discarded. Do you wish to continue?", vbYesNo, "Clear data") = vbNo Then
        Exit Sub
    End If
    
    OnStart
    BackUpStowagePlan
    Set selectedCells = Selection
    For Each cell In selectedCells
        colorIndex = cell.Interior.colorIndex
        If colorIndex <> xlColorIndexNone Then
            If mCells Is Nothing Then
                Set mCells = cell
            Else
                Set mCells = Union(mCells, cell)
            End If
            If cell.MergeCells Then
                cell.MergeArea.UnMerge
            End If
        End If
    Next
    If Not mCells Is Nothing Then
        mCells.Value2 = vbNullString
        mCells.Interior.colorIndex = xlColorIndexNone
        UnitsAndWeigths
    End If
    OnEnd
End Sub

Private Sub CleareLoadedUnitsSectionForPort(portCellRow As Long)
    'To add packages clear
    Dim cell As Range
    Set cell = Union(Range("B" & portCellRow & ":" & "AT" & portCellRow), Range("CS" & portCellRow & ":" & "DB" & portCellRow))
    cell.Value2 = ""
End Sub

