VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ManagerImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "StowagePlan.core"
Option Explicit
Implements IManager

Private Type TManager
    DischargingPorts      As Scripting.Dictionary
    LoadPortsCodes()      As String
    UnitsCells            As Range
    WeightsCells          As Range
End Type

Private this As TManager

Private Sub IManager_InitiateProperties()
    this.LoadPortsCodes = GetLoadingPortsSequence()
    Set this.DischargingPorts = GetInitialState()
    SetHoldsProperty
    SetPkgsProperty
End Sub

Private Sub IManager_UpdateTableCommand()
    STOWAGE_PLAN_SHEEET.[AL2].Select
    ResetFormat
    ResetStowageSummaryTable
    UpdateTable this.DischargingPorts
     If Not this.UnitsCells Is Nothing Then
        this.UnitsCells.NumberFormatLocal = UNITS_FORMAT
    End If
    If Not this.WeightsCells Is Nothing Then
        this.WeightsCells.NumberFormatLocal = WEIGHT_FORMAT
    End If
    DischargingPortSequence
End Sub

Private Sub IManager_CreatePreArivalPlan()
    CreatePreArrivalPlan
End Sub

Private Sub IManager_ExportDischargingPlanToFile()
    ExportDischargePlan
End Sub

Private Sub IManager_ExportDeparturePlanToFile()
    ExportDeparturePlan
End Sub

Private Sub IManager_CreateStowagePlanBackUp()
    BackUpStowagePlan
End Sub

Private Sub IManager_MoveStowageTableRow(ByVal offset As Long)
    With ActiveSheet
        If .Name = STOWPLAN_SHEET_NAME Then
            MoveRow TABLE_TOP_ROW, TABLE_BOTTOM_ROW, offset
        End If
    End With
End Sub

Private Sub IManager_ToggleTableEmptyRows()
    ToggleEmptyRows
End Sub

Private Sub IManager_CreateStowageDirectionShape(ByVal direction As Variant)
    StowDirectionCreate direction
End Sub

Private Sub IManager_CreateCargoBoxShape(ByVal boxId As String, ByVal boxTag As String)
    Select Case boxId
    Case ADD_RECTANGLE_CARGO_BOX_BUTTON_ID, _
         ADD_CALLOUT_CARGO_BOX_BUTTON_ID
        AddStaticCargoBox boxTag
    Case ADD_INFO_BOX_BUTTON_ID
        AddInfoBox
    End Select
End Sub

Private Function GetInitialState() As Scripting.Dictionary
    Dim cargoTableValues    As Variant:                 cargoTableValues = CARGO_SUMMARY_TABLE_RANGE.Value2
    Dim portsByColor        As Scripting.Dictionary:    Set portsByColor = New Scripting.Dictionary
    Dim row                 As Long
    Dim currentPort         As Port
    
    Dim portNameCell As Range
    For row = LBound(cargoTableValues, 1) To UBound(cargoTableValues, 1)
        Set portNameCell = CARGO_SUMMARY_TABLE_RANGE.Item(row, 1)
        If portNameCell.Value2 <> vbNullString Then
            Set currentPort = New Port
            currentPort.PortName = cargoTableValues(row, 1)
            currentPort.PortCode = cargoTableValues(row, 13)
            currentPort.color = portNameCell.Interior.color
            currentPort.RowNumber = portNameCell.row

            portsByColor.Add currentPort.color, currentPort
        End If
    Next row
    
    Set GetInitialState = portsByColor
End Function

Private Sub SetHoldsProperty()
    Dim cell            As Range
    Dim holdNumber      As Long
    Dim colorCode       As Long
    Dim currentPort     As Port
    Dim currentHold     As Hold
    
    For holdNumber = 1 To HOLDS
        If SheetAndRangeExists(STOWPLAN_SHEET_NAME, "HOLD" & holdNumber) Then
            For Each cell In HOLD_RANGE(holdNumber)
                colorCode = CLng(cell.Interior.color)
                If cell.Interior.colorIndex <> xlColorIndexNone And this.DischargingPorts.Exists(colorCode) And WorksheetFunction.IsNumber(cell.Value2) = True Then
                    Set currentPort = this.DischargingPorts.Item(colorCode)
                    Set currentHold = currentPort.HoldData.Item(holdNumber)
                    If Int(cell.Value2) = cell.Value2 Then
                        If this.UnitsCells Is Nothing Then
                            Set this.UnitsCells = cell
                        Else
                            Set this.UnitsCells = Union(this.UnitsCells, cell)
                        End If
                        currentHold.Count = currentHold.Count + cell.Value2
                        currentPort.TotalUnits = currentPort.TotalUnits + cell.Value2
                    Else
                        If this.WeightsCells Is Nothing Then
                            Set this.WeightsCells = cell
                        Else
                            Set this.WeightsCells = Union(this.WeightsCells, cell)
                        End If
                        currentHold.Weight = currentHold.Weight + cell.Value2
                        currentPort.TotalUnitsWeight = currentPort.TotalUnitsWeight + cell.Value2
                    End If
                End If
            Next
        End If
    Next holdNumber
End Sub

Private Sub SetPkgsProperty()

    On Error GoTo NotWellFormed
    
    Dim colorCode               As Long
    Dim pkgsShape               As Shape
    Dim textArray()             As String
    Dim first                   As Long
    Dim last                    As Long
    Dim lengthOfArray           As Long
    Dim currentPort             As Port
    
    For Each pkgsShape In STOWAGE_PLAN_SHAPES
        If Right$(pkgsShape.Name, Len(PACKAGE_TAG)) = PACKAGE_TAG Then
            colorCode = CLng(pkgsShape.Fill.ForeColor)
            If this.DischargingPorts.Exists(colorCode) Then
                textArray = MultiSplit(Trim$(pkgsShape.TextFrame2.TextRange), Chr$(10), " ")
                Set currentPort = this.DischargingPorts.Item(colorCode)
                first = LBound(textArray)
                last = UBound(textArray)
                lengthOfArray = last - first + 1
            
                If lengthOfArray < 8 Then
                    Err.Raise vbObjectError + 513, _
                              "SetPkgsProperty", _
                              "There seems to be a problem with the format of package text boxes for " & _
                              STOWAGE_PLAN_SHEEET.Cells.Item(currentPort.RowNumber, 2).Value2 & _
                              " with value: " & vbNewLine & pkgsShape.TextFrame2.TextRange.Text
                End If
                
                If Not IsNumeric(textArray(3)) Or Not IsNumeric(textArray(6)) Then
                    Err.Raise vbObjectError + 513, _
                              "SetPkgsProperty", _
                              "Expected values are not numbers. Please check packages for " & _
                              STOWAGE_PLAN_SHEEET.Cells.Item(currentPort.RowNumber, 2).Value2 & _
                              " with value: " & vbNewLine & pkgsShape.TextFrame2.TextRange.Text
                End If
                
                currentPort.TotalPkgs = currentPort.TotalPkgs + CInt(textArray(3))
                currentPort.TotalPkgsWeight = currentPort.TotalPkgsWeight + CDbl(textArray(6))
                
IfNotWellFormed:

            End If
        End If
    Next pkgsShape
    On Error GoTo 0
    
    Exit Sub
    
NotWellFormed:
    MsgBox Err.Description, vbCritical
    Err.Clear
    Resume IfNotWellFormed
End Sub

Private Function GetLoadingPortsSequence() As String()
    Dim cell                As Range
    Dim LoadPortsCodes()    As String:      ReDim LoadPortsCodes(0)
    For Each cell In LDG_PORTS_CODES_RANGE
        If cell.Value2 <> vbNullString Then
            ReDim Preserve LoadPortsCodes(UBound(LoadPortsCodes) + 1)
            LoadPortsCodes(UBound(LoadPortsCodes) - 1) = cell.Value2
        End If
    Next cell

    If UBound(LoadPortsCodes) - 1 >= 0 Then
        ReDim Preserve LoadPortsCodes(UBound(LoadPortsCodes) - 1)
    End If
    GetLoadingPortsSequence = LoadPortsCodes
End Function

Public Sub PrintDischPorts()
    Dim portObj As Variant
    For Each portObj In this.DischargingPorts.Items
        'Debug.Print portObj.Color, portObj.PortName, portObj.PortCode, portObj.RowNumber
    Next portObj
End Sub


