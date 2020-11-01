Attribute VB_Name = "Feature_PreArrivalPlan_Create"
'@Folder "StowagePlan.feature.pre_arrival_plan"
Option Explicit
Option Private Module

Private Sub CopySheet()
    Dim wsNamesForExport() As String
    Dim wsForExport As Variant
    
    'Add additional sheets for export if needed
    wsNamesForExport = CreateTextArrayFromSourceTexts(DISCHARGE_PLAN_SHEET_NAME, DISCHARGE_PLAN_MAIN_DECK_SHEET_NAME)
    wsForExport = Array(STOWAGE_PLAN_SHEEET, MAIN_DECK_SHEEET)
    
    Dim index As Long
    
    For index = 0 To UBound(wsNamesForExport)
        If WorksheetExists(wsNamesForExport(index)) Then
            STOWAG_PLAN_BOOK.Worksheets.Item(wsNamesForExport(index)).Delete
        End If
    Next index
    
    For index = 0 To UBound(wsForExport)
        wsForExport(index).Copy After:=LAST_WORKSHEET
        LAST_WORKSHEET.Name = wsNamesForExport(index)
    Next index
End Sub

Private Function GetFirstDischargingPort() As Range
    Dim Port    As Range
    For Each Port In PORTS_LIST_RANGE
        If Trim$(Port) <> vbNullString Then
            Set GetFirstDischargingPort = Port
            Exit Function
        End If
    Next Port
End Function

Public Sub CreatePreArrivalPlan()
    Dim selectedPort As Range
    Set selectedPort = GetFirstDischargingPort
        
    If Not ProperSubSet(PORTS_LIST_RANGE, selectedPort) Then
        MsgBox "Please select port from the ports list"
        Exit Sub
    End If
    
    If selectedPort.Rows.Count <> 1 Then
        MsgBox "Invalid selection. Select only one port"
        Exit Sub
    End If
    
    If Trim$(selectedPort.Value2) = vbNullString Then
        MsgBox "Invalid selection. Empty cell"
        Exit Sub
    End If
    
    OptimizeVBA True
    
    'ask for date of discharging plan ex InputBox
    
    'Dim TheString As String, TheDate As Date
    'TheString = Application.InputBox("Enter A Date")
    'If IsDate(TheString) Then
    '    TheDate = DateValue(TheString)
    'Else
    '    MsgBox "Invalid date"
    'End If
    
    CopySheet
    STOWAG_PLAN_BOOK.Save
    
    Dim dischargingPortCell As Range
    Set dischargingPortCell = DISCHARGE_PLAN_SHEET.Range("BU3")
    dischargingPortCell.Value2 = selectedPort.Value2
    DISCHARGE_PLAN_SHEET.Range("BO3").Value2 = "Arrival:"
    DISCHARGE_PLAN_SHEET.Range("AL2").Value2 = "DISCHARGING PLAN"
    
    Dim selectedPortColor   As Long
    Dim rowIndexByColor     As Object
    
    Set rowIndexByColor = GetRowIndexByColorMap(PORTS_LIST_RANGE)
    selectedPortColor = CLng(selectedPort.Interior.color)
    
    Dim cell            As Range
    Dim Hold            As Long
    Dim colorCode       As Long
    Dim holdRange       As Range
    Dim mCells          As Range
    Dim shp             As Shape
    Dim portRow         As Long
    
    DISCHARGE_PLAN_SHEET.Activate
    DISCHARGE_PLAN_SHEET.Range("A1").Select
    
    Dim colorCodeKey As Variant
    For Each colorCodeKey In rowIndexByColor.Keys
        If colorCodeKey <> selectedPortColor Then
            For Hold = 1 To HOLDS
                Set mCells = Nothing
                If SheetAndRangeExists(DISCHARGE_PLAN_SHEET_NAME, "HOLD" & Hold) Then
                    Set holdRange = DISCHARGE_PLAN_HOLD_RANGE(Hold)
                    For Each cell In holdRange
                        colorCode = CLng(cell.Interior.color)
                        If cell.Interior.colorIndex <> xlColorIndexNone And colorCode = colorCodeKey Then
                            If mCells Is Nothing Then
                                Set mCells = cell
                            Else
                                Set mCells = Union(mCells, cell)
                            End If
                        End If
                    Next
                    If Not mCells Is Nothing Then
                        mCells.Interior.colorIndex = 15
                    End If
                End If
            Next Hold

            'TODO Cycle trough stowage plan and deck 5 sheets and change shapes colors
            For Each shp In DISCHARGE_PLAN_SHAPES
                If Right$(shp.Name, Len(PACKAGE_TAG)) = PACKAGE_TAG Then
                    colorCode = CLng(shp.Fill.ForeColor)
                    If rowIndexByColor.Exists(colorCode) And colorCode = colorCodeKey Then
                        shp.Fill.ForeColor.RGB = RGB(192, 192, 192)
                    End If
                End If
            Next shp
        End If
    Next colorCodeKey
    
    
    For portRow = TABLE_TOP_ROW To TABLE_BOTTOM_ROW
        If portRow <> selectedPort.row Then
            DISCHARGE_PLAN_CARGO_TABLE_ROW(portRow).Interior.colorIndex = 15
        End If
    Next portRow
    
    OptimizeVBA False
 
End Sub

Sub StartDate()

    Dim strDate As String
    Dim acceptDate As Integer
    
    Do
        Do
            strDate = InputBox("Please Enter the PRODUCTION REPORTING DATE as MM/DD/YYYY", "Production Reporting Date", Format(Date, "dd-mm-yyyy"))
            If Not IsDate(strDate) Then MsgBox "Please enter a production date!", vbCritical
        Loop Until IsDate(strDate)
        strDate = Format(CDate(strDate), "mm/dd/yyyy")
        acceptDate = MsgBox("The PRODUCTION DATE you entered is " & strDate & vbNewLine & "Accept this date?", vbYesNo)
    Loop Until acceptDate = vbYes
    
End Sub



