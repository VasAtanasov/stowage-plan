Attribute VB_Name = "Global_Tools"
'@Folder "StowagePlan.utils"
Option Explicit
Option Private Module

Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    ribbon.ActivateTab controlId:="CargoPlanTab"
End Sub

Public Function NameRangeExists(namedRange As String) As Boolean
    Dim nm As Name
    NameRangeExists = False
    For Each nm In ThisWorkbook.Names
        If nm.Name = namedRange Then
            NameRangeExists = True
            Exit For
        End If
    Next nm
End Function

Public Function SheetAndRangeExists(WorksheetName As String, rangeName As String) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = Worksheets(WorksheetName).Range(rangeName)
    On Error GoTo 0
    SheetAndRangeExists = Not rng Is Nothing
End Function

Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Public Function GetStowageTableColorsMap(ByRef portsList As Range)
    Dim colorIndexByColorCode As Object
    Set colorIndexByColorCode = CreateObject("Scripting.Dictionary")
    Dim cell As Range
    Dim colorIndex As Integer
    For Each cell In portsList
        colorIndex = cell.Interior.colorIndex
        If colorIndex > 0 And cell.Value2 <> vbNullString And Not colorIndexByColorCode.Exists(cell.Interior.color) Then
            colorIndexByColorCode.Add cell.Interior.color, colorIndex
        End If
    Next
    Set GetStowageTableColorsMap = colorIndexByColorCode
End Function

Public Function GetDischargingPortsMap(ByRef portsList As Range) As Object
    Dim porstByColorIndex As Object
    Set porstByColorIndex = CreateObject("Scripting.Dictionary")
    Dim cell As Range
    Dim colorIndex As Integer
    For Each cell In portsList
        colorIndex = cell.Interior.colorIndex
        If colorIndex > 0 And cell.Value2 <> vbNullString And Not porstByColorIndex.Exists(colorIndex) Then
            porstByColorIndex.Add colorIndex, cell.row
        End If
    Next
    Set GetDischargingPortsMap = porstByColorIndex
End Function

Public Function MultiSplit(SourceText As String, ParamArray Delimiters()) As String()
    Dim v As Variant
    For Each v In Delimiters
        SourceText = Replace(SourceText, v, "•")
    Next
    MultiSplit = FilterArrayString(Split(SourceText, "•"))
End Function

Private Function FilterArrayString(ByRef arr)
    Dim strTmp As Variant
    Dim outputArray() As String
    ReDim Preserve outputArray(0)
    For Each strTmp In arr
        If Trim(strTmp) <> "" Then
            ReDim Preserve outputArray(UBound(outputArray) + 1)
            outputArray(UBound(outputArray) - 1) = strTmp
        End If
    Next strTmp
    ReDim Preserve outputArray(UBound(outputArray) - 1)
    FilterArrayString = outputArray
End Function

Public Sub SaveAsXLSX(fileName As String)
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs fileName, 51             '51 = xlsx
    Application.DisplayAlerts = True
End Sub

Public Sub CreateBackUpFolder()
    Dim NewFolderPath As String
    NewFolderPath = Environ("UserProfile") & Application.PathSeparator & BACKUP_FOLDR_NAME
    If Len(Dir(NewFolderPath, vbDirectory)) = 0 Then
        MkDir NewFolderPath
    End If
End Sub

Public Sub CreateCurrentVoyageFolder(voyageNumber As String)
    Dim NewFolderPath As String
    NewFolderPath = ThisWorkbook.Path & Application.PathSeparator & voyageNumber
    If Len(Dir(NewFolderPath, vbDirectory)) = 0 Then
        MkDir NewFolderPath
    End If
End Sub

Public Sub RenameDefaultShapes()
    Dim shp As Shape
    Dim shpRng As Range
    Dim i As Integer
    i = 1
    Dim wsName As Variant
    For Each wsName In Array(STOWPLAN_SHEET_NAME, MAIN_DECK_SHEET_NAME, PANEL_PLANE_SHEET_NAME)
        For Each shp In Sheets(wsName).Shapes
            If Left(shp.Name, Len(STOW_DORECTION_TAG)) <> STOW_DORECTION_TAG And Right(shp.Name, Len(PACKAGE_TAG)) <> PACKAGE_TAG Then
                shp.Name = STOWAGE_PLAN_DEFAULT_SHAPE_TAG & "_" & (Int((999999 - 100 + 1) * Rnd + 1))
                i = i + 1
                Debug.Print shp.Name
            End If
        Next shp
    Next wsName
End Sub

Public Function DblSafeCompare(ByVal Value1 As Variant, ByVal Value2 As Variant) As Boolean
    'Compares two variants, dates and floats are compared at high accuracy
    Const AccuracyLevel As Double = 0.00000001
    'We accept an error of 0.000001% of the value
    Const AccuracyLevelSingle As Single = 0.0001
    'We accept an error of 0.0001 on singles
    If VarType(Value1) <> VarType(Value2) Then Exit Function
    Select Case VarType(Value1)
    Case vbSingle
        DblSafeCompare = Abs(Value1 - Value2) <= (AccuracyLevelSingle * Abs(Value1))
    Case vbDouble
        DblSafeCompare = Abs(Value1 - Value2) <= (AccuracyLevel * Abs(Value1))
    Case vbDate
        DblSafeCompare = Abs(CDbl(Value1) - CDbl(Value2)) <= (AccuracyLevel * Abs(CDbl(Value1)))
    Case vbNull
        DblSafeCompare = True
    Case Else
        DblSafeCompare = Value1 = Value2
    End Select
End Function

Sub CheckTableColorsConflict()
    Dim cell    As Range
    Dim row     As Long
    Dim i       As Long
    
    For Each cell In STOWAGE_PLAN_SHEEET.Range(TABLE_LEFT_COL & TABLE_TOP_ROW & ":" & TABLE_LEFT_COL & TABLE_BOTTOM_ROW)
        row = cell.row + 1
        For i = row To TABLE_BOTTOM_ROW
            If STOWAGE_PLAN_SHEEET.Range(TABLE_LEFT_COL & i).Interior.colorIndex = cell.Interior.colorIndex Then
                MsgBox "Matching color indexes at row " & cell.row & " row " & i
                Exit Sub
            End If
        Next i
    Next cell
End Sub

Sub ChangeColor()
    'Debug.Print Sheets(STOWPLAN_SHEET_NAME).Range(CARGO_SUMMARY_TABLE_ROW(23)).Interior.colorIndex
    'Sheets(STOWPLAN_SHEET_NAME).Range(CARGO_SUMMARY_TABLE_ROW(22)).Interior.colorIndex = 35
End Sub

Sub MergeUnmergeCells()
    Dim rng As Range
    Set rng = Selection
    Dim rsltVal As String

    If rng.MergeCells Then
        rng.UnMerge
    Else
        For Each cl In rng.Cells
            rsltVal = Trim(rsltVal) & " " & Trim(cl.value)
        Next cl
        
        rng.ClearContents
        rng.Merge
        rng.value = Trim(rsltVal)
        rng.HorizontalAlignment = xlCenter
        rng.VerticalAlignment = xlCenter
    End If
End Sub

