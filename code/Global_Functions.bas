Attribute VB_Name = "Global_Functions"
'@Folder "StowagePlan.utils"
Option Explicit
Option Private Module

Public Function TotalUnitsAndWeightsSum(ByVal sumRange As Range) As Double()
    Dim resultArray(1 To 2) As Double
    Dim cell As Range
    For Each cell In sumRange
        If cell.Interior.colorIndex <> xlColorIndexNone And WorksheetFunction.IsNumber(cell.Value2) = True Then
            If CInt(cell.Value2) = cell.Value2 Then
                resultArray(1) = resultArray(1) + cell.Value2
            Else
                resultArray(2) = resultArray(2) + cell.Value2
            End If
        End If
    Next
    TotalUnitsAndWeightsSum = resultArray
End Function

Public Function GetRowIndexByColorMap(ByVal portsList As Range) As Object
    Dim rowIndexByColor As Object
    Set rowIndexByColor = CreateObject("Scripting.Dictionary")
    Dim cell As Range
    Dim color As Long
    For Each cell In portsList
        color = cell.Interior.color
        If color > 0 And cell.Value2 <> vbNullString And Not rowIndexByColor.Exists(color) Then
            rowIndexByColor.Add CLng(color), cell.row
        End If
    Next
    Set GetRowIndexByColorMap = rowIndexByColor
End Function

Public Function SumNonContiguous(ByVal sumRange As Range, ByVal nth As Long) As Double
    Dim rangeArray As Variant: rangeArray = sumRange.Value2
    
    Dim sum     As Double
    Dim row     As Long
    Dim col     As Long
    
    For row = LBound(rangeArray, 1) To UBound(rangeArray, 1)
        For col = LBound(rangeArray, 2) To UBound(rangeArray, 2)
            If (col - 1) Mod 4 = nth And WorksheetFunction.IsNumber(rangeArray(row, col)) = True Then
                sum = sum + rangeArray(row, col)
            End If
        Next col
    Next row
    SumNonContiguous = sum
End Function

Public Function GetNumberOfPorts() As Long
    Dim Count As Long
    Dim stowagePlanDischPortsList As Variant
    stowagePlanDischPortsList = PORTS_LIST_RANGE.Value2
    Dim row As Long
    For row = LBound(stowagePlanDischPortsList, 1) To UBound(stowagePlanDischPortsList, 1)
        If stowagePlanDischPortsList(row, 1) <> vbNullString Then
            Count = Count + 1
        End If
    Next row
    GetNumberOfPorts = Count
End Function

Public Function ProperSubSet(ByRef range1 As Range, ByRef range2 As Range) As Boolean
    Dim cell As Range
    For Each cell In range2
        If Intersect(cell, range1) Is Nothing Then
            ProperSubSet = False
            Exit Function
        End If
    Next cell
    ProperSubSet = True
End Function

Public Function CreateTextArrayFromSourceTexts(ParamArray SourceTexts() As Variant) As String()
    Dim TargetTextArray() As String
    ReDim TargetTextArray(0 To UBound(SourceTexts))
    
    Dim SourceTextsCellNumber As Long
    
    For SourceTextsCellNumber = 0 To UBound(SourceTexts)
        TargetTextArray(SourceTextsCellNumber) = SourceTexts(SourceTextsCellNumber)
    Next SourceTextsCellNumber
    
    CreateTextArrayFromSourceTexts = TargetTextArray
End Function

Public Function GetGUID() As String
    GetGUID = Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)
End Function

