Attribute VB_Name = "MainDeckModule"
'@Folder "StowagePlan"
Option Explicit
Option Private Module

Const LEFT_COL As String = "B"
Const LEFT_COL_INDEX As Integer = 2
Const TARGET_ROW_INDEX As Integer = 3

Public Sub SetDestinationPortAction(control As IRibbonControl)
    SetDestinationPortForm.Show
End Sub

Public Sub SetDestinationPort()
    Debug.Print SetDestinationPortForm.PortsListDropDown.value, SetDestinationPortForm.PortsListDropDown.BackColor
End Sub

Sub DischargingPortSequence()
    Dim dischPorts As Variant
    dischPorts = DIS_PORTS_CODES_RANGE.Value2
    
    Dim colorByPortCode As Object
    Set colorByPortCode = CreateObject("Scripting.Dictionary")
    Dim cell As Range
    Dim colorValue As Long
    
    For Each cell In DIS_PORTS_CODES_RANGE
        colorValue = cell.Interior.color
        If colorValue > 0 And cell.Value2 <> vbNullString And Not colorByPortCode.Exists(cell.Value2) Then
            colorByPortCode.Add cell.Value2, colorValue
        End If
    Next
    
    Dim portCodes() As String
    ReDim portCodes(1 To 1)
    
    Dim dischSequence() As String
    ReDim dischSequence(0)
    
    Dim row As Long
    For row = LBound(dischPorts, 1) To UBound(dischPorts, 1)
        If dischPorts(row, 1) <> vbNullString Then
            ReDim Preserve dischSequence(UBound(dischSequence) + 2)
            dischSequence(UBound(dischSequence) - 2) = dischPorts(row, 1)
            dischSequence(UBound(dischSequence) - 1) = ">>>>"
        End If
    Next row
    ReDim Preserve dischSequence(UBound(dischSequence) - 2)
    
    Dim i As Long, colIndex As Integer
    Dim wsName As Variant
    For Each wsName In Array(MAIN_DECK_SHEET_NAME)
        With Sheets(wsName)
            With .Range(.Cells(TARGET_ROW_INDEX, LEFT_COL_INDEX).Address)
                .EntireRow.Interior.color = xlNone
                .EntireRow.Value2 = ""
            End With
            For i = 0 To UBound(dischSequence)
                colIndex = i + LEFT_COL_INDEX
                Set cell = .Range(.Cells(TARGET_ROW_INDEX, colIndex).Address)
                cell.Value2 = dischSequence(i)
                cell.HorizontalAlignment = xlCenter
                If colorByPortCode.Exists(dischSequence(i)) Then
                    cell.Interior.color = colorByPortCode(dischSequence(i))
                Else
                    cell.Interior.colorIndex = 15
                End If
            Next i
        End With
    Next wsName
End Sub

