Attribute VB_Name = "ValidateCalculations"
'@Folder "StowagePlan"
Option Explicit
Option Private Module

Public Sub ValidateUnitsAndWeights(control As IRibbonControl)
    Call Validate
End Sub

Sub OnValidationSuccess()
    Application.ScreenUpdating = True
End Sub

Sub OnValidationError()
    Application.ScreenUpdating = True
End Sub

Private Sub Validate()
    Application.ScreenUpdating = False
    
    If Not IsDataByPortValid Then
        OnValidationError
        Exit Sub
    End If
    
    Dim hatchSummaryUnits As Integer
    Dim hatchSummaryUnitsWeights As Double
    Dim hatchSummaryPackages As Integer
    Dim hatchSummaryPackagesWeights As Double
    Dim hatchSummaryTotal As Integer
    Dim hatchSummaryTotalWehight As Double
    
    hatchSummaryUnits = HATCH_SUMMARY_SHEET.Range("D21").Value2
    hatchSummaryUnitsWeights = HATCH_SUMMARY_SHEET.Range("E21").Value2
    hatchSummaryPackages = HATCH_SUMMARY_SHEET.Range("F21").Value2
    hatchSummaryPackagesWeights = HATCH_SUMMARY_SHEET.Range("G21").Value2
    
    hatchSummaryTotal = hatchSummaryUnits + hatchSummaryPackages
    hatchSummaryTotalWehight = hatchSummaryUnitsWeights + hatchSummaryPackagesWeights
    
    Debug.Print "Hatch Summary:", hatchSummaryUnits, hatchSummaryUnitsWeights, hatchSummaryPackages, hatchSummaryPackagesWeights, hatchSummaryTotal, hatchSummaryTotalWehight
    
    Dim tableSummaryUnits As Integer
    Dim tableSummaryUnitsWeights As Double
    Dim tableSummaryPackages As Integer
    Dim tableSummaryPackagesWeights As Double
    Dim tableSummaryTotal As Integer
    Dim tableSummaryTotalWehight As Double
    
    tableSummaryUnits = Sheets(STOWPLAN_SHEET_NAME).Range("CI26").Value2
    tableSummaryUnitsWeights = Sheets(STOWPLAN_SHEET_NAME).Range("CN26").Value2
    tableSummaryPackages = Sheets(STOWPLAN_SHEET_NAME).Range("CS26").Value2
    tableSummaryPackagesWeights = Sheets(STOWPLAN_SHEET_NAME).Range("CX26").Value2
    
    tableSummaryTotal = tableSummaryUnits + tableSummaryPackages
    tableSummaryTotalWehight = tableSummaryUnitsWeights + tableSummaryPackagesWeights
    
    Debug.Print "Table Summary:", tableSummaryUnits, tableSummaryUnitsWeights, tableSummaryPackages, tableSummaryPackagesWeights, tableSummaryTotal, tableSummaryTotalWehight
    
    Dim bottomSummaryUnits As Integer
    Dim bottomSummaryUnitsWeights As Double
    Dim bottomSummaryTotal As Integer
    Dim bottomSummaryTotalWehight As Double
    
    bottomSummaryUnits = Sheets(STOWPLAN_SHEET_NAME).Range("CT140").Value2
    bottomSummaryUnitsWeights = Sheets(STOWPLAN_SHEET_NAME).Range("CT141").Value2
    
    bottomSummaryTotal = bottomSummaryUnits + tableSummaryPackages
    bottomSummaryTotalWehight = bottomSummaryUnitsWeights + tableSummaryPackagesWeights
    
    Debug.Print "Bottom Summary:", bottomSummaryUnits, bottomSummaryUnitsWeights, , , bottomSummaryTotal, bottomSummaryTotalWehight
    
    If tableSummaryUnits <> hatchSummaryUnits Or hatchSummaryUnits <> bottomSummaryUnits Then
        OnValidationError
        MsgBox "Difference in units:" & vbNewLine & _
               "Hatch Summary Units: " & hatchSummaryUnits & vbNewLine & _
               "Table Summary Units: " & tableSummaryUnits & vbNewLine & _
               "Table Bottom Units: " & bottomSummaryUnits, _
               vbCritical
        Exit Sub
    End If
    
    If Not DblSafeCompare(hatchSummaryUnitsWeights, tableSummaryUnitsWeights) Or Not DblSafeCompare(tableSummaryUnitsWeights, bottomSummaryUnitsWeights) Then
        OnValidationError
        MsgBox "Difference in units weights:" & vbNewLine & _
               "Hatch Summary Units Weights: " & hatchSummaryUnitsWeights & vbNewLine & _
               "Table Summary Units Weights: " & tableSummaryUnitsWeights & vbNewLine & _
               "Table Bottom Units Weights: " & bottomSummaryUnitsWeights, _
               vbCritical
        Exit Sub
    End If
    
    If hatchSummaryPackages <> tableSummaryPackages Then
        OnValidationError
        MsgBox "Difference in packages:" & vbNewLine & _
               "Hatch Summary Packages: " & hatchSummaryPackages & vbNewLine & _
               "Table Summary Packages: " & tableSummaryPackages & vbNewLine, _
               vbCritical
        Exit Sub
    End If
    
    If Not DblSafeCompare(hatchSummaryPackagesWeights, tableSummaryPackagesWeights) Then
        OnValidationError
        MsgBox "Difference in units weights:" & vbNewLine & _
               "Hatch Summary Packages Weights: " & hatchSummaryPackagesWeights & vbNewLine & _
               "Table Summary Packages Weights: " & tableSummaryPackagesWeights & vbNewLine, _
               vbCritical
        Exit Sub
    End If
    
    OnValidationSuccess
    MsgBox "Input data is valid", vbInformation
End Sub

Private Function IsDataByPortValid() As Boolean
    Dim dischPortsCount As Long: dischPortsCount = GetNumberOfPorts
    
    Dim tableSummary As Variant
    tableSummary = CARGO_SUMMARY_TABLE_RANGE.Value2
    
    Dim hatchSummary As Variant
    Dim hatchSummaryPorts As Variant
    hatchSummary = HATCH_SUMMARY_SHEET.Range("H21:BO21").Value2
    hatchSummaryPorts = HATCH_SUMMARY_SHEET.Range("H4:BO4").Value2
    
    Dim tableSummaryData() As Variant
    ReDim tableSummaryData(1 To dischPortsCount, 1 To 5)
    Dim dataPositions(), dataType(), dataUnit()
    dataPositions = Array(1, 86, 91, 96, 101)
    dataType = Array("port name", "total units", "total units weight", "total packages", "total packges weight")
    dataUnit = Array("", "U/s", "t", "Pkgs", "t")
    
    Dim position As Integer
        
    Dim row As Long, col As Long, i As Long
    For row = LBound(tableSummary, 1) To UBound(tableSummary, 1)
        If tableSummary(row, 1) <> vbNullString Then
            
            For i = 1 To 5
                position = dataPositions(i - 1)
                tableSummaryData(row, i) = IIf(tableSummary(row, position) <> vbNullString Or tableSummary(row, position) <> "", tableSummary(row, position), 0)
            Next i
            
            Debug.Print "Port: " & tableSummaryData(row, 1)
            Debug.Print "Table Summary: ", tableSummaryData(row, 1), tableSummaryData(row, 2), tableSummaryData(row, 3), tableSummaryData(row, 4), tableSummaryData(row, 5)
            
            col = (row - 1) * 4 + 1
            
            tableSummaryData(row, 1) = tableSummaryData(row, 1) = Left(hatchSummaryPorts(1, col), Len(tableSummaryData(row, 1)))
            
            If Not tableSummaryData(row, 1) Then
                MsgBox "Non matchin port names. Difference in discharging ports sequence or missing port name." & vbNewLine & _
                       "Table Summary Port: " & tableSummary(row, 1) & vbNewLine & _
                       "Hatch Summary Port: " & hatchSummaryPorts(1, col), _
                       vbCritical
                Exit Function
            End If
            
            tableSummaryData(row, 2) = tableSummaryData(row, 2) = hatchSummary(1, col)
            tableSummaryData(row, 3) = DblSafeCompare(CDbl(tableSummaryData(row, 3)), CDbl(hatchSummary(1, col + 1)))
            tableSummaryData(row, 4) = tableSummaryData(row, 4) = hatchSummary(1, col + 2)
            tableSummaryData(row, 5) = DblSafeCompare(CDbl(tableSummaryData(row, 5)), CDbl(hatchSummary(1, col + 3)))

            
            Debug.Print "Hatch Summary: ", hatchSummaryPorts(1, col), hatchSummary(1, col), hatchSummary(1, col + 1), hatchSummary(1, col + 2), hatchSummary(1, col + 3)
            Debug.Print "Is Matching Summary: ", tableSummaryData(row, 1), tableSummaryData(row, 2), tableSummaryData(row, 3), tableSummaryData(row, 4), tableSummaryData(row, 5)
            Debug.Print "================================================================================"
            
            For i = 1 To 5
                If Not tableSummaryData(row, i) Then
                    MsgBox "There is a difference in " & dataType(i - 1) & " data." & vbNewLine & _
                                                                         "Port Name: " & tableSummary(row, 1) & vbNewLine & _
                                                                         "Table Summary Data: " & tableSummary(row, dataPositions(i - 1)) & " " & dataUnit(i - 1) & vbNewLine & _
                                                                         "Hatch Summary Data: " & hatchSummary(1, col + i - 2) & " " & dataUnit(i - 1), _
                                                                         vbCritical
                    Exit Function
                End If
            Next i
        End If
    Next row
    
    IsDataByPortValid = True
    
End Function


