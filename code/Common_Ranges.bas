Attribute VB_Name = "Common_Ranges"
'@Folder "StowagePlan.common"
Option Explicit

Public Function STOWAG_PLAN_BOOK() As Workbook
    Set STOWAG_PLAN_BOOK = Workbooks.Item(WORKBOOK_NAME)
End Function

Public Function LAST_WORKSHEET() As Worksheet
    Set LAST_WORKSHEET = STOWAG_PLAN_BOOK.Worksheets.Item(STOWAG_PLAN_BOOK.Worksheets.Count)
End Function

' -----------------------------------------------------------------------------------
' Stowage plan sheet common ranges
' Start

Public Function STOWAGE_PLAN_SHEEET() As Worksheet
    Set STOWAGE_PLAN_SHEEET = STOWAG_PLAN_BOOK.Worksheets.Item(STOWPLAN_SHEET_NAME)
End Function

Public Function CARGO_SUMMARY_TABLE_RANGE() As Range
    Set CARGO_SUMMARY_TABLE_RANGE = STOWAGE_PLAN_SHEEET.Range(CARGO_SUMMARY_TABLE_RANGE_NAME)
End Function

Public Function UPPER_DECK_RANGE() As Range
    Set UPPER_DECK_RANGE = STOWAGE_PLAN_SHEEET.Range(UPPER_DECK_RANGE_NAME)
End Function

Public Function LOWER_DECK_RANGE() As Range
    Set LOWER_DECK_RANGE = STOWAGE_PLAN_SHEEET.Range(LOWER_DECK_RANGE_NAME)
End Function

Public Function PORTS_LIST_RANGE() As Range
    Set PORTS_LIST_RANGE = STOWAGE_PLAN_SHEEET.Range(PORTS_LIST_RANGE_NAME)
End Function

Public Function HOLD_RANGE(ByVal Hold As Long) As Range
    Set HOLD_RANGE = STOWAGE_PLAN_SHEEET.Range("HOLD" & Hold)
End Function

Public Function HOLD_SUMMARY_RANGE() As Range
    Set HOLD_SUMMARY_RANGE = STOWAGE_PLAN_SHEEET.Range(HOLD_SUMMARY_RANGE_NAME)
End Function

Public Function TOTAL_UNITS_SUMMARY_RANGE() As Range
    Set TOTAL_UNITS_SUMMARY_RANGE = STOWAGE_PLAN_SHEEET.Range(TOTAL_UNITS_SUMMARY_RANGE_NAME)
End Function

Public Function LOADING_SUMMARY_RANGE() As Range
    Set LOADING_SUMMARY_RANGE = STOWAGE_PLAN_SHEEET.Range(LOADING_SUMMARY_RANGE_NAME)
End Function

Public Function PACKAGES_SUMMARY_RANGE() As Range
    Set PACKAGES_SUMMARY_RANGE = STOWAGE_PLAN_SHEEET.Range(PACKAGES_SUMMARY_RANGE_NAME)
End Function

Public Function DIS_PORTS_CODES_RANGE() As Range
    Set DIS_PORTS_CODES_RANGE = STOWAGE_PLAN_SHEEET.Range(DIS_PORTS_CODES_RANGE_NAME)
End Function

Public Function LDG_PORTS_CODES_RANGE() As Range
    Set LDG_PORTS_CODES_RANGE = STOWAGE_PLAN_SHEEET.Range(LDG_PORTS_CODES_RANGE_NAME)
End Function

Public Function PORT_TOTAL_UNITS_CELL(ByVal cellRow As Long) As Range
    Set PORT_TOTAL_UNITS_CELL = STOWAGE_PLAN_SHEEET.Range(COL_FOR_PORT_TOTAL_UNITS & cellRow)
End Function

Public Function PORT_TOTAL_WEIGHTS_CELL(ByVal cellRow As Long) As Range
    Set PORT_TOTAL_WEIGHTS_CELL = STOWAGE_PLAN_SHEEET.Range(COL_FOR_PORT_TOTAL_WEIGHTS & cellRow)
End Function

Public Function PORT_TOTAL_UNITS_FOR_HOLD_CELL(ByVal Hold As Long, ByVal cellRow As Long) As Range
    Dim columnNameForUnits As String
    Select Case Hold
    Case 1
        columnNameForUnits = HOLD1_FOR_UNITS
    Case 2
        columnNameForUnits = HOLD2_FOR_UNITS
    Case 3
        columnNameForUnits = HOLD3_FOR_UNITS
    Case 4
        columnNameForUnits = HOLD4_FOR_UNITS
    End Select
    Set PORT_TOTAL_UNITS_FOR_HOLD_CELL = STOWAGE_PLAN_SHEEET.Range(columnNameForUnits & cellRow)
End Function

Public Function PORT_TOTAL_WEIGHTS_FOR_HOLD_CELL(ByVal Hold As Long, ByVal cellRow As Long) As Range
    Dim columnNameForWeights As String
    Select Case Hold
    Case 1
        columnNameForWeights = HOLD1_FOR_WEIGTHS
    Case 2
        columnNameForWeights = HOLD2_FOR_WEIGTHS
    Case 3
        columnNameForWeights = HOLD3_FOR_WEIGTHS
    Case 4
        columnNameForWeights = HOLD4_FOR_WEIGTHS
    End Select
    Set PORT_TOTAL_WEIGHTS_FOR_HOLD_CELL = STOWAGE_PLAN_SHEEET.Range(columnNameForWeights & cellRow)
End Function

Public Function PORT_TOTAL_PACKAGES_COUNT_CELL(ByVal portRow As Long) As Range
    Set PORT_TOTAL_PACKAGES_COUNT_CELL = STOWAGE_PLAN_SHEEET.Range(COL_FOR_PORT_TOTAL_PKGS_COUNT & portRow)
End Function

Public Function PORT_TOTAL_PACKAGES_WEIGHT_CELL(ByVal portRow As Long) As Range
    Set PORT_TOTAL_PACKAGES_WEIGHT_CELL = STOWAGE_PLAN_SHEEET.Range(COL_FOR_PORT_TOTAL_PKGS_WEIGHTS & portRow)
End Function

Public Function STOWAGE_PLAN_LOADING_PORTS_TOTAL() As Range
    Set STOWAGE_PLAN_LOADING_PORTS_TOTAL = STOWAGE_PLAN_SHEEET.Range(STOWAGE_PLAN_TOTAL_LOADED_CELL)
End Function

Public Function STOWAGE_PLAN_CARGO_TABLE_ROW(ByVal rowNuber As Long) As Range
    Set STOWAGE_PLAN_CARGO_TABLE_ROW = STOWAGE_PLAN_SHEEET.Range(TABLE_LEFT_COL & rowNuber & ":" & TABLE_RIGHT_COL & rowNuber)
End Function

Public Function STOWAGE_PLAN_SHAPES() As Variant
    Set STOWAGE_PLAN_SHAPES = STOWAGE_PLAN_SHEEET.Shapes
End Function

' End
' -----------------------------------------------------------------------------------

' -----------------------------------------------------------------------------------
' Hatch Summary sheet common ranges
' Start

Public Function HATCH_SUMMARY_SHEET() As Worksheet
    Set HATCH_SUMMARY_SHEET = STOWAG_PLAN_BOOK.Worksheets.Item(HATCH_SUMMARY_SHEET_NAME)
End Function

Public Function HATCH_SUMMARY_TABLE_RANGE() As Range
    Set HATCH_SUMMARY_TABLE_RANGE = HATCH_SUMMARY_SHEET.Range(HATCH_SUMMARY_TABLE_RANGE_NAME)
End Function

' End
' -----------------------------------------------------------------------------------

' -----------------------------------------------------------------------------------
' Main deck sheet common ranges
' Start

Public Function MAIN_DECK_SHEEET() As Worksheet
    Set MAIN_DECK_SHEEET = STOWAG_PLAN_BOOK.Worksheets.Item(MAIN_DECK_SHEET_NAME)
End Function

' End
' -----------------------------------------------------------------------------------

' -----------------------------------------------------------------------------------
' Discharging Plan sheet common ranges
' Start

Public Function DISCHARGE_PLAN_SHEET() As Worksheet
    Set DISCHARGE_PLAN_SHEET = STOWAG_PLAN_BOOK.Worksheets.Item(DISCHARGE_PLAN_SHEET_NAME)
End Function

Public Function DISCHARGING_PLAN_VOYAGE_NUMBER() As Range
    Set DISCHARGING_PLAN_VOYAGE_NUMBER = DISCHARGE_PLAN_SHEET.Range(DISCH_PLAN_VOYAGE_NUMBER_CELL)
End Function

Public Function DISCHARGE_PLAN_SHAPES() As Variant
    Set DISCHARGE_PLAN_SHAPES = DISCHARGE_PLAN_SHEET.Shapes
End Function

Public Function DISCHARGE_PLAN_HOLD_RANGE(ByVal Hold As Long) As Range
    Set DISCHARGE_PLAN_HOLD_RANGE = DISCHARGE_PLAN_SHEET.Range("HOLD" & Hold)
End Function

Public Function DISCHARGE_PLAN_CARGO_TABLE_ROW(ByVal rowNuber As Long) As Range
    Set DISCHARGE_PLAN_CARGO_TABLE_ROW = DISCHARGE_PLAN_SHEET.Range(TABLE_LEFT_COL & rowNuber & ":" & TABLE_RIGHT_COL & rowNuber)
End Function

' End
' -----------------------------------------------------------------------------------

' -----------------------------------------------------------------------------------
' Discharging Plan Main Deck sheet common ranges
' Start
