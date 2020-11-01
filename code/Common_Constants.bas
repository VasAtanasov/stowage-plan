Attribute VB_Name = "Common_Constants"
'@Folder "StowagePlan.common"
Option Explicit

Public tableColors As Variant

Public Const VESSEL_CODE                                As String = "ERSH"
Public Const WORKBOOK_NAME                              As String = "Stowage plan.xlsm"
Public Const BACKUP_FOLDR_NAME                          As String = "backup-stowage"

Public Const DECKS                                      As Long = 12
Public Const HOLDS                                      As Long = 4
Public Const TABLE_TOP_ROW                              As Long = 9
Public Const TABLE_BOTTOM_ROW                           As Long = 23
Public Const TABLE_LEFT_COL                             As String = "B"
Public Const TABLE_RIGHT_COL                            As String = "DB"
Public Const UNITS_FIRST_COL                            As Long = 46
Public Const WEIGHTS_FIRST_COL                          As Long = 51
Public Const HOLD_COL_SPREAD                            As Long = 10


Public Const LOADING_SUMMARY_TOP_ROW                    As Long = 7
Public Const LOADING_SUMMARY_BOTTOM_ROW                 As Long = 19

'Sheets names
Public Const STOWPLAN_SHEET_NAME                        As String = "Stowage Plan"
Public Const DISCHARGE_PLAN_SHEET_NAME                  As String = "Discharging Plan"
Public Const DISCHARGE_PLAN_MAIN_DECK_SHEET_NAME        As String = "Discharging Plan Main Deck"
Public Const MAIN_DECK_SHEET_NAME                       As String = "Main Deck"
Public Const PANEL_PLANE_SHEET_NAME                     As String = "Panel Plan"
Public Const HATCH_SUMMARY_SHEET_NAME                   As String = "Hatch Summary"

'Stowage plan ranges names
Public Const UPPER_DECK_RANGE_NAME                      As String = "UPPER_DECK"
Public Const LOWER_DECK_RANGE_NAME                      As String = "LOWER_DECK"
Public Const CARGO_SUMMARY_TABLE_RANGE_NAME             As String = TABLE_LEFT_COL & TABLE_TOP_ROW & ":" & TABLE_RIGHT_COL & TABLE_BOTTOM_ROW
Public Const PORTS_LIST_RANGE_NAME                      As String = "B" & TABLE_TOP_ROW & ":" & "M" & TABLE_BOTTOM_ROW '"PORTS_LIST"
Public Const HOLD_SUMMARY_RANGE_NAME                    As String = "AU" & TABLE_TOP_ROW & ":CH" & TABLE_BOTTOM_ROW '"HOLD_SUMMARY"
Public Const TOTAL_UNITS_SUMMARY_RANGE_NAME             As String = "CI" & TABLE_TOP_ROW & ":CR" & TABLE_BOTTOM_ROW '"TOTAL_UNITS_SUMMARY"
Public Const LOADING_SUMMARY_RANGE_NAME                 As String = "Q" & TABLE_TOP_ROW & ":AT" & TABLE_BOTTOM_ROW '"LOADING_SUMMARY"
Public Const PACKAGES_SUMMARY_RANGE_NAME                As String = "CS" & TABLE_TOP_ROW & ":DB" & TABLE_BOTTOM_ROW '"PACKAGE_SUMMARY"
Public Const DIS_PORTS_CODES_RANGE_NAME                 As String = "N" & TABLE_TOP_ROW & ":P" & TABLE_BOTTOM_ROW
Public Const LDG_PORTS_CODES_RANGE_NAME                 As String = "$Q$7:$AT$7"

Public Const HOLD1_FOR_UNITS                    As String = "BY"
Public Const HOLD2_FOR_UNITS                    As String = "BO"
Public Const HOLD3_FOR_UNITS                    As String = "BE"
Public Const HOLD4_FOR_UNITS                    As String = "AU"

Public Const HOLD1_FOR_WEIGTHS                  As String = "CD"
Public Const HOLD2_FOR_WEIGTHS                  As String = "BT"
Public Const HOLD3_FOR_WEIGTHS                  As String = "BJ"
Public Const HOLD4_FOR_WEIGTHS                  As String = "AZ"

Public Const COL_FOR_PORT_TOTAL_UNITS           As String = "CI"
Public Const COL_FOR_PORT_TOTAL_WEIGHTS         As String = "CN"

Public Const COL_FOR_PORT_TOTAL_PKGS_COUNT      As String = "CS"
Public Const COL_FOR_PORT_TOTAL_PKGS_WEIGHTS    As String = "CX"

Public Const STOWAGE_PLAN_TOTAL_LOADED_CELL     As String = "AP26"

'Hatch summary ranges names
Public Const HATCH_SUMMARY_TABLE_RANGE_NAME     As String = "B" & LOADING_SUMMARY_TOP_ROW & ":BO" & LOADING_SUMMARY_BOTTOM_ROW

'Discharging Plan cells
Public Const DISCH_PLAN_VOYAGE_NUMBER_CELL      As String = "BU2"

'Common constants
Public Const UNITS_FORMAT                       As String = "0""U/s"""
Public Const WEIGHT_FORMAT                      As String = "0.0""mt"""
Public Const PACKING_UNITS                      As String = "U/s"
Public Const PACKING_PKGS                       As String = "pkgs"
Public Const PACKAGE_TAG                        As String = "_PKGS"
Public Const INFO_BOX_TAG                       As String = "_INFO"
Public Const STOW_DORECTION_TAG                 As String = "STOW_DIRECTION"
Public Const STOWAGE_PLAN_DEFAULT_SHAPE_TAG     As String = "STOWAGE_PLAN_DEFAULT_SHAPE"

'Useful functions
Public Function STOWAGE_PLAN_DATE() As String
    STOWAGE_PLAN_DATE = Format$(STOWAGE_PLAN_SHEEET.Range("CG3").Value2, "yyyy-mm-dd")
End Function

Public Function CURRENT_PORT() As String
    CURRENT_PORT = STOWAGE_PLAN_SHEEET.Range("BU3").Value2
End Function

Public Function CURRENT_VOY() As String
    CURRENT_VOY = STOWAGE_PLAN_SHEEET.Range("BU2").Value2
End Function

