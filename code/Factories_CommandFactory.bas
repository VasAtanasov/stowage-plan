Attribute VB_Name = "Factories_CommandFactory"
'@Folder "StowagePlan.factories"
Option Explicit

Public Function CreateCommand(ByVal control As IRibbonControl, ByRef manager As IManager) As ICommand
    Select Case control.ID
    Case UPDATE_TABLE_BUTTON_ID
        Set CreateCommand = New CommandUpdateTable
    Case PRE_ARRIVAL_BUTTON_ID
        Set CreateCommand = New CommandPreArrivalPlan
    Case EXPORT_DISCHARGE_PLAN_BUTTON_ID
        Set CreateCommand = New CommandExportDischargingPlan
    Case EXPORT_DEPARTURE_PLAN_BUTTON_ID
        Set CreateCommand = New CommandExportDeparturePlan
    Case BACK_UP_STOWAGE_PLAN_BUTTON_ID
        Set CreateCommand = New CommandBackUpStowagePlan
    Case MOVE_ROW_UP_BUTTON_ID, MOVE_ROW_DOWN_BUTTON_ID
        Set CreateCommand = New CommandMoveRow
    Case TOGGLE_ROWS_BUTTON_ID
        Set CreateCommand = New CommandToggleEmptyRows
    Case HEAD_TO_FORE_BUTTON_ID, _
         HEAD_TO_AFT_BUTTON_ID, _
         HEAD_TO_PORT_BUTTON_ID, _
         HEAD_TO_STBD_BUTTON_ID
        Set CreateCommand = New CommandCreateDirection
    Case ADD_RECTANGLE_CARGO_BOX_BUTTON_ID, _
         ADD_CALLOUT_CARGO_BOX_BUTTON_ID, _
         ADD_INFO_BOX_BUTTON_ID
        Set CreateCommand = New CommandSetBox
    Case Else
        Set CreateCommand = Nothing
    End Select
    If Not CreateCommand Is Nothing Then
        CreateCommand.InitiateProperties manager, control
    End If
End Function


