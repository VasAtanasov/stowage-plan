Attribute VB_Name = "Feature_Table_ToggleRows"
'@Folder "StowagePlan.feature.table"
Option Explicit
Option Private Module

Public Sub ToggleEmptyRows()
    With ActiveSheet
        If .Name <> STOWPLAN_SHEET_NAME Then
            Exit Sub
        End If
        
        OnStart
        Dim emptyCells As Range
        Dim cell As Range
        For Each cell In STOWAGE_PLAN_SHEEET.Range(TABLE_LEFT_COL & TABLE_TOP_ROW & ":" & TABLE_LEFT_COL & TABLE_BOTTOM_ROW)
            If cell.Value2 = vbNullString Then
                If emptyCells Is Nothing Then
                    Set emptyCells = cell
                Else
                    Set emptyCells = Union(emptyCells, cell)
                End If
            End If
        Next cell
        If Not emptyCells Is Nothing Then
            emptyCells.EntireRow.Hidden = Not emptyCells.EntireRow.Hidden
        End If
        OnEnd
    End With
End Sub

