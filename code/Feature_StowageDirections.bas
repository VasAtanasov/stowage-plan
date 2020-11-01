Attribute VB_Name = "Feature_StowageDirections"
'@Folder "StowagePlan.feature.drawings.directions"
Option Explicit

Public Sub StowDirectionCreate(ByVal rtn As Variant)
    Dim ws As Worksheet
    Dim shpFF As FreeformBuilder
    Dim shp As Shape
    Dim cellTop As Variant
    Dim cellLeft As Variant
    Dim cell As Range
    Dim destinationPort As String
    
    With ActiveCell
        For Each cell In DIS_PORTS_CODES_RANGE
            If cell.Interior.colorIndex = .Interior.colorIndex Then
                destinationPort = cell.Value2
                Exit For
            End If
        Next cell
    End With
    
    If destinationPort = vbNullString Then
        MsgBox "There are no discharging ports codes."
        Exit Sub
    End If
    
    cellTop = ActiveCell.Top
    cellLeft = ActiveCell.Left
    Set ws = ActiveSheet
    Set shpFF = ws.Shapes.BuildFreeform(msoEditingAuto, cellLeft, cellTop)

    With shpFF
        .AddNodes msoSegmentLine, msoEditingAuto, cellLeft + 100, cellTop
        .AddNodes msoSegmentLine, msoEditingAuto, cellLeft + 75, cellTop - 10
        Set shp = .ConvertToShape
        'Set rotation on creation
        shp.Rotation = rtn
    End With
    With shp
        .Name = GetCreationTime(destinationPort)
        With .Line
            .Weight = 0.5
            .ForeColor.RGB = RGB(0, 0, 0)
            .Visible = msoTrue
        End With
    End With
End Sub

Private Function GetCreationTime(ByVal PortCode As String) As String
    GetCreationTime = STOW_DORECTION_TAG & "_" & Format$(Now, "yyyymmddhhmmss") & "_" & PortCode
End Function

