Attribute VB_Name = "Feature_Box_Info"
'@Folder "StowagePlan.feature.drawings.box"
Option Explicit

Private destinationPort     As String
Private infoBoxPresenter    As InfoBox

Public Sub AddInfoBox()
    Dim cell As Range
    With ActiveCell
        For Each cell In DIS_PORTS_CODES_RANGE
            If cell.Interior.colorIndex = .Interior.colorIndex Then
                destinationPort = cell.Value2
                Exit For
            End If
        Next cell
    End With
    
    If destinationPort = vbNullString Then
        MsgBox "Discharging port color not selected."
        Exit Sub
    End If
    
    Set infoBoxPresenter = CreateInfoBox
    infoBoxPresenter.Show
End Sub

Public Sub InfoBoxShape()
    Dim cellLeftPsn     As Long:        cellLeftPsn = ActiveCell.Left - 50
    Dim cellTopPsn      As Long:        cellTopPsn = ActiveCell.Top - 15
    Dim ws              As Worksheet:   Set ws = ActiveSheet
    Dim box             As Shape:       Set box = GenerateShape
    Dim cellColor       As Variant:     cellColor = ActiveCell.Interior.color
    
    With box
        .Name = Format$(Now, "yyyymmddhhmmss") & "_" & destinationPort & INFO_BOX_TAG
        With .Line
            .Weight = 0.5
            .ForeColor.RGB = RGB(0, 0, 0)
            .Visible = msoFalse
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = cellColor
            .Transparency = 1
        End With
        With .TextFrame2
            .AutoSize = msoAutoSizeShapeToFitText
            .WordWrap = True
            .HorizontalAnchor = msoAnchorCenter
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange
                .Characters.ParagraphFormat.Alignment = msoAlignCenter
                .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                .Characters.Font.Size = 14
                .Characters.Text = infoBoxPresenter.TextValue
            End With
        End With
    End With
End Sub
