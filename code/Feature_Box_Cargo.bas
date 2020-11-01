Attribute VB_Name = "Feature_Box_Cargo"
'@Folder "StowagePlan.feature.drawings.box"
Option Explicit

Private lastLoadingPort     As String
Private destinationPort     As String
Private cargoBoxPresenter   As CargoBox

Public Sub AddStaticCargoBox(ByVal boxTag As String)
    Dim cell As Range
    For Each cell In LDG_PORTS_CODES_RANGE
        If cell.Value2 <> vbNullString Then
            lastLoadingPort = cell.Value2
        End If
    Next cell
    
    With ActiveCell
        For Each cell In DIS_PORTS_CODES_RANGE
            If cell.Interior.colorIndex = .Interior.colorIndex Then
                destinationPort = cell.Value2
                Exit For
            End If
        Next cell
    End With
    
    If lastLoadingPort = vbNullString Then
        MsgBox "Loading ports codes seems to be empty."
        Exit Sub
    End If
    
    If destinationPort = vbNullString Then
        MsgBox "Discharging port color not selected."
        Exit Sub
    End If

    Set cargoBoxPresenter = CreateCargoBox(boxTag)
    cargoBoxPresenter.Show
End Sub

Public Sub AddStaticCargoShape(ByVal boxTag As String)
    Dim cellLeftPsn     As Long:        cellLeftPsn = ActiveCell.Left - 50
    Dim cellTopPsn      As Long:        cellTopPsn = ActiveCell.Top - 15
    Dim cellColor       As Variant:     cellColor = ActiveCell.Interior.color
    Dim boxShape        As Shape:       Set boxShape = GenerateShape(boxTag)
    With boxShape
        .Name = "PKG_BOX_" & Format$(Now, "yyyymmddhhmmss") & PACKAGE_TAG
        With .Line
            .Weight = 0.5
            .ForeColor.RGB = RGB(0, 0, 0)
            .Visible = msoTrue
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = cellColor
            .Transparency = 0.1
        End With
        
        With .TextFrame2
            .AutoSize = msoAutoSizeShapeToFitText
            .WordWrap = True
            .HorizontalAnchor = msoAnchorCenter
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange
                .Characters.ParagraphFormat.Alignment = msoAlignCenter
                .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                .Characters.Font.Size = 12
                .Characters.Text = cargoBoxPresenter.TextBoxValue(destinationPort, lastLoadingPort, PACKING_PKGS)
            End With
        End With
    End With
End Sub



