Attribute VB_Name = "Feature_Box_Base"
'@Folder("StowagePlan.feature.drawings.box")
Public Function GenerateShape(Optional ByVal shapeTag As String = "") As Shape
    Dim cellLeftPsn     As Long:        cellLeftPsn = ActiveCell.Left - 50
    Dim cellTopPsn      As Long:        cellTopPsn = ActiveCell.Top - 15
    Dim ws              As Worksheet:   Set ws = ActiveSheet
    
    Set GenerateShape = ws.Shapes.AddShape(GetShapeBytTag(shapeTag), cellLeftPsn, cellTopPsn, 120, 40)
End Function

Public Function GetShapeBytTag(ByVal shapeTag As String) As Long
    Select Case shapeTag
        Case "msoShapeRectangularCallout"
            GetShapeBytTag = msoShapeRectangularCallout
        Case Else
            GetShapeBytTag = msoShapeRectangle
    End Select
End Function

