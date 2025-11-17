Attribute VB_Name = "SelectSimilarShapes"
Sub SelectSimilarShapes()
    Dim slide As slide
    Dim shape As shape
    Dim selectedShape As shape
    Dim shapeType As MsoShapeType
    Dim shapeColor As Long
    
    ' Check if a shape is selected
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Set selectedShape = ActiveWindow.Selection.ShapeRange(1)
        shapeType = selectedShape.Type
        
        ' Assuming you want to match the fill color
        shapeColor = selectedShape.Fill.ForeColor.RGB
        
        ' Loop through all shapes on the active slide
        Set slide = ActiveWindow.View.slide
        For Each shape In slide.Shapes
            If shape.Type = shapeType And shape.Fill.ForeColor.RGB = shapeColor Then
                shape.Select msoFalse ' Add shape to the selection
            End If
        Next shape
    Else
        MsgBox "Please select a shape first.", vbExclamation
    End If
End Sub

