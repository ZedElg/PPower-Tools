Attribute VB_Name = "ConvertShapesToTextBoxes"
Sub ConvertSelectedShapesToTextBoxes()
    Dim shp As shape
    Dim newShape As shape
    Dim selectedShapes As ShapeRange
    Dim slide As slide
    Dim shapeText As String
    Dim shapeLeft As Single
    Dim shapeTop As Single
    Dim shapeWidth As Single
    Dim shapeHeight As Single
    
    ' Check if there are any shapes selected
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Set selectedShapes = ActiveWindow.Selection.ShapeRange
        
        ' Loop through each selected shape
        For Each shp In selectedShapes
            ' Check if the shape is a shape (not already a text box or picture)
            If shp.Type = msoAutoShape Or shp.Type = msoFreeform Or shp.Type = msoLine Then
                ' Store the current shape's content and position
                shapeText = shp.TextFrame.textRange.Text
                shapeLeft = shp.Left
                shapeTop = shp.Top
                shapeWidth = shp.Width
                shapeHeight = shp.Height
                
                ' Delete the current shape
                shp.Delete
                
                ' Add a new text box with the same content and position
                Set newShape = ActiveWindow.Selection.SlideRange(1).Shapes.AddTextbox(msoTextOrientationHorizontal, shapeLeft, shapeTop, shapeWidth, shapeHeight)
                newShape.TextFrame.textRange.Text = shapeText
            End If
        Next shp
        
    Else
        MsgBox "Please select some shapes first.", vbExclamation
    End If
End Sub

