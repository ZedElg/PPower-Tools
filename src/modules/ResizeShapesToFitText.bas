Attribute VB_Name = "ResizeShapesToFitText"
Sub ResizeShapesToFitText()
    Dim shp As shape
    Dim selectedShapes As ShapeRange

    ' Check if there are any shapes selected
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Set selectedShapes = ActiveWindow.Selection.ShapeRange
        
        ' Loop through each selected shape
        For Each shp In selectedShapes
            ' Check if the shape has a text frame with text
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    ' Resize the shape to fit the text
                    shp.TextFrame.AutoSize = ppAutoSizeShapeToFitText
                End If
            End If
        Next shp
        
    Else
        MsgBox "Please select some shapes first.", vbExclamation
    End If
End Sub

