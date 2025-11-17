Attribute VB_Name = "MakeShapesSameSize"
Sub MakeShapesSameSize()
    Dim sel As Selection
    Dim shp As shape
    Dim baseWidth As Single
    Dim baseHeight As Single
    
    ' Get the current selection
    Set sel = Application.ActiveWindow.Selection
    
    ' Check if the selection contains shapes
    If sel.Type = ppSelectionShapes Then
        ' Get the dimensions of the first selected shape
        baseWidth = sel.ShapeRange(1).Width
        baseHeight = sel.ShapeRange(1).Height
        
        ' Loop through each selected shape
        For Each shp In sel.ShapeRange
            ' Check if the shape is a horizontal or vertical line
            If shp.Type = msoLine Then
                ' Determine if the line is horizontal or vertical
                If Abs(shp.Width) > Abs(shp.Height) Then
                    ' Horizontal line
                    shp.Width = baseWidth
                    shp.Height = 0 ' Ensure it's horizontal
                Else
                    ' Vertical line
                    shp.Height = baseHeight
                    shp.Width = 0 ' Ensure it's vertical
                End If
            Else
                ' Apply both width and height for non-line shapes
                shp.Width = baseWidth
                shp.Height = baseHeight
            End If
        Next shp
    Else
        MsgBox "Please select shapes to apply dimensions.", vbExclamation
    End If
End Sub

