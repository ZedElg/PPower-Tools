Attribute VB_Name = "AlignShapesToCenterAndMiddle"
Sub AlignShapesToCenterAndMiddle()
    Dim sel As Selection
    Dim selRange As ShapeRange
    Dim minX As Single, minY As Single
    Dim maxX As Single, maxY As Single
    Dim centerX As Single, centerY As Single
    Dim shp As shape
    
    ' Get the current selection
    Set sel = Application.ActiveWindow.Selection
    
    ' Check if the selection contains shapes
    If sel.Type = ppSelectionShapes Then
        ' Get the shape range from the selection
        Set selRange = sel.ShapeRange
        
        ' Initialize min and max values
        minX = selRange(1).Left
        minY = selRange(1).Top
        maxX = selRange(1).Left + selRange(1).Width
        maxY = selRange(1).Top + selRange(1).Height
        
        ' Determine the bounding box of the selection
        For Each shp In selRange
            If shp.Left < minX Then minX = shp.Left
            If shp.Top < minY Then minY = shp.Top
            If shp.Left + shp.Width > maxX Then maxX = shp.Left + shp.Width
            If shp.Top + shp.Height > maxY Then maxY = shp.Top + shp.Height
        Next shp
        
        ' Calculate the center of the bounding box
        centerX = (minX + maxX) / 2
        centerY = (minY + maxY) / 2
        
        ' Align each shape to the center of the bounding box
        For Each shp In selRange
            shp.Left = centerX - (shp.Width / 2)
            shp.Top = centerY - (shp.Height / 2)
        Next shp
    Else
        MsgBox "Please select shapes to align.", vbExclamation
    End If
End Sub

