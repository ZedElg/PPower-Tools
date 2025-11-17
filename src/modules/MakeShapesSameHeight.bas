Attribute VB_Name = "MakeShapesSameHeight"
Sub MakeShapesSameHeight()
    Dim sel As Selection
    Dim shp As shape
    Dim baseHeight As Single
    
    ' Get the current selection
    Set sel = Application.ActiveWindow.Selection
    
    ' Check if the selection contains shapes
    If sel.Type = ppSelectionShapes Then
        ' Get the height of the first selected shape
        baseHeight = sel.ShapeRange(1).Height
        
        ' Loop through each selected shape and apply the height
        For Each shp In sel.ShapeRange
            shp.Height = baseHeight
        Next shp
    Else
        MsgBox "Please select shapes to make the same height.", vbExclamation
    End If
End Sub

