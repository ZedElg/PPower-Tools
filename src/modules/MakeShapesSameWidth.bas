Attribute VB_Name = "MakeShapesSameWidth"
Sub MakeShapesSameWidth()
    Dim sel As Selection
    Dim shp As shape
    Dim baseWidth As Single
    
    ' Get the current selection
    Set sel = Application.ActiveWindow.Selection
    
    ' Check if the selection contains shapes
    If sel.Type = ppSelectionShapes Then
        ' Get the width of the first selected shape
        baseWidth = sel.ShapeRange(1).Width
        
        ' Loop through each selected shape and apply the width
        For Each shp In sel.ShapeRange
            shp.Width = baseWidth
        Next shp
    Else
        MsgBox "Please select shapes to make the same width.", vbExclamation
    End If
End Sub

