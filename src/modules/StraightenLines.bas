Attribute VB_Name = "StraightenLines"
Sub StraightenLines()
    Dim sel As Selection
    Dim shp As shape
    
    ' Get the current selection
    Set sel = Application.ActiveWindow.Selection
    
    ' Check if the selection contains shapes
    If sel.Type = ppSelectionShapes Then
        ' Loop through each selected shape
        For Each shp In sel.ShapeRange
            ' Check if the shape is a line
            If shp.Type = msoLine Then
                ' Determine if the line should be horizontal or vertical
                If Abs(shp.Width) > Abs(shp.Height) Then
                    ' Make the line horizontal
                    shp.Height = 0
                Else
                    ' Make the line vertical
                    shp.Width = 0
                End If
            End If
        Next shp
    Else
        MsgBox "Please select lines to straighten.", vbExclamation
    End If
End Sub

