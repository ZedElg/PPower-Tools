Attribute VB_Name = "ReplaceTextInSelectedShapes"
Sub ReplaceTextInSelectedShapes()
    Dim shp As shape
    Dim selectedShapes As ShapeRange

    ' Check if there is any selection
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        Set selectedShapes = ActiveWindow.Selection.ShapeRange

        ' Loop through each selected shape
        For Each shp In selectedShapes
            ' Check if the shape has a text frame
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    ' Replace text with three dots
                    shp.TextFrame.textRange.Text = "..."
                End If
            End If
        Next shp

    Else
        MsgBox "Please select some text boxes first.", vbExclamation
    End If
End Sub

