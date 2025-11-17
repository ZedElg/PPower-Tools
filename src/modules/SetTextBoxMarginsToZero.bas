Attribute VB_Name = "SetTextBoxMarginsToZero"
Sub SetTextBoxMarginsToZero()
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
                    ' Set all margins to zero
                    With shp.TextFrame
                        .MarginLeft = 0
                        .MarginRight = 0
                        .MarginTop = 0
                        .MarginBottom = 0
                    End With
                End If
            End If
        Next shp

    Else
        MsgBox "Please select some text boxes first.", vbExclamation
    End If
End Sub

