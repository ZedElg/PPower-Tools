Attribute VB_Name = "SwapPositionsOfTwoObjects"
Sub SwapPositionsOfTwoObjects()
    Dim shp1 As shape
    Dim shp2 As shape
    Dim tempLeft As Single
    Dim tempTop As Single

    ' Check if there are exactly two shapes selected
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        If ActiveWindow.Selection.ShapeRange.Count = 2 Then
            Set shp1 = ActiveWindow.Selection.ShapeRange(1)
            Set shp2 = ActiveWindow.Selection.ShapeRange(2)

            ' Store the position of the first shape
            tempLeft = shp1.Left
            tempTop = shp1.Top

            ' Swap positions
            shp1.Left = shp2.Left
            shp1.Top = shp2.Top
            shp2.Left = tempLeft
            shp2.Top = tempTop

        Else
            MsgBox "Please select exactly two shapes.", vbExclamation
        End If
    Else
        MsgBox "Please select some shapes first.", vbExclamation
    End If
End Sub

