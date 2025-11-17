Attribute VB_Name = "GroupByRows"
Sub GroupByRows()
    Dim sel As Selection
    Dim selRange As ShapeRange
    Dim rows As Collection
    Dim shp As shape
    Dim rowKey As String
    Dim rowShapes As Collection
    Dim tolerance As Single
    Dim i As Long
    Dim j As Long

    ' Tolerance for considering shapes in the same row
    tolerance = 10 ' Adjust this value if necessary

    ' Get the current selection
    Set sel = Application.ActiveWindow.Selection

    ' Check if the selection contains shapes
    If sel.Type = ppSelectionShapes Then
        ' Get the shape range from the selection
        Set selRange = sel.ShapeRange

        ' Initialize the collection to hold row groupings
        Set rows = New Collection

        ' Loop through each shape in the selection
        For Each shp In selRange
            ' Round the Top position to the nearest multiple of tolerance
            rowKey = CStr(Int(shp.Top / tolerance) * tolerance)

            ' Check if there is already a collection for this row
            On Error Resume Next
            Set rowShapes = rows(rowKey)
            On Error GoTo 0

            ' If there is no collection for this row, create one
            If rowShapes Is Nothing Then
                Set rowShapes = New Collection
                rows.Add rowShapes, rowKey
            End If

            ' Add the shape to the row collection
            rowShapes.Add shp
            Set rowShapes = Nothing
        Next shp

        ' Group the shapes in each row
        For i = 1 To rows.Count
            Set rowShapes = rows(i)
            If rowShapes.Count > 1 Then
                Dim groupRange() As Variant
                ReDim groupRange(1 To rowShapes.Count)

                ' Fill the array with shape names
                For j = 1 To rowShapes.Count
                    groupRange(j) = rowShapes(j).Name
                Next j

                ' Group the shapes in this row
                selRange.Parent.Shapes.Range(groupRange).Group
            End If
        Next i
    Else
        MsgBox "Please select shapes to group by rows.", vbExclamation
    End If
End Sub

