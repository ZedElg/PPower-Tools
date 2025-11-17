Attribute VB_Name = "GroupByColumns"
Sub GroupByColumns()
    Dim sel As Selection
    Dim selRange As ShapeRange
    Dim columns As Collection
    Dim shp As shape
    Dim columnKey As String
    Dim columnShapes As Collection
    Dim tolerance As Single
    Dim i As Long
    Dim j As Long

    ' Tolerance for considering shapes in the same column
    tolerance = 10 ' Adjust this value if necessary

    ' Get the current selection
    Set sel = Application.ActiveWindow.Selection

    ' Check if the selection contains shapes
    If sel.Type = ppSelectionShapes Then
        ' Get the shape range from the selection
        Set selRange = sel.ShapeRange

        ' Initialize the collection to hold column groupings
        Set columns = New Collection

        ' Loop through each shape in the selection
        For Each shp In selRange
            ' Round the Left position to the nearest multiple of tolerance
            columnKey = CStr(Int(shp.Left / tolerance) * tolerance)

            ' Check if there is already a collection for this column
            On Error Resume Next
            Set columnShapes = columns(columnKey)
            On Error GoTo 0

            ' If there is no collection for this column, create one
            If columnShapes Is Nothing Then
                Set columnShapes = New Collection
                columns.Add columnShapes, columnKey
            End If

            ' Add the shape to the column collection
            columnShapes.Add shp
            Set columnShapes = Nothing
        Next shp

        ' Group the shapes in each column
        For i = 1 To columns.Count
            Set columnShapes = columns(i)
            If columnShapes.Count > 1 Then
                Dim groupRange() As Variant
                ReDim groupRange(1 To columnShapes.Count)

                ' Fill the array with shape names
                For j = 1 To columnShapes.Count
                    groupRange(j) = columnShapes(j).Name
                Next j

                ' Group the shapes in this column
                selRange.Parent.Shapes.Range(groupRange).Group
            End If
        Next i
    Else
        MsgBox "Please select shapes to group by columns.", vbExclamation
    End If
End Sub

