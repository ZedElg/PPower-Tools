Attribute VB_Name = "SelectShapesWithSameSize"
Sub SelectShapesWithSameSize()
    Dim sel As Selection
    Dim baseShape As shape
    Dim shp As shape
    Dim sl As slide
    Dim sameSizeShapes As Collection
    Dim i As Long

    ' Get the current selection
    On Error Resume Next
    Set sel = Application.ActiveWindow.Selection
    On Error GoTo 0

    ' Check if the selection is valid
    If sel Is Nothing Then
        MsgBox "No object is selected. Please select a single shape.", vbExclamation
        Exit Sub
    ElseIf sel.Type = ppSelectionNone Then
        MsgBox "No object is selected. Please select a single shape.", vbExclamation
        Exit Sub
    ElseIf sel.Type <> ppSelectionShapes Or sel.ShapeRange.Count <> 1 Then
        MsgBox "Please select a single shape to find others with the same size.", vbExclamation
        Exit Sub
    End If

    ' Get the selected shape
    Set baseShape = sel.ShapeRange(1)

    ' Get the slide the shape is on
    Set sl = baseShape.Parent

    ' Initialize the collection to hold shapes of the same size
    Set sameSizeShapes = New Collection

    ' Loop through each shape on the slide
    For Each shp In sl.Shapes
        ' Check if the shape has the same width and height as the selected shape
        If shp.Width = baseShape.Width And shp.Height = baseShape.Height Then
            sameSizeShapes.Add shp
        End If
    Next shp

    ' Deselect all shapes first
    Application.ActiveWindow.Selection.Unselect

    ' Select all shapes that have the same size
    For i = 1 To sameSizeShapes.Count
        sameSizeShapes(i).Select msoFalse
    Next i
End Sub

