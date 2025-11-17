Attribute VB_Name = "ToggleHarveyBall"
Sub ToggleHarveyBall()
    Dim sel As Selection
    Dim grpShape As shape
    Dim quarterCircle As shape
    Dim shp As shape
    Dim isHarveyBall As Boolean
    
    ' Initialize the Harvey Ball flag
    isHarveyBall = False
    
    ' Get the current selection
    On Error Resume Next
    Set sel = Application.ActiveWindow.Selection
    On Error GoTo 0
    
    ' Check if any selection exists
    If sel Is Nothing Then
        MsgBox "No Harvey Ball selected.", vbExclamation
        Exit Sub
    End If
    
    ' Check if the selection contains shapes and a single shape is selected
    If sel.Type = ppSelectionNone Or sel.Type = ppSelectionSlides Then
        MsgBox "No Harvey Ball selected.", vbExclamation
        Exit Sub
    End If

    If sel.Type = ppSelectionShapes And sel.ShapeRange.Count = 1 Then
        ' Get the selected shape
        Set grpShape = sel.ShapeRange(1)
        
        ' Check if the shape is a group
        If grpShape.Type = msoGroup Then
            ' Loop through the shapes in the group to find the quarter circle
            For Each shp In grpShape.groupItems
                If shp.AutoShapeType = msoShapePie Then
                    Set quarterCircle = shp
                    isHarveyBall = True ' Set the flag to true if a quarter circle is found
                    Exit For
                End If
            Next shp
            
            ' If the quarter circle is found, adjust its angle
            If isHarveyBall Then
                ' Add 90 degrees to the start angle
                quarterCircle.Adjustments.Item(2) = quarterCircle.Adjustments.Item(2) + 90
                
                ' Ensure the angle stays within 0-360 range
                If quarterCircle.Adjustments.Item(2) >= 360 Then
                    quarterCircle.Adjustments.Item(2) = quarterCircle.Adjustments.Item(2) - 360
                End If
            Else
                MsgBox "No quarter circle found in the selected group.", vbExclamation
            End If
        Else
            MsgBox "Please select a grouped shape containing a quarter circle.", vbExclamation
        End If
    Else
        MsgBox "No Harvey Ball selected.", vbExclamation
    End If
End Sub

