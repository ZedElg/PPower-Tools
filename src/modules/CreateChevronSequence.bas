Attribute VB_Name = "CreateChevronSequence"
Sub CreateChevronSequence()
    Dim osld As slide
    Dim i As Integer
    Dim numChevrons As Variant
    Dim shp As shape
    Dim xOffset As Single
    Dim yOffset As Single
    Dim chevronWidth As Single
    Dim chevronHeight As Single

    Set osld = Application.ActiveWindow.View.slide

    numChevrons = InputBox("How many chevrons do you want?", "Number of Chevrons", 3)

    ' Cancel check
    If numChevrons = vbNullString Then Exit Sub
    If Not IsNumeric(numChevrons) Then Exit Sub
    If numChevrons <= 0 Then Exit Sub

    numChevrons = CInt(numChevrons)

    chevronWidth = 100
    chevronHeight = 50
    xOffset = 100
    yOffset = 100

    ' First shape, flat beginning
    Set shp = osld.Shapes.AddShape(msoShapePentagon, xOffset, yOffset, chevronWidth, chevronHeight)
    shp.Fill.ForeColor.RGB = RGB(211, 211, 211)
    shp.Line.ForeColor.RGB = RGB(211, 211, 211)
    shp.Adjustments.Item(1) = 0.15

    ' Remaining chevrons
    For i = 2 To numChevrons
        Set shp = osld.Shapes.AddShape( _
            msoShapeChevron, _
            xOffset + (i - 1) * (chevronWidth + 10), _
            yOffset, _
            chevronWidth, _
            chevronHeight)

        shp.Fill.ForeColor.RGB = RGB(211, 211, 211)
        shp.Line.ForeColor.RGB = RGB(211, 211, 211)
        shp.Adjustments.Item(1) = 0.15
    Next i

End Sub


