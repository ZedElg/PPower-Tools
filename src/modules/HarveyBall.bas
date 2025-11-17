Attribute VB_Name = "HarveyBall"
Sub HarveyBall()
    Dim osld As slide
    Dim ovalShape As shape
    Dim quarterCircle As shape
    Dim groupItems As ShapeRange
    Dim grpShape As shape

    ' Set the slide to the active slide
    Set osld = Application.ActiveWindow.View.slide
    
    ' Add an oval shape
    Set ovalShape = osld.Shapes.AddShape(msoShapeOval, 100, 100, 28.35, 28.35) ' 1cm x 1cm is approximately 28.35 points
    ovalShape.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White color
    ovalShape.Line.ForeColor.RGB = RGB(255, 255, 255) ' Outline same color as fill
    ovalShape.ZOrder msoSendToBack ' Send to back

    ' Add a quarter circle shape
    Set quarterCircle = osld.Shapes.AddShape(msoShapePie, 100, 100, 28.35, 28.35)
    quarterCircle.Fill.ForeColor.RGB = RGB(211, 211, 211) ' Light grey color
    quarterCircle.Line.Visible = msoFalse ' Hide the outline
    quarterCircle.Adjustments.Item(1) = 270
    quarterCircle.Adjustments.Item(2) = 0

    ' Group the shapes together
    Set groupItems = osld.Shapes.Range(Array(ovalShape.Name, quarterCircle.Name))
    Set grpShape = groupItems.Group
    
    ' Center the group on the slide
    grpShape.Left = (osld.Master.Width - grpShape.Width) / 2
    grpShape.Top = (osld.Master.Height - grpShape.Height) / 2
End Sub

