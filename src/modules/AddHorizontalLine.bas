Attribute VB_Name = "AddHorizontalLine"
Sub AddHorizontalLine()
    Dim slide As slide
    Dim lineShape As shape
    
    ' Define the dimensions and position of the line
    Dim lineLength As Single
    Dim lineThickness As Single
    Dim lineLeft As Single
    Dim lineTop As Single
    
    lineLength = 4 * 28.3465 ' Convert cm to points
    lineThickness = 0.5 ' Line thickness in points
    lineLeft = 100 ' Left position (adjust as needed)
    lineTop = 100 ' Top position (adjust as needed)
    
    ' Get the active slide
    Set slide = ActiveWindow.View.slide
    
    ' Create a horizontal line
    Set lineShape = slide.Shapes.AddLine(lineLeft, lineTop, lineLeft + lineLength, lineTop)
    
    ' Set the line color to black
    lineShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    
    ' Set the line thickness to 0.5pt
    lineShape.Line.Weight = lineThickness
End Sub
