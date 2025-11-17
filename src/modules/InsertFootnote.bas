Attribute VB_Name = "InsertFootnote"
Sub InsertFootnote()
    Dim slide As slide
    Dim footnoteBox As shape
    Dim textRange As textRange
    
    ' Define the dimensions of the text box
    Dim boxHeight As Single
    Dim boxWidth As Single
    Dim boxLeft As Single
    Dim boxTop As Single
    
    ' Set dimensions (convert cm to points)
    boxHeight = 0.34 * 28.3465 ' Convert cm to points
    boxWidth = 20.22 * 28.3465 ' Convert cm to points
    boxLeft = 1.54 * 28.3465 ' Convert cm to points
    boxTop = 18.06 * 28.3465 ' Convert cm to points
    
    ' Get the active slide
    Set slide = ActiveWindow.View.slide
    
    ' Create the text box
    Set footnoteBox = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, boxLeft, boxTop, boxWidth, boxHeight)
    
    ' Add text to the text box
    Set textRange = footnoteBox.TextFrame.textRange
    textRange.Text = "Source: ..."
    
    ' Format the text
    With textRange.Characters(1, 7).Font ' "Source:" part
        .Bold = msoTrue
    End With
    
    ' Set the font size for the entire text
    textRange.Font.Size = 8 ' Set the font size to 8
    
    ' Set the text box properties
    footnoteBox.Line.Visible = msoFalse ' No border
    footnoteBox.Fill.Visible = msoFalse ' No fill
    footnoteBox.TextFrame.AutoSize = ppAutoSizeShapeToFitText ' Resize the box to fit the text
    
    ' Set margins to zero
    With footnoteBox.TextFrame
        .MarginBottom = 0
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
    End With
    
    ' Set vertical anchor to middle
    footnoteBox.TextFrame.VerticalAnchor = msoAnchorMiddle
    
    ' Set horizontal alignment to left
    footnoteBox.TextFrame.textRange.ParagraphFormat.Alignment = ppAlignLeft
    
    ' Adjust the top position to ensure it is exactly 18.06 cm
    footnoteBox.Top = boxTop
End Sub

