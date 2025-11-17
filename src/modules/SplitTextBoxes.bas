Attribute VB_Name = "SplitTextBoxes"
Sub SplitTextBoxes()
    Dim oshp As shape
    Dim osld As slide
    Dim paraCount As Long
    Dim i As Long, j As Long
    Dim dupRange As ShapeRange
    Dim newShp As shape
    Dim yPos As Single
    Dim txt As String

    ' Check selection
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Please select a single text box first.", vbExclamation
        Exit Sub
    End If
    
    If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
        MsgBox "Please select exactly one text box.", vbExclamation
        Exit Sub
    End If
    
    Set oshp = ActiveWindow.Selection.ShapeRange(1)
    Set osld = oshp.Parent
    
    If Not oshp.HasTextFrame Then
        MsgBox "The selected shape has no text frame.", vbExclamation
        Exit Sub
    End If
    
    If Not oshp.TextFrame2.HasText Then
        MsgBox "The selected text box is empty.", vbExclamation
        Exit Sub
    End If

    paraCount = oshp.TextFrame2.textRange.Paragraphs.Count
    
    ' If only one paragraph, leave it unchanged
    If paraCount <= 1 Then Exit Sub
    
    ' Start stacking at original top
    yPos = oshp.Top

    For i = 1 To paraCount
        ' Skip empty paragraphs
        If Len(Trim$(oshp.TextFrame2.textRange.Paragraphs(i).Text)) = 0 Then
            GoTo NextParagraph
        End If
        
        ' Duplicate original textbox
        Set dupRange = oshp.Duplicate
        Set newShp = dupRange(1)
        
        ' Keep original width and left, move to current y position
        newShp.Left = oshp.Left
        newShp.Top = yPos
        newShp.Width = oshp.Width
        
        ' Delete all paragraphs except i in the duplicate
        With newShp.TextFrame2.textRange
            Dim total As Long
            total = .Paragraphs.Count
            
            For j = total To 1 Step -1
                If j <> i Then
                    .Paragraphs(j).Delete
                End If
            Next j
            
            ' Remove trailing line break if present
            txt = .Text
            If Len(txt) > 0 Then
                If Right$(txt, 1) = vbCr Or Right$(txt, 1) = vbLf Then
                    .Characters(Len(txt), 1).Delete
                End If
            End If
        End With
        
        ' Let height adjust to text, keep width
        newShp.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        
        ' Next box goes directly under this one
        yPos = newShp.Top + newShp.Height

NextParagraph:
    Next i
    
    ' Remove original combined textbox
    oshp.Delete
End Sub

