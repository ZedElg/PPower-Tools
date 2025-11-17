Attribute VB_Name = "MergeTextBoxes"
Option Explicit

Sub MergeTextBoxes()
    Dim sel As Selection
    Dim sr As ShapeRange
    Dim baseShp As shape
    Dim shp As shape
    Dim mergedText As String
    Dim i As Long

    ' Check selection
    Set sel = Application.ActiveWindow.Selection
    If sel Is Nothing Then
        MsgBox "Please select at least two text boxes.", vbExclamation
        Exit Sub
    End If

    If sel.Type <> ppSelectionShapes Then
        MsgBox "Please select at least two text boxes.", vbExclamation
        Exit Sub
    End If

    Set sr = sel.ShapeRange
    If sr.Count < 2 Then
        MsgBox "Please select at least two text boxes.", vbExclamation
        Exit Sub
    End If

    ' First selected shape is the base
    Set baseShp = sr(1)
    If Not baseShp.HasTextFrame Then
        MsgBox "The first selected shape has no text frame.", vbExclamation
        Exit Sub
    End If

    ' Build merged text from all shapes that have text
    mergedText = ""

    For i = 1 To sr.Count
        Set shp = sr(i)
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                If Len(mergedText) > 0 Then
                    mergedText = mergedText & vbCrLf
                End If
                mergedText = mergedText & shp.TextFrame.textRange.Text
            End If
        End If
    Next i

    If Len(mergedText) = 0 Then
        MsgBox "No text found in the selected shapes.", vbExclamation
        Exit Sub
    End If

    ' Put merged text into the first shape,
    ' it keeps its own formatting for all text
    baseShp.TextFrame.textRange.Text = mergedText

    ' Delete all other shapes from the selection
    For i = sr.Count To 1 Step -1
        If sr(i).Name <> baseShp.Name Then
            sr(i).Delete
        End If
    Next i

    ' Optional, auto size height to fit merged text
    baseShp.TextFrame.AutoSize = ppAutoSizeShapeToFitText

    ' Select the merged text box
    baseShp.Select
End Sub

