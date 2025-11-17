Attribute VB_Name = "CreateSticky"
Option Explicit

Const STICKY_BASE_LEFT As Single = 10
Const STICKY_BASE_TOP As Single = 10
Const STICKY_GAP As Single = 5

Sub CreateSticky()
    Dim sld As slide
    Dim stickyShape As shape
    Dim textRange As textRange
    Dim existingShape As shape
    
    Dim stickyWidth As Single
    Dim stickyHeight As Single
    
    Dim candidateLeft As Single
    Dim candidateTop As Single
    Dim isOverlap As Boolean
    Dim slot As Long
    
    Const MAX_SLOTS As Long = 50
    Const TOP_TOLERANCE As Single = 0.5
    
    Set sld = ActiveWindow.View.slide
    
    stickyWidth = 4 * 28.3465     ' 4 cm
    stickyHeight = 1.5 * 28.3465  ' 1.5 cm
    
    candidateTop = STICKY_BASE_TOP
    
    ' Find first free slot from left to right
    For slot = 0 To MAX_SLOTS
        candidateLeft = STICKY_BASE_LEFT + slot * (stickyWidth + STICKY_GAP)
        isOverlap = False
        
        For Each existingShape In sld.Shapes
            ' Only treat our sticky rectangles as blockers
            If existingShape.Type = msoShapeRectangle Then
                If existingShape.Name Like "Sticky_*" Then
                    ' Same row
                    If Abs(existingShape.Top - candidateTop) < TOP_TOLERANCE Then
                        ' Horizontal overlap
                        If existingShape.Left < candidateLeft + stickyWidth _
                           And existingShape.Left + existingShape.Width > candidateLeft Then
                            isOverlap = True
                            Exit For
                        End If
                    End If
                End If
            End If
        Next existingShape
        
        If Not isOverlap Then Exit For
    Next slot
    
    ' Create sticky at first free slot
    Set stickyShape = sld.Shapes.AddShape(msoShapeRectangle, _
                                          candidateLeft, candidateTop, _
                                          stickyWidth, stickyHeight)
    
    stickyShape.Name = "Sticky_" & sld.Shapes.Count
    
    stickyShape.Fill.ForeColor.RGB = RGB(255, 255, 0)
    stickyShape.Line.ForeColor.RGB = RGB(211, 211, 211)
    
    Set textRange = stickyShape.TextFrame.textRange
    textRange.Text = "..."
    textRange.Font.Color.RGB = RGB(0, 0, 0)
    textRange.Font.Size = 10
    stickyShape.TextFrame.textRange.ParagraphFormat.Alignment = ppAlignLeft
    stickyShape.TextFrame.VerticalAnchor = msoAnchorMiddle
End Sub



