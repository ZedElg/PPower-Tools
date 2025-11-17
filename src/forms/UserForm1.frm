VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "List of Shortcuts"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7185
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim shortcuts As Variant
    Dim i As Integer

    ' Define the shortcuts: name, shortcut, and description
    shortcuts = Array( _
        Array("Hyperlink", "CTRL + K", "Insert hyperlink"), _
        Array("Duplicate", "CTRL + D", "Duplicate object"), _
        Array("New Slide", "CTRL + M", "Insert a new slide"), _
        Array("Copy Format", "CTRL + ALT + C", "Copy Format"), _
        Array("Paste Format", "CTRL + ALT + V", "Paste Format"), _
        Array("Text Box", "ALT + Q", "Insert a Text Box"), _
        Array("Copy", "Ctrl + C", "Copy the selected object or text"), _
        Array("Paste", "Ctrl + V", "Paste the copied object or text"), _
        Array("Undo", "Ctrl + Z", "Undo the last action"), _
        Array("Save", "Ctrl + S", "Save the presentation"), _
        Array("New Slide", "Ctrl + M", "Insert a new slide"))

    ' Add the headers
    lstShortcuts.AddItem "Name" & vbTab & "Shortcut" & vbTab & "Description"
    
    ' Add the shortcuts
    For i = LBound(shortcuts) To UBound(shortcuts)
        lstShortcuts.AddItem shortcuts(i)(0) & vbTab & shortcuts(i)(1) & vbTab & shortcuts(i)(2)
    Next i
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


