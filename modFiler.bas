Attribute VB_Name = "modFiler"
Option Compare Database
Option Explicit

Function Filer(fdType As Integer)

'Open a File Dialog to enable user to select a file or folder.
'   fdType = 2 : Save As
'   fdType = 3 : Select File
'   fdType = 4 : Select Folder

Dim fDialog As Office.FileDialog

Dim varItem As Variant

Dim dTitle As String

If fdType >= 2 And fdType <= 4 Then
    Set fDialog = Application.FileDialog(fdType)
    
    If fdType = 2 Then
        dTitle = "Save File As"
    
    ElseIf fdType = 3 Then
        dTitle = "Select File"
    
    ElseIf fdType = 4 Then
        dTitle = "Select Folder"
    
    Else
        dTitle = ""
        
    End If
    
    With fDialog
        .Title = dTitle
        
        If .Show = False Then
            Exit Function
            
        End If
        
        Filer = .SelectedItems(1)
    
    End With

End If

End Function
