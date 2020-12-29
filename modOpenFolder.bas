Attribute VB_Name = "modOpenFolder"
Option Compare Database
Option Explicit


Public Sub OpenFolder(fPath As String)

If FolderExists(fPath) = True Then
    Shell "C:\WINDOWS\explorer.exe """ & fPath & "", vbNormalFocus
Else
    MsgBox "Folder: " & vbNewLine & fPath & " does not exist."
    
End If


End Sub
