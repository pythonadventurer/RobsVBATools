Attribute VB_Name = "modFolderExists"
Option Compare Database
Option Explicit

Function FolderExists(strPath As String) As Boolean
    'Check if the given folder exists.
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    
    If Dir(strPath, vbDirectory) = "." Then
        FolderExists = True
    Else
        FolderExists = False
    End If
    
    
End Function
