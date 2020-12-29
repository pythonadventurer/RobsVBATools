Attribute VB_Name = "modFolderExists"
Option Compare Database
Option Explicit

Function FolderExists(strPath As String) As Boolean
    'Check if the given folder exists.
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
    
End Function
