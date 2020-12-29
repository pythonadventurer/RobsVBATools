Attribute VB_Name = "modFileExists"
Option Compare Database
Option Explicit

Function FileExists(FilePath As String) As Boolean
'Check if a file exists.
If Dir(FilePath) = "" Then
    FileExists = False
Else
    FileExists = True
End If

End Function
