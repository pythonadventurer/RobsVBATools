Attribute VB_Name = "modCopyFiles"
Option Compare Database
Option Explicit

Sub CopyFiles(src As String, dest As String)
'Copy all files from one folder to another

    Dim FSO As Object
    Dim fsoFiles As Variant
    Dim fil As Variant
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fils = FSO.GetFolder(src).Files

    For Each fil In fils
        FSO.CopyFile fil.Path, dest
        

    Next fil
    
End Sub
