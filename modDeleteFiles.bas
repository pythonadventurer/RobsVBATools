Attribute VB_Name = "modDeleteFiles"
Option Compare Database
Option Explicit

Sub DeleteFiles(Folder, FileType As String)

'Recursive deletion of all files whose extention matches FileType in a directory tree.

'*Source:*
'**https://stackoverflow.com/questions/22645347/loop-through-all-subfolders-using-vba**

    Dim SubFolder
    
    For Each SubFolder In Folder.Subfolders
        DeleteFiles SubFolder, FileType
    
    Next
    
    Dim File
    
    For Each File In Folder.Files
        If Right(File.Name, 3) = FileType Then
            Debug.Print "DELETING: " & File
            
            Kill File
            
        End If
            
    Next
    
End Sub
