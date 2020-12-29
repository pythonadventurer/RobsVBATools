Attribute VB_Name = "modReadWriteText"
Option Compare Database
Option Explicit
Function ReadText(TextFile As String) As String
    
    'Read lines of text from a file.
    'Based on: http://thedbguy.blogspot.com/2016/05/how-to-read-text-file.html
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim FileStream As Object
    Set FileStream = fso.OpenTextFile(TextFile, 1) '1=ForReading, 8=ForAppending
    ReadText = FileStream.ReadAll
    Set FileStream = Nothing
    Set fso = Nothing

End Function


Sub WriteText(TextLines As Collection, TextFile As String)

    'Write lines of text to a file.
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim FileStream As Object
    Set FileStream = fso.CreateTextFile(TextFile)
    Dim Line As Variant
    For Each Line In TextLines
        FileStream.WriteLine Line
    
    Next Line
    
    FileStream.Close
    
End Sub
