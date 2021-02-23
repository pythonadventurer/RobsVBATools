Attribute VB_Name = "modLogging"
Option Compare Database
Option Explicit

Sub SendToLog(LogFile As String, LogText As String)
'Writes text LogText to file LogFile.
'Automatically adds the current date and time to the log entry.
'If LogFile already exists, LogText will be appended to it.
'If LogFile doesn't exist, it will be created.

Dim LogEntry As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim FileNum As Integer

LogEntry = Format(CStr(Now), "yyyy-mm-dd hh:mm:ss") & " " & LogText

LogFile = FSO.BuildPath(funcAppLog, LogFile)

FileNum = FreeFile
    
If IsNull(Dir(LogFile)) Then
    Open LogFile For Output As FileNum
    
Else
    Open LogFile For Append As FileNum
    

End If

Print #FileNum, LogEntry
    
Close FileNum

End Sub

