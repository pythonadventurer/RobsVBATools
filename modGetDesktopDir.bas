Attribute VB_Name = "modGetDesktopDir"
Option Compare Database
Option Explicit

Function GetDesktopDir() As String

'Return the current user's desktop directory
GetDesktopDir = "C:\Users\" & Environ("UserName") & "\Desktop"
  
End Function
