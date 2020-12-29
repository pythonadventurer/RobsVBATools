Attribute VB_Name = "modScreenRes"
Option Compare Database
'Source:
'https://stackoverflow.com/questions/11843310/how-to-retrieve-screen-size-resolution-in-ms-access-vba-to-re-size-a-form

Public Declare PtrSafe Function GetSystemMetrics32 Lib "user32" _
    Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long

Function ScreenRes() As Collection

Dim w As Long, h As Long
Dim hw As New Collection

w = GetSystemMetrics32(0) ' width
h = GetSystemMetrics32(1) ' height
hw.Add (w)
hw.Add (h)
     
Set ScreenRes = hw


End Function

Sub ScreenResExample()

Dim MyScreen As New Collection

Set MyScreen = ScreenRes

Debug.Print "Height: " & MyScreen(1)
Debug.Print "Width: " & MyScreen(2)

End Sub
