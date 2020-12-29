Attribute VB_Name = "modSetAccessWindow"
Option Compare Database
Option Explicit
'Source:
'https://www.tek-tips.com/viewthread.cfm?qid=33710

Public Declare PtrSafe Function MoveWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Sub SetAccessWindow()
    Dim retval As Long
    
    'Screen width & height
    Dim sWidth As Long, sHeight As Long
    
    'Access window width & height
    Dim wWidth As Long, wHeight As Long
    
    'Access window x and y position
    Dim xPos As Long, yPos As Long
        
    'Desired width and height of Access application window
    wWidth = 640
    wHeight = 700
    
    'Adjust x and y position based on size of screen, so window will be centered.
    'Function GetSystemMetrics32 is in modScreenRes.
    sWidth = GetSystemMetrics32(0)
    sHeight = GetSystemMetrics32(1)
    
    xPos = sWidth / 2 - wWidth / 2
    yPos = sHeight / 2 - wHeight / 2
    
    'Size and center the Access window
    retval = MoveWindow(Application.hWndAccessApp, xPos, yPos, wWidth, wHeight, 1)
    
End Sub

