Attribute VB_Name = "modWait"
Option Compare Database
Option Explicit
Declare Sub Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long)
Sub Wait(SlpSec As Long)

Sleep SlpSec * 1000 '1000 = 1 second

End Sub


