VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Number of minutes for timer
Const Interval As Integer = 1
'Module for setting a countdown timer.
'Designed to be attached to a form that has the following controls:
'btnStartTimer, btnStopTimer, btnResetTimer, lblTimeLeft, chkRun (invisible), txtSecondsLeft (invisible)

Private Sub btnResetTimer_Click()
    Call ResetTimer
        
End Sub

Private Sub btnStartTimer_Click()
'chkRun is an invisible check box controlled by btnStartTimer and btnStopTimer.
'On Timer event subtracts 1 second from the timer if chkRun = -1 (is checked)
    Me.chkRun = -1
    
End Sub

Private Sub btnStopTimer_Click()
    Me.chkRun = 0
    
End Sub

Private Sub Form_Load()
    
Call ResetTimer
    
End Sub

Private Sub Form_Timer()

'Check if the Run box is checked, if so, update the time.
If Me.chkRun = -1 Then
    If Me.txtSecondsLeft > 0 Then
        Me.txtSecondsLeft = Me.txtSecondsLeft - 1
        Me.lblTimeLeft.Caption = SecondsTimer(Me.txtSecondsLeft)
    Else
        Call ResetTimer
        
        Me.chkRun = -1
        
    End If
End If

End Sub

Sub ResetTimer()
    
'Reset back to the number of minutes set in constant Interval.
Me.chkRun = 0
Me.txtSecondsLeft = Interval * 60
Me.lblTimeLeft.Caption = SecondsTimer(Me.txtSecondsLeft)

End Sub
Function SecondsTimer(TotalSeconds As Integer) As String

Dim Hours As Integer
Dim Minutes As Integer
Dim Seconds As Integer

Hours = Int(TotalSeconds / 3600)
Minutes = Int((TotalSeconds - Hours * 3600) / 60)
Seconds = TotalSeconds - Hours * 3600 - Minutes * 60
SecondsTimer = Format(Hours, "00") & ":" & Format(Minutes, "00") & ":" & Format(Seconds, "00")


End Function

