Attribute VB_Name = "modOpenLink"
Option Compare Database
Option Explicit

Sub OpenLink(Ctl As Control)
    'Open a hyperlink in a control
    If Ctl.IsHyperlink = True Then
        If IsNull(Ctl) = False Then
            Application.FollowHyperlink Ctl
        End If
    End If

End Sub
