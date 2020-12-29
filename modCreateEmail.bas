Attribute VB_Name = "modCreateEmail"
Option Compare Database
Option Explicit

Sub CreateEmail(Optional sRecipients As Variant, _
                Optional sSubject As String, _
                Optional sBody As String, _
                Optional sTemp As Variant, _
                Optional sBCC As Variant, _
                Optional cAttachments As Variant)
                
Dim oOutlook As Object
Dim oOutlookMsg As Object

'##################################################
'# Create and display an email message in Outlook #
'##################################################
'
'If an email template is specified using the argument sTemp, and that template includes an email subject,
'the template's subject will be used. If a subject was specified in sSubject, it will be ignored.
'
'User must at least pass an argument for sTemp.  If not, then must pass sRecipient, sSubject AND sBody.
'
'Specify multiple recipients for the argument sRecipients by passing a string with the email addresses
'separated by semicolons.
'Example: "rob@professionalco-op.com;dave@professionalco-op.com"
'

If IsMissing(sTemp) And (IsMissing(sRecipients) Or IsMissing(sSubject) Or IsMissing(sBody)) Then
    MsgBox "Either a recipient AND subject AND body must be specified, OR a template!"
    
    Exit Sub
    
End If

Set oOutlook = CreateObject("Outlook.application")

If Not IsMissing(sTemp) Then

'If a message template has been specified
    Set oOutlookMsg = oOutlook.CreateItemFromTemplate(sTemp)
    
    oOutlookMsg.Display
    
    'If the template includes a subject AND the user included a subject argument, the subject in
    'the template will be used, and the user-provided subject ignored.
    If IsNull(oOutlookMsg.Subject) Then
        If Not IsMissing(sSubject) Then
            oOutlookMsg.Subject = sSubject
        End If
    End If
        
Else
    Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
    
    oOutlookMsg.Display
    
    Signature = oOutlookMsg.HTMLbody

    oOutlookMsg.HTMLbody = sBody & Signature
    
    If Not IsMissing(sSubject) Then
        oOutlookMsg.Subject = sSubject

    End If
End If

If Not IsMissing(sRecipients) Then
    oOutlookMsg.Recipients.Add (sRecipients)
    
End If
    
If Not IsMissing(sBCC) Then
    oOutlookMsg.Recipients.Add (sBCC)
    
End If

' Add attachments to the message.
If Not IsMissing(cAttachments) Then
    If IsArray(cAttachments) Then
        For i = LBound(cAttachments) To UBound(cAttachments)
            If cAttachments(i) <> "" And cAttachments(i) <> "False" Then
                oOutlookMsg.Attachments.Add (cAttachments(i))
            End If
        Next i
    Else
        If cAttachments <> "" And cAttachments(i) <> "False" Then
            oOutlookMsg.Attachments.Add (cAttachments)
        End If
    End If
End If

Set oOutlook = Nothing

Set oOutlookMsg = Nothing

End Sub

