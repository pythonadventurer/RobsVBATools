Attribute VB_Name = "modChangeLog"
Option Compare Database
Option Explicit

Sub UpdateChangelog(Frm As Form, RecordID As Variant, Optional strMode As String)

'Track changes to data.
'
'To include a form's control in the changelog, set its Tag property to 'Changelog'.
'Do this for all controls in the form that are to have changes tracked.
'
'
'Call this sub from a form's AfterUpdate event. Example:
'
'    Private Sub Form_BeforeUpdate(Cancel As Integer)
'        Call UpdateChangelog(Me, Me.Company)
'    End Sub
'
'Recordid = the table's primary key, or another field that
'will ensure the user can identify which record was changed.
'

'To include a note in the Notes field of the changelog, include
'VBA code in the controls BeforeUpdate event
'to add a comma and the note to the tag.  Example :
'
'    Private Sub AssistantsName_BeforeUpdate(Cancel As Integer)
'         Dim StrName As String
'         StrName = InputBox("Who requested the change?")
'         Ctrl.Tag = Ctrl.Tag & ",Change requested by " & StrName
'    End Sub
'
'The presence of the comma will prompt the UpdateChangelog sub to
'to treat the tag as an array, with element 1 as the note.


'Structure of the required table tblChangeLog:

'|Name        |Type|
'|LogID       |AutoNumber|
'|DateTime    |Date/Time|
'|UserID      |Long Integer|
'|FormName    |Text|
'|ControlName |Text|
'|OldValue    |Text|
'|NewValue    |Text|
'|RecordID    |Text|
'|Notes       |Text|


Dim Ctl As Control
Dim varBefore As Variant
Dim varAfter As Variant
Dim strControlName As String
Dim strSQL As String
Dim UserID As Long
UserID = DLookup("user_id", "user")


On Error GoTo errHandler

For Each Ctl In Frm.Controls
    With Ctl
        'Only form controls tagged "Changelog" will be tracked in the changelog.
        If .Tag Like "Changelog*" Then
                       
            'There is a note attached, so split the tag into an array using the comma
            If InStr(.Tag, ",") > 0 Then
                Dim TagContent() As String
                Dim Notes As String
                TagContent = Split(.Tag, ",")
                
                'TagContent(0) = "Changelog", TagContent(1) is the note.
                Notes = TagContent(1)
          
            'No note attached
            Else
                Notes = ""
         
            End If
         
            strControlName = .Name
            
            'Change nulls to empty strings so the expression "varAfter <> varBefore", below,
            'will evaluate to either True or False, and never Null.
            varBefore = Nz(.OldValue, "")
            varAfter = Nz(.Value, "")
            
            'Convert vars to strings
            varBefore = CStr(varBefore)
            varAfter = CStr(varAfter)
            
            'Escape any apostrophes present..
            varBefore = Replace(varBefore, "'", "''")
            varAfter = Replace(varAfter, "'", "''")
            
            
            'Check for strMode = "Delete". In the case of deletions, varBefore will equal varAfter.
            'Setting varAfter to "" will make them unequal, and trigger the changelog update.
            If Not (IsMissing(strMode)) Then
                If strMode = "Delete" Then
                    varAfter = ""
                    Notes = "Deleted"
                
                End If
            End If
            
            Notes = Replace(Notes, "'", "''")
            
            'Track only fields that have changed.
            If varAfter <> varBefore Then
                strSQL = "INSERT INTO tblChangelog (UserID, " & _
                                                 "FormName, " & _
                                                 "ControlName, " & _
                                                 "OldValue, " & _
                                                 "NewValue, " & _
                                                 "RecordID, " & _
                                                 "Notes) " _
                                    & "VALUES(" & UserID & ", '" & _
                                                  Frm.Name & " ', '" & _
                                                  strControlName & "', '" & _
                                                  varBefore & "', '" & _
                                                  varAfter & "', '" & _
                                                  RecordID & "', '" & _
                                                  Notes & "');"
                                
                DoCmd.SetWarnings False
                DoCmd.RunSQL strSQL
                DoCmd.SetWarnings True
            End If
        End If
    End With
Next

Set Ctl = Nothing

Exit Sub

errHandler:
  MsgBox "Error : " & Err.Number & vbNewLine & _
                      Err.Description & vbNewLine & _
         "Procedure : UpdateChangelog", vbOKOnly, "Error"
End Sub

