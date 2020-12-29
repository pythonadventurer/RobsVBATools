Attribute VB_Name = "modUpdateTableLinks"
Option Compare Database
Option Explicit

Sub UpdateTableLinks(dbpath As String, currPath As String, newPath As String)

Dim db As DAO.Database

Set db = OpenDatabase(dbpath)

Dim tdf As TableDef

For Each tdf In db.TableDefs
    If InStr(tdf.Connect, currPath) > 0 And InStr(tdf.Name, "~") = 0 Then
        tdf.Connect = Replace(tdf.Connect, currPath, newPath)

        tdf.RefreshLink

        Forms!frmMain.txtLog = Forms!frmMain.txtLog & "Updated linked table '" & tdf.Name & "' to: " & vbNewLine & tdf.Connect & vbNewLine
    
    End If

Next tdf

Set db = Nothing

End Sub
