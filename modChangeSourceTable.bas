Attribute VB_Name = "modChangeSourceTable"
Option Compare Database
Option Explicit

Sub ChangeSourceTable(LinkName As String, SrcDb As String, SrcTable As String)

'Changes the source table of the selected link, while keeping the
'link name the same.

'LinkName : the name of the link, as it appears in the Navigation Pane.

'SrcDb : the name of the database that contains the table to be
'linked to.

'SrcTable : the name of the table in SrcDb that is to be linked to.

Dim tdfNew As TableDef

'Deletes ONLY the link, NOT the table itself.
DoCmd.RunSQL "DROP TABLE " & LinkName & ";"
Set tdfNew = CurrentDb.CreateTableDef(LinkName)
tdfNew.Connect = ";DATABASE=" & SrcDb
tdfNew.SourceTableName = SrcTable
CurrentDb.TableDefs.Append tdfNew

End Sub
