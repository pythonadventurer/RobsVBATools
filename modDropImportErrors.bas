Attribute VB_Name = "modDropImportErrors"
Option Compare Database
Option Explicit

Sub DropImportErrors()
Dim tbl_name As DAO.TableDef, str As String
With CurrentDb
    For Each tbl_name In .TableDefs
            str = tbl_name.Name
            If InStr(str, "ImportErrors") <> 0 Then
            str = "DROP TABLE " & str & ""
            DoCmd.RunSQL str
            End If
    Next
End With
End Sub
