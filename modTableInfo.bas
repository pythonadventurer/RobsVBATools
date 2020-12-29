Attribute VB_Name = "modTableInfo"
Option Compare Database

Sub ListTableFields(TableName As String)

Dim db As Database
Dim tdf As TableDef
Dim fld As Field

Set fs = CreateObject("Scripting.FileSystemObject")

Set txtFile = fs.CreateTextFile("C:\Users\robf\Documents\ListTables.txt", True)

Set db = CurrentDb

Set tdf = db.TableDefs(TableName)

txtFile.WriteLine (TableName)

Debug.Print TableName

For Each fld In tdf.Fields
    
    txtFile.WriteLine (fld.Name & "," & FieldType(fld.Type))
    
    Debug.Print fld.Name & "," & FieldType(fld.Type)
    
Next

txtFile.Close

Set tdf = Nothing
Set db = Nothing



End Sub

Function FieldType(IntegerType As Integer) As String

Select Case IntegerType
    Case dbBoolean
        FieldType = "Boolean"
    
    Case dbCurrency
        FieldType = "Currency"
        
    Case dbDate
        FieldType = "Date"
    
    Case dbInteger
        FieldType = "Integer"
        
    Case dbLong
        FieldType = "Long Integer"
        
    Case dbMemo
        FieldType = "Memo"
        
    Case dbNumeric
        FieldType = "Numeric"
        
    Case dbText
        FieldType = "Text"
        
End Select


End Function
