Attribute VB_Name = "modDocZim"
Option Compare Database
Option Explicit
'Change this constant as needed depending on where the documenation is to be created.
Const DocumentationDirectory As String = "C:\Users\robf\Documents\Working\Notebooks\DatabaseDocumentation"

'Create plain text documentation of the database, including:
'
'    * Control names, types and data sources
'    * Table fields and types
'    * SQL for all queries
'    * Code in forms and modules
'
'The documentation is created as a Zim Wiki notebook, which is searchable.
  
Function colDataControls() As Object
'List of controls that have the ControlSource property.
    Set colDataControls = New Collection
    colDataControls.Add "ListBox"
    colDataControls.Add "CheckBox"
    colDataControls.Add "ComboBox"
    colDataControls.Add "OptionButton"
    colDataControls.Add "OptionGroup"
    colDataControls.Add "TextBox"
    colDataControls.Add "Togglebutton"

End Function
Function ItemInCollection(varItem As Variant, varColl As Collection) As Boolean

'Determine if the given item, varItem, exists in collection varColl.
Dim varCheckItem As Variant
Dim Found As Boolean
Found = False
For Each varCheckItem In varColl
    If varCheckItem = varItem Then
        Found = True
    End If
Next varCheckItem
ItemInCollection = Found

End Function
Sub WriteText(TextLines As Collection, TextFile As String)

    'Write lines of text to a file.
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim FileStream As Object
    Set FileStream = FSO.CreateTextFile(TextFile)
    Dim Line As Variant
    For Each Line In TextLines
        FileStream.WriteLine Line
    
    Next Line
    
    FileStream.Close
    
End Sub
Sub CreateZimNotebook(strNotebookFolder As String, strNotebookName As String)

Dim NotebookText As New Collection

NotebookText.Add "[Notebook]"
NotebookText.Add "version=0.4"
NotebookText.Add "name=" & strNotebookName
NotebookText.Add "interwiki="
NotebookText.Add "home=Home"
NotebookText.Add "icon="
NotebookText.Add "document_root="
NotebookText.Add "shared=True"
NotebookText.Add "endofline=dos"
NotebookText.Add "disable_trash=False"
NotebookText.Add "profile="

Call WriteText(NotebookText, strNotebookFolder & "\notebook.zim")

End Sub


Function FieldTypeName(fld As DAO.Field)
    'Returns the type of the given field.
    Dim strReturn As String

    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work prior to Access 2007 and later.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn
End Function

Sub DocFormZim(strFormName As String, strOutputFile As String)

    Dim objForm As Form
    Dim objControl As Control
    Dim objModule As Module
    Dim colTxtLines As New Collection
    
    'Form must be open in Design view before it can be assigned to a variable
    'and its controls accessed.
    DoCmd.OpenForm strFormName, acDesign, , , , acHidden
    Set objForm = Forms(strFormName)

    'Create the heading for the form's documenation
    colTxtLines.Add "====== " & strFormName & " ======" & vbNewLine
    colTxtLines.Add "===== Recordsource ====="
    colTxtLines.Add "''" & Replace(objForm.Recordsource, vbNewLine, " ") & "''" & vbNewLine
    colTxtLines.Add "===== Controls ====="
    colTxtLines.Add "|Name|Type|Source|Caption|"
    
    Dim CtrlName As String
    Dim CtrlType As String
    Dim CtrlSource As String
    Dim CtrlCaption As String
    
    For Each objControl In objForm.Controls
        CtrlName = objControl.Name
        CtrlType = TypeName(objControl)
        
        'If the control is a type that has the ControlSource property,
        'include the control source (see function colDataControls, above)
        If ItemInCollection(TypeName(objControl), colDataControls) Then
            CtrlSource = objControl.ControlSource
            
        ElseIf TypeName(objControl) = "SubForm" Then
            CtrlSource = objControl.SourceObject
            
        Else
            CtrlSource = "None"
        End If
        
       'Include captions for label and button controls
        If TypeName(objControl) = "Label" Or TypeName(objControl) = "CommandButton" Then
            CtrlCaption = objControl.Caption
        
        Else
            CtrlCaption = "None"
            
        End If
        
        colTxtLines.Add "|" & CtrlName & "|" & CtrlType & "|" & CtrlSource & "|" & CtrlCaption & "|"

    Next objControl
    
    colTxtLines.Add vbNewLine
    
        'If the form has a module, list its code.
    If objForm.HasModule Then
        Set objModule = objForm.Module
        If objModule.CountOfLines > 0 Then
            colTxtLines.Add vbNewLine & "===== Code ====="
            Dim n As Integer
            For n = 1 To objModule.CountOfLines - 1
                colTxtLines.Add "''" & n & ": " & objModule.Lines(n, 1) & "''"
                
            Next n
        End If
    End If
    
    DoCmd.Close acForm, strFormName
     

    'Create the documentation file for the form
    Call WriteText(colTxtLines, strOutputFile)
        
    Set colTxtLines = Nothing
    
End Sub
Sub DocTableZim(strTableName As String, strOutputFile As String)

Dim db As DAO.Database
Dim tdfTableDef As TableDef
Dim fldField As Field
Dim colTxtLines As New Collection

Set db = CurrentDb

Set tdfTableDef = db.TableDefs(strTableName)

colTxtLines.Add "====== " & tdfTableDef.Name & " ======" & vbNewLine
colTxtLines.Add "===== Fields =====" & vbNewLine
colTxtLines.Add "|Name|Type|"
For Each fldField In tdfTableDef.Fields
    colTxtLines.Add "|" & fldField.Name & "|" & FieldTypeName(fldField) & "|"
Next fldField
    
colTxtLines.Add vbNewLine

Call WriteText(colTxtLines, strOutputFile)

Set db = Nothing
Set colTxtLines = Nothing

End Sub
Sub DocModuleZim(strModuleName As String, strOutputFile As String)

Dim modModule As Module
Dim colTxtLines As New Collection
Dim n As Integer

DoCmd.OpenModule strModuleName
Set modModule = Modules(strModuleName)
colTxtLines.Add "====== " & strModuleName & " =====" & vbNewLine
colTxtLines.Add "===== Code ====="

For n = 1 To modModule.CountOfLines
    colTxtLines.Add "''" & n & ": " & modModule.Lines(n, 1) & "''"

Next n

DoCmd.Close acModule, strModuleName

Call WriteText(colTxtLines, strOutputFile)

Set colTxtLines = Nothing

End Sub
Sub DocQueryZim(strQueryName As String, OutputFile As String)

Dim db As DAO.Database
Dim qryQueryDef As QueryDef
Dim colTxtLines As New Collection

Set db = CurrentDb
Set qryQueryDef = db.QueryDefs(strQueryName)

colTxtLines.Add "====== " & qryQueryDef.Name & "======" & vbNewLine
colTxtLines.Add "===== SQL ====="
colTxtLines.Add "'''"
colTxtLines.Add qryQueryDef.sql
colTxtLines.Add "'''"
Call WriteText(colTxtLines, OutputFile)

Set db = Nothing
Set colTxtLines = Nothing


End Sub
Sub DocReportZim(strReportName As String, strOutputFile As String)

Dim objReport As Report
Dim objControl As Control
Dim objModule As Module
Dim colTxtLines As New Collection

DoCmd.OpenReport strReportName, acDesign, , , acHidden
Set objReport = Reports(strReportName)

colTxtLines.Add "====== " & strReportName & " ======" & vbNewLine
colTxtLines.Add "===== Recordsource ====="
colTxtLines.Add "'''"
colTxtLines.Add objReport.Recordsource
colTxtLines.Add "'''" & vbNewLine
colTxtLines.Add "===== Controls ====="
colTxtLines.Add "|Name|Type|Source|Caption|"

Dim CtrlName As String
Dim CtrlType As String
Dim CtrlSource As String
Dim CtrlCaption As String

For Each objControl In objReport.Controls
    CtrlName = objControl.Name
    CtrlType = TypeName(objControl)
    
    'If the control is a type that has the ControlSource property,
    'include the control source (see function colDataControls, above)
    If ItemInCollection(TypeName(objControl), colDataControls) Then
        CtrlSource = objControl.ControlSource
        
    ElseIf TypeName(objControl) = "SubForm" Then
        CtrlSource = objControl.SourceObject
        
    Else
        CtrlSource = "None"
    End If
    
   'Include captions for label and button controls
    If TypeName(objControl) = "Label" Or TypeName(objControl) = "CommandButton" Then
        CtrlCaption = objControl.Caption
    
    Else
        CtrlCaption = "None"
        
    End If
    
    colTxtLines.Add "|" & CtrlName & "|" & CtrlType & "|" & CtrlSource & "|" & CtrlCaption & "|"

Next objControl

colTxtLines.Add vbNewLine

'If the report has a module, list its code.
If objReport.HasModule Then
    Set objModule = objReport.Module
    If objModule.CountOfLines > 0 Then
        colTxtLines.Add vbNewLine & "===== Code ====="
        Dim n As Integer
        For n = 1 To objModule.CountOfLines - 1
            colTxtLines.Add "''" & n & ": " & objModule.Lines(n, 1) & "''"
            
        Next n
    End If
End If

DoCmd.Close acReport, strReportName

'Create the documentation file for the form
Call WriteText(colTxtLines, strOutputFile)
    
Set colTxtLines = Nothing

End Sub
Sub DocZim()

Dim resp As Integer
resp = MsgBox("WARNING -- This will DELETE and RE-CREATE all existing documentation for this database in folder: " & vbNewLine & DocumentationDirectory & " !" & _
              vbNewLine & vbNewLine & "Are you sure?", vbYesNo, "Confirm")
              
If resp = vbNo Then
    Exit Sub

End If

Dim objFileSystem As Object
Set objFileSystem = CreateObject("Scripting.FileSystemObject")

'Get the name of the current Db, without directories,
'to use as folder name within DocFolder. For example, if the
'full database name  is:

'  R:\Working\Development\2020_dev.accdb

'then varDbName will be "2020_dev.accdb"

'Get the last part of the current database path by splitting
'the full path into array, then getting the highest indexed
'item in the array which is the last part of the database path.
Dim arrDbPath As Variant
Dim varDbName As String
arrDbPath = Split(CurrentDb.Name, "\")
varDbName = arrDbPath(UBound(arrDbPath))

'Snip off the *.accdb extension
varDbName = Left(varDbName, Len(varDbName) - 6)

'Create the folders that will contain the docs for each object type
Dim varFolders As New Collection
Dim varFolder As Variant
Dim varDbDocFolder As String
varFolders.Add ("Tables")
varFolders.Add ("Queries")
varFolders.Add ("Forms")
varFolders.Add ("Reports")
varFolders.Add ("Modules")
varDbDocFolder = objFileSystem.BuildPath(DocumentationDirectory, varDbName)
 
 
On Error GoTo ErrorHandler

objFileSystem.CreateFolder (varDbDocFolder)
 
For Each varFolder In varFolders
    objFileSystem.CreateFolder (objFileSystem.BuildPath(varDbDocFolder, varFolder))
    
    Dim varFileText As New Collection
    varFileText.Add "====== " & varFolder & " ======"
    Call WriteText(varFileText, objFileSystem.BuildPath(varDbDocFolder, varFolder & ".txt"))
    Set varFileText = Nothing
    
Next varFolder

Call CreateZimNotebook(varDbDocFolder, varDbName)

Dim varObject As Variant

'Create documentation for Forms
Debug.Print "Creating Forms documentation..."
For Each varObject In CurrentProject.AllForms

    Call DocFormZim(varObject.Name, objFileSystem.BuildPath(varDbDocFolder, "Forms\" & _
                                                           Replace(varObject.Name, " ", "_") & ".txt"))
    Debug.Print varObject.Name & " completed."
    
Next varObject

resp = MsgBox("Forms documentation complete. Continue?", vbYesNo, "Confirm")

If resp = vbNo Then
    Exit Sub
End If

'Create documentation for Modules
Debug.Print "Creating Modules documenation...."
For Each varObject In CurrentProject.AllModules
    Call DocModuleZim(varObject.Name, objFileSystem.BuildPath(varDbDocFolder, "Modules\" & _
                                                           Replace(varObject.Name, " ", "_") & ".txt"))
    Debug.Print varObject.Name & " completed."
Next varObject

resp = MsgBox("Modules documentation complete. Continue?", vbYesNo, "Confirm")

If resp = vbNo Then
    Exit Sub
End If

'Create documentation for Queries
Debug.Print "Creating Queries documentation..."
Dim varQueryDef As QueryDef

For Each varQueryDef In CurrentDb.QueryDefs
    'Exclude system queries
    If Left(varQueryDef.Name, 1) <> "~" Then
        Call DocQueryZim(varQueryDef.Name, objFileSystem.BuildPath(varDbDocFolder, "Queries\" & _
                                                           Replace(varQueryDef.Name, " ", "_") & ".txt"))
        Debug.Print varQueryDef.Name & " completed."
    End If
    
Next varQueryDef

resp = MsgBox("Queries documentation complete. Continue?", vbYesNo, "Confirm")

If resp = vbNo Then
    Exit Sub
End If

'Create documentation for Reports
Debug.Print "Creating Reports documentation...."

For Each varObject In CurrentProject.AllReports
    Call DocReportZim(varObject.Name, objFileSystem.BuildPath(varDbDocFolder, "Reports\" & _
                                                           Replace(varObject.Name, " ", "_") & ".txt"))
    Debug.Print varObject.Name & " completed."
    
Next varObject

resp = MsgBox("Reports documentation complete. Continue?", vbYesNo, "Confirm")

If resp = vbNo Then
    Exit Sub
End If

'Create documentation for Tables
Debug.Print "Creating Tables documentaton..."

Dim varTableDef As TableDef

For Each varTableDef In CurrentDb.TableDefs
    Call DocTableZim(varTableDef.Name, objFileSystem.BuildPath(varDbDocFolder, "Tables\" & _
                                                           Replace(varTableDef.Name, " ", "_") & ".txt"))
                                                           
    Debug.Print varTableDef.Name; " completed."
    
Next varTableDef

MsgBox ("Documentation of database '" & varDbName & "' completed.")

Exit Sub

ErrorHandler:
    If Err.Number = 58 Then 'Folder already exists
        objFileSystem.DeleteFolder (varDbDocFolder)
        
        Resume
        
    End If
    
End Sub



