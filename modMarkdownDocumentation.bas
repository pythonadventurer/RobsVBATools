Attribute VB_Name = "modMarkdownDocumentation"
Option Compare Database

'##########################
'# Markdown Documentation #
'##########################

'All code needed to create documentation for the database in a
'folder of Markdown documents and images.

Sub CreateMarkdownDocs(DestDir As String)
'Create a Markdown page for every object in the database.

Dim colTables As New Collection
Dim colQueries As New Collection
Dim colForms As New Collection
Dim colReports As New Collection
Dim colModules As New Collection

'# Tables
Dim Tdf As TableDef
For Each Tdf In CurrentDb.TableDefs
 'Exclude system and hidden tables
    If Left(Tdf.Name, 1) <> "~" And Left(Tdf.Name, 4) <> "MSys" Then
        Call DocumentTable(Tdf.Name, DestDir)
        Debug.Print "Processed table : " & Tdf.Name
        colTables.Add Item:=Tdf.Name
        
    End If
    
Next Tdf

'# Queries
Dim Qry As QueryDef
For Each Qry In CurrentDb.QueryDefs
    If Left(Qry.Name, 1) <> "~" Then
        Call DocumentQuery(Qry.Name, DestDir)
        Debug.Print "Processed query : " & Qry.Name
        colQueries.Add Item:=Qry.Name
        
    End If
    
Next Qry

'# Forms
Dim Frm As Variant

For Each Frm In CurrentProject.AllForms
    Call DocumentForm(Frm.Name, DestDir)
    Debug.Print "Processed form : " & Frm.Name
    colForms.Add Item:=Frm.Name
    
Next Frm

'# Reports
Dim Rpt As Variant

For Each Rpt In CurrentProject.AllReports
    Call DocumentReport(Rpt.Name, DestDir)
    Debug.Print "Processed report : " & Rpt.Name
    colReports.Add Item:=Rpt.Name
    
Next Rpt

'# Modules
Dim Mdl As Variant

For Each Mdl In CurrentProject.AllModules
    Call DocumentModule(Mdl.Name, DestDir)
    Debug.Print "Processed module : " & Mdl.Name
    colModules.Add Item:=Mdl.Name
    
    
Next Mdl

'# Index Page
Dim DatabaseName As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim PageText As String
Dim IndexPage As String
DatabaseName = FSO.GetFileName(CurrentDb.Name)


IndexPage = FSO.BuildPath(DestDir, "Object_Index.md")

PageText = "Created: " & Format(Now, "YYYY-MM-DD HH:MM:ss") & vbNewLine & vbNewLine

PageText = PageText & "# Tables" & vbNewLine
Dim Obj As Variant
For Each Obj In colTables
    PageText = PageText & "[[" & Obj & "]]" & vbNewLine
Next Obj
PageText = PageText & vbNewLine

PageText = PageText & "# Queries" & vbNewLine
For Each Obj In colQueries
    PageText = PageText & "[[" & Obj & "]]" & vbNewLine
Next Obj
PageText = PageText & vbNewLine

PageText = PageText & "# Forms" & vbNewLine
For Each Obj In colForms
    PageText = PageText & "[[" & Obj & "]]" & vbNewLine
Next Obj
PageText = PageText & vbNewLine

PageText = PageText & "# Reports" & vbNewLine
For Each Obj In colReports
    PageText = PageText & "[[" & Obj & "]]" & vbNewLine
Next Obj
PageText = PageText & vbNewLine

PageText = PageText & "# Modules" & vbNewLine
For Each Obj In colModules
    PageText = PageText & "[[" & Obj & "]]" & vbNewLine
Next Obj
PageText = PageText & vbNewLine


Dim PageObject As Object
Set PageObject = FSO.CreateTextFile(IndexPage)

PageObject.Write (PageText)
Debug.Print "Created index page: " & IndexPage

Dim LinkList As New Collection
Set LinkList = GetLinkList(IndexPage)

Dim objFolder As Object
Dim objFile As Object

Set objFolder = FSO.GetFolder(DestDir)
For Each objFile In objFolder.Files
    'Skip the index file!
    If InStr(objFile.Name, "_index.md") = 0 Then
        Call AddLinks(objFile.path, LinkList)
    End If
Next objFile

Debug.Print "Links created."




End Sub
Sub DocumentTable(TableName As String, DestDir As String)
'Create a Markdown page with the info about the table:
' - Fields
'   - Field Type
' - Connect string (if linked table)


Dim Db As DAO.Database
Set Db = CurrentDb
Dim Tdf As TableDef
Set Tdf = Db.TableDefs(TableName)

Dim PageText As String
PageText = ""

'## Creation Date
'Dim Created As String
'Created = Format(Now, "YYYY-MM-DD HH:MM:ss")
'PageText = PageText + "Created: " & Created & vbNewLine & vbNewLine
If Tdf.Connect <> "" Then
    PageText = PageText & "*Linked Table*" & vbNewLine
End If
PageText = PageText & vbNewLine & "## Fields" & vbNewLine & vbNewLine

'Create a Markdown table showing field names and types
Dim fld As Field
PageText = PageText & "Name | Type" & vbNewLine
PageText = PageText & "-|-" & vbNewLine

For Each fld In Tdf.Fields
    Dim FieldName As String
    FieldName = fld.Name
    
    'If there is a  hashtag in the field name, it will mess up the table
    'layout.  Change format to code-like.
    If InStr(FieldName, "#") <> 0 Then
        FieldName = "`" & FieldName & "`"
    End If
    
    PageText = PageText & FieldName & " | " & FieldTypeName(fld) & vbNewLine
Next fld

If Tdf.Connect <> "" Then
    PageText = PageText & vbNewLine & "## Table Link" & vbNewLine & Tdf.Connect
End If

Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim Page As String
Page = FSO.BuildPath(DestDir, TableName + ".md")

Dim PageObject As Object
Set PageObject = FSO.CreateTextFile(Page)
PageObject.Write (PageText)
Debug.Print "Created page: " & Page


End Sub

Sub DocumentForm(FormName As String, DestDir As String)
'Create a Markdown page in DestDir with the following info about the form:
' - Record source (name of query, or SQL)
' - Controls
'    - Control Name
'    - Control Type
'        - TextBox
'            - Control Source
'        - Label
'           - Caption
'        - ComboBox
'            - Control Source
'            - Row Source
'        - CheckBox
'            - Control Source
'        - CommandButton
'            - Caption
'            - OnClick code
'        - SubForm
'           - Source Object
'           - LinkMasterFields
'           - LinkChildFields
'- Code (if any)


Dim PageText As String
PageText = ""

'## Creation Date
Dim Created As String
Created = Format(Now, "YYYY-MM-DD HH:MM:ss")
PageText = PageText + "Created: " & Created & vbNewLine & vbNewLine

Dim Frm As Form
DoCmd.OpenForm FormName, acDesign
Set Frm = Forms(FormName)

'## RecordSource
Dim RecordSource As String
RecordSource = Frm.Properties("RecordSource")
If RecordSource = "" Then
    RecordSource = "Unbound"
End If
If InStr(RecordSource, "SELECT ") <> 0 Then
    'Recordsource = SQL, so format it as code.
    RecordSource = "```SQL" & vbNewLine & RecordSource & vbNewLine & "```"
End If

PageText = PageText & "## Recordsource" & vbNewLine & RecordSource & vbNewLine & vbNewLine

'## Controls
PageText = PageText & "## Controls" & vbNewLine

Dim Ctrl As Control
For Each Ctrl In Frm.Controls
    Dim ControlSource As String
    Dim RowSource As String
    PageText = PageText & vbNewLine & "### " & Ctrl.Name & vbNewLine
    
    'These are the only control types included in the documentation.
    If TypeName(Ctrl) = "TextBox" Or TypeName(Ctrl) = "Label" _
                              Or TypeName(Ctrl) = "ComboBox" _
                              Or TypeName(Ctrl) = "CheckBox" _
                              Or TypeName(Ctrl) = "CommandButton" _
                              Or TypeName(Ctrl) = "Subform" Then
        
        PageText = PageText & "- Type: " & TypeName(Ctrl) & vbNewLine
    End If
    Select Case TypeName(Ctrl)
        Case "TextBox"
            ControlSource = Ctrl.Properties("ControlSource")
            
            If ControlSource = "" Then
               ControlSource = "Unbound"
            
            ElseIf InStr(ControlSource, "SELECT ") <> 0 Then
               ControlSource = vbNewLine & "```SQL" & vbNewLine & ControlSource & vbNewLine & "```"
                
            End If
            
            PageText = PageText & "- ControlSource : " & ControlSource & vbNewLine
            
        Case "Label"
            PageText = PageText & "- Caption : " & Ctrl.Properties("Caption") & vbNewLine
        
        Case "ComboBox"
            ControlSource = Ctrl.Properties("ControlSource")
            
            If ControlSource = "" Then
               ControlSource = "Unbound"
            
            ElseIf InStr(ControlSource, "SELECT ") <> 0 Then
               ControlSource = vbNewLine & "```SQL" & vbNewLine & ControlSource & vbNewLine & "```"
                
            End If
            
            RowSource = Ctrl.Properties("RowSource")
            
            If RowSource = "" Then
               RowSource = "Unbound"
            
            ElseIf InStr(RowSource, "SELECT ") <> 0 Then
               RowSource = vbNewLine & "```SQL" & vbNewLine & RowSource & vbNewLine & "```"
                
            End If
                                 
            PageText = PageText & "- ControlSource : " & ControlSource & vbNewLine
            PageText = PageText & "- RowSource : " & RowSource & vbNewLine

         Case "CheckBox"
            If Ctrl.Properties("ControlSource") = "" Then
                 PageText = PageText & "- ControlSource : Unbound" & vbNewLine
            
             Else
                 PageText = PageText & "- ControlSource : " & Ctrl.Properties("ControlSource") & vbNewLine
            
             End If
        
        Case "SubForm"
            PageText = PageText & "- SourceObject : " & Ctrl.Properties("SourceObject") & vbNewLine
            PageText = PageText & "- LinkMasterFields : " & Ctrl.Properties("LinkMasterFields") & vbNewLine
            PageText = PageText & "- LinkChildFields : " & Ctrl.Properties("LinkChildFields") & vbNewLine

        Case "CommandButton"
            PageText = PageText & "- Caption : " & Ctrl.Properties("Caption") & vbNewLine
            
    End Select
            
            
Next Ctrl

'## Code
If Frm.HasModule = True Then
    Dim Mdl As Variant
    Set Mdl = Frm.Module
    
    Dim ModCode As String
    ModCode = "```VB" & vbNewLine

    Dim ModLines As Integer
    For ModLines = 1 To Mdl.CountOfLines
        ModCode = ModCode & Mdl.Lines(ModLines, 1) & vbNewLine
    Next ModLines

    ModCode = ModCode & vbNewLine & "```" & vbNewLine
    
    PageText = PageText & vbNewLine & "## Code" & vbNewLine
    PageText = PageText & ModCode & vbNewLine
    
End If

DoCmd.Close acForm, FormName

'Create the Markdown page and content
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim Page As String
Page = FSO.BuildPath(DestDir, FormName + ".md")

Dim PageObject As Object
Set PageObject = FSO.CreateTextFile(Page)
PageObject.Write (PageText)
Debug.Print "Created page: " & Page


    
End Sub

Sub DocumentQuery(QueryName As String, DestDir As String)

Dim Db As DAO.Database
Set Db = CurrentDb
Dim Qry As QueryDef
Set Qry = Db.QueryDefs(QueryName)

Dim PageText As String
PageText = ""

'## Creation Date
Dim Created As String
Created = Format(Now, "YYYY-MM-DD HH:MM:ss")
PageText = PageText & "Created: " & Created & vbNewLine & vbNewLine

'## SQL
PageText = PageText & "## SQL" & vbNewLine
PageText = PageText & "```SQL" & vbNewLine
PageText = PageText & Qry.SQL & vbNewLine
PageText = PageText & "```" & vbNewLine

'Create the Markdown page and content
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim Page As String
Page = FSO.BuildPath(DestDir, QueryName + ".md")

Dim PageObject As Object
Set PageObject = FSO.CreateTextFile(Page)
PageObject.Write (PageText)
Debug.Print "Created page: " & Page


End Sub

Sub DocumentModule(ModuleName As String, DestDir As String)

Dim PageText As String
PageText = ""

'## Creation Date
Dim Created As String
Created = Format(Now, "YYYY-MM-DD HH:MM:ss")
PageText = PageText & "Created: " & Created & vbNewLine & vbNewLine

Dim Mdl As Module
DoCmd.OpenModule (ModuleName)
Set Mdl = Modules(ModuleName)

Dim ModCode As String
ModCode = "```VB" & vbNewLine

Dim ModLines As Integer
For ModLines = 1 To Mdl.CountOfLines
    ModCode = ModCode & Mdl.Lines(ModLines, 1) & vbNewLine
Next ModLines

ModCode = ModCode & vbNewLine & "```" & vbNewLine

PageText = PageText & vbNewLine & "## Code" & vbNewLine
PageText = PageText & ModCode & vbNewLine

DoCmd.Close acModule, ModuleName

'Create the Markdown page and content
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim Page As String
Page = FSO.BuildPath(DestDir, ModuleName + ".md")

Dim PageObject As Object
Set PageObject = FSO.CreateTextFile(Page)
PageObject.Write (PageText)
Debug.Print "Created page: " & Page


End Sub
Sub DocumentReport(ReportName As String, DestDir As String)
Dim PageText As String
PageText = ""

'## Creation Date
Dim Created As String
Created = Format(Now, "YYYY-MM-DD HH:MM:ss")
PageText = PageText + "Created: " & Created & vbNewLine & vbNewLine

Dim Rpt As Report
DoCmd.OpenReport ReportName, acDesign
Set Rpt = Reports(ReportName)

'## RecordSource
Dim RecordSource As String
RecordSource = Rpt.Properties("RecordSource")
If RecordSource = "" Then
    RecordSource = "Unbound"
End If
PageText = PageText & "## Recordsource" & vbNewLine & RecordSource & vbNewLine & vbNewLine

'## Controls
PageText = PageText & "## Controls" & vbNewLine

Dim Ctrl As Control
For Each Ctrl In Rpt.Controls
    Dim ControlSource As String
    Dim RowSource As String
    PageText = PageText & vbNewLine & "### " & Ctrl.Name & vbNewLine
    
    'These are the only control types included in the documentation.
    If TypeName(Ctrl) = "TextBox" Or TypeName(Ctrl) = "Label" _
                              Or TypeName(Ctrl) = "ComboBox" _
                              Or TypeName(Ctrl) = "CheckBox" _
                              Or TypeName(Ctrl) = "CommandButton" _
                              Or TypeName(Ctrl) = "Subform" Then
        
        PageText = PageText & "- Type: " & TypeName(Ctrl) & vbNewLine
    End If
    Select Case TypeName(Ctrl)
        Case "TextBox"
            If Ctrl.Properties("ControlSource") = "" Then
                PageText = PageText & "- ControlSource : Unbound" & vbNewLine
                                       
            Else
                PageText = PageText & "- ControlSource : " & Ctrl.Properties("ControlSource") & vbNewLine
                
            End If

        Case "Label"
            PageText = PageText & "- Caption : " & Ctrl.Properties("Caption") & vbNewLine
        
        Case "ComboBox"
            ControlSource = Ctrl.Properties("ControlSource")
            
            If ControlSource = "" Then
               ControlSource = "Unbound"
            
            ElseIf InStr(ControlSource, "SELECT ") <> 0 Then
               ControlSource = vbNewLine & "```SQL" & vbNewLine & ControlSource & vbNewLine & "```"
                
            End If
            
            RowSource = Ctrl.Properties("RowSource")
            
            If RowSource = "" Then
               RowSource = "Unbound"
            
            ElseIf InStr(RowSource, "SELECT ") <> 0 Then
               RowSource = vbNewLine & "```SQL" & vbNewLine & RowSource & vbNewLine & "```"
                
            End If
                                 
            PageText = PageText & "- ControlSource : " & ControlSource & vbNewLine
            PageText = PageText & "- RowSource : " & RowSource & vbNewLine

         Case "CheckBox"
            If Ctrl.Properties("ControlSource") = "" Then
                 PageText = PageText & "- ControlSource : Unbound" & vbNewLine
            
             Else
                 PageText = PageText & "- ControlSource : " & Ctrl.Properties("ControlSource") & vbNewLine
            
             End If
        
        Case "SubForm"
            PageText = PageText & "- SourceObject : " & Ctrl.Properties("SourceObject") & vbNewLine
            PageText = PageText & "- LinkMasterFields : " & Ctrl.Properties("LinkMasterFields") & vbNewLine
            PageText = PageText & "- LinkChildFields : " & Ctrl.Properties("LinkChildFields") & vbNewLine

        Case "CommandButton"
            PageText = PageText & "- Caption : " & Ctrl.Properties("Caption") & vbNewLine
            
    End Select
            
            
Next Ctrl

'## Code
If Rpt.HasModule = True Then
    Dim Mdl As Variant
    Set Mdl = Rpt.Module
    
    Dim ModCode As String
    ModCode = "```VB" & vbNewLine

    Dim ModLines As Integer
    For ModLines = 1 To Mdl.CountOfLines
        ModCode = ModCode & Mdl.Lines(ModLines, 1) & vbNewLine
    Next ModLines

    ModCode = ModCode & vbNewLine & "```" & vbNewLine
    
    PageText = PageText & vbNewLine & "## Code" & vbNewLine
    PageText = PageText & ModCode & vbNewLine
    
End If

DoCmd.Close acReport, ReportName

'Create the Markdown page and content
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim Page As String
Page = FSO.BuildPath(DestDir, ReportName + ".md")

Dim PageObject As Object
Set PageObject = FSO.CreateTextFile(Page)
PageObject.Write (PageText)
Debug.Print "Created page: " & Page

End Sub

Function FieldTypeName(fld As DAO.Field) As String
'Purpose: Converts the numeric results of DAO Field.Type to text.
'Modified version of this:
'   http://allenbrowne.com/func-06.html

    Dim strReturn As String    'Name to return

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
        
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn
End Function

Function GetLinkList(LinkPage As String) As Collection
'Get a list of objects to create links in documents

'Get a list of objects to link to
Dim LinkList As New Collection
Dim FSO As Object
Dim TextStream As Object
Dim TextLine As String

Set FSO = CreateObject("Scripting.FileSystemObject")
Set TextStream = FSO.OpenTextFile(LinkPage)
Do While Not (TextStream.AtEndOfStream)
    TextLine = TextStream.ReadLine
    If Left(TextLine, 2) = "[[" Then
        LinkList.Add Item:=Mid(TextLine, 3, Len(TextLine) - 4)
        
    End If
Loop

TextStream.Close

Set GetLinkList = LinkList


End Function

Sub AddLinks(PageFile As String, Links As Collection)

'Read the text from the file to be searched for links
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim TextStream As Object
Dim FileText As String
Dim PageFileName As String
FileText = ""
Set TextStream = FSO.OpenTextFile(PageFile, ForReading)

'Get the name of the object from the page name so it can
'be excluded form the links.  Otherwise it would create a link
'to itself below.
PageFileName = FSO.GetFileName(PageFile)
PageFileName = Left(PageFileName, Len(PageFileName) - 3)

With TextStream
    Do Until .AtEndOfStream
        FileText = FileText & .ReadLine & vbNewLine
    Loop
End With
TextStream.Close

'Check the file contents for names of objects to be linked
Set TextStream = FSO.OpenTextFile(PageFile, ForAppending)
TextStream.WriteLine "# Related Objects"
Dim Link As Variant
For Each Link In Links
    If InStr(FileText, Link) <> 0 And Link <> PageFileName Then
        TextStream.WriteLine "[[" + Link + "]]"
    End If
Next Link

TextStream.Close
Set TextStream = Nothing


End Sub

