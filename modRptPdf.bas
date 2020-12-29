Attribute VB_Name = "modRptPdf"
Option Compare Database
Option Explicit

Sub rptPDF(rptSrcName As String, rptDestName As String, rptDestPath As String)

'Creates a PDF of a report in the specified location
    
'    rptSrcName:  Name of the report to output to pDF
'    rptDestPath: Location where the PDF is to be created
'    rptDestName: Filename for the PDF

DoCmd.OutputTo acOutputReport, rptSrcName, acFormatPDF, _
               rptDestPath & "\" & rptDestName & ".pdf"

End Sub

