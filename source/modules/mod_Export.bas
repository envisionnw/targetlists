Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Export
' Description:  Export functions and routines
'
' Source/date:  Bonnie Campbell, 4/2/2015
' Revisions:    BLC - 4/2/2015 - initial version
' =================================

' ---------------------------------
' SUB:          OutputPDF
' Description:  Exports a report as a PDF
' Assumptions:  -
' Parameters:   rpt        - Report being converted to PDF (report)
'               strPDFName - desired name for PDF report (string)
'               strPath    - directory path location to save PDF report (string)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Bmo, December 5, 2013
' http://stackoverflow.com/questions/20394194/exporting-to-pdf-from-an-access-report-with-a-name-from-the-data
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/1/2015 - initial version
' ---------------------------------

Public Sub OutputPDF(rpt As Report, strFilter As String, strWhere As String, strOrderBy As String, _
                        strPDFName As String, strPath As String)

On Error GoTo Err_Handler
    
    Dim strReportName As String

    DoCmd.OpenReport rpt, acViewPreview, strFilter, strWhere, acHidden, strOrderBy

    strReportName = strPath & rpt.name & ".pdf"

    DoCmd.OutputTo acOutputReport, "", acFormatPDF, strReportName, False

    DoCmd.Close acReport, rpt
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - OutputPDF[mod_Export])"
    End Select
    Resume Exit_Sub
End Sub