Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Print
' Description:  printing functions & procedures
'
' Source/date:  Bonnie Campbell, 1/30/2015
' Revisions:    BLC - 1/30/2015 - initial version
' =================================

' ---------------------------------
'  Properties
' ---------------------------------
'***App Window Constants***
Public Const WIN_NORMAL = 1         'Open Normal
Public Const WIN_MAX = 3            'Open Maximized
Public Const WIN_MIN = 2            'Open Minimized

'***Error Codes***
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&

' ---------------------------------
' FUNCTION:     apiShellExecute
' Description:  XX
' Assumptions:  -
' Parameters:   hwnd -
'               lpOperation -
'               lpFile -
'               lpParameters -
'               lpDirectory -
'               nShowCmd -
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Dirk Goldgar, MVP, December 31, 2013
' https://social.msdn.microsoft.com/Forums/office/en-US/2423c0af-3eec-4320-8e37-2ac9d28d5f98/access-vba-print-copies-of-external-file?forum=accessdev
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015 - initial version
' ---------------------------------
Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

' ---------------------------------
' SUB:          PrintFile
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015 - initial version
' ---------------------------------
Public Sub xPrintFile(fileName As String, filePath As String)
    
On Error GoTo Err_Handler


Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PrintFile[mod_Print])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' FUNCTION:     PrintFile
' Description:  XX
' Assumptions:  -
' Parameters:   strFilePath - full file path for file being printed
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Dirk Goldgar, MVP, December 31, 2013
' https://social.msdn.microsoft.com/Forums/office/en-US/2423c0af-3eec-4320-8e37-2ac9d28d5f98/access-vba-print-copies-of-external-file?forum=accessdev
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015 - initial version
' ---------------------------------
Public Function PrintFile(strFilePath As String)

    ' Uses ShellExecute to print, rather than open the file.

    Dim lRet As Long, varTaskID As Variant
    Dim strRet As String

    lRet = apiShellExecute(hWndAccessApp, "print", _
            strFilePath, vbNullString, vbNullString, 0&)
            
    If lRet > ERROR_SUCCESS Then
        strRet = vbNullString
        lRet = -1
    Else
        Select Case lRet
            Case ERROR_NO_ASSOC:
                strRet = "Error: No associated application.  Couldn't print!"
            Case ERROR_OUT_OF_MEM:
                strRet = "Error: Out of Memory/Resources. Couldn't print!"
            Case ERROR_FILE_NOT_FOUND:
                strRet = "Error: File not found.  Couldn't print!"
            Case ERROR_PATH_NOT_FOUND:
                strRet = "Error: Path not found. Couldn't print!"
            Case ERROR_BAD_FORMAT:
                strRet = "Error:  Bad File Format. Couldn't print!"
            Case Else:
        End Select
    End If
    
    PrintFile = lRet & _
                IIf(strRet = "", vbNullString, ", " & strRet)

Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PrintFile[mod_Print])"
    End Select
    Resume Exit_Function
End Function