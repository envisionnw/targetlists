Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Strings
' Description:  string functions & procedures
'
' Source/date:  Bonnie Campbell, 2/10/2015
' Revisions:    BLC - 2/10/2015 - initial version
' =================================

' ---------------------------------
' FUNCTION:     CountInString
' Description:  Counts the number of instances of character(s) in a string
' Assumptions:  -
' Parameters:   strInspect - string to check
'               strFind - string to count
' Returns:      count - number o instances strFind is found in strInspect
' Throws:       none
' References:   none
' Source/date:
'
' http://stackoverflow.com/questions/5193893/count-specific-character-occurrences-in-string
' Scott Huish, June 20, 2011
' http://www.mrexcel.com/forum/excel-questions/558568-count-occurrence-string-within-string-using-visual-basic-applications.html
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015 - initial version
' ---------------------------------
Public Function CountInString(ByVal strInspect As String, ByVal strFind As String) As Integer
On Error GoTo Err_Handler:
     Dim count As Integer

    'default
    count = 0
    
    If Len(strInspect) > 0 Then
        count = UBound(Split(strInspect, strFind))
    End If
    
    CountInString = count

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CountInString[mod_Strings])"
    End Select
    Resume Exit_Function
End Function