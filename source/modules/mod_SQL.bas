Option Compare Database
Option Explicit
' =================================
' MODULE:       mod_SQL
' Description:  SQL functions & procedures
'
' Source/date:  Bonnie Campbell, 4/7/2015
' Revisions:    BLC - 4/7/2015 - initial version
' =================================

' ---------------------------------
'  Subroutines & Functions
' ---------------------------------

' ---------------------------------
' SUB:          ConcatRelated
' Description:  Used in SQL queries to generate concatenated string of related records
' Assumptions:  used in Access SQL or control
' Parameters:   strField - field to retrieve results from & concatenate (string)
'               strTable - table or query name (string)
'               strWHERE - limiting WHERE clause (string)
'               strOrderBy - sorting ORDER BY clause (string)
'               strSeparator - character to use between concatenated values (string)
' Returns:      SQL (string, variant, or NULL if no matches)
' Notes:        1. Use square brackets around field/table names with spaces or odd characters.
'               2. strField can be a Multi-valued field (A2007 and later), but strOrderBy cannot.
'               3. Nulls are omitted, zero-length strings (ZLSs) are returned as ZLSs.
'               4. Returning more than 255 characters to a recordset triggers this Access bug:
'                  http://allenbrowne.com/bug-16.html
' Usage:        SQL string:
'                SELECT CompanyName,  ConcatRelated("OrderDate", "tblOrders", "CompanyID = "
'                   & [CompanyID]) FROM tblCompany;
'               Access textbox control:
'                =ConcatRelated("OrderDate", "tblOrders", "CompanyID = " & [CompanyID])
' Throws:       none
' References:   none
' Source/date:
' Allen Browne, June, 2008
' http://allenbrowne.com/func-concat.html
' Adapted:      Bonnie Campbell, April 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/7/2015  - initial version
' ---------------------------------
Public Function ConcatRelated(strField As String, _
    strTable As String, _
    Optional strWhere As String, _
    Optional strOrderBy As String, _
    Optional strSeparator = ", ") As Variant
On Error GoTo Err_Handler
    Dim rs As DAO.Recordset         'Related records
    Dim rsMV As DAO.Recordset       'Multi-valued field recordset
    Dim strSQL As String            'SQL statement
    Dim strOut As String            'Output string to concatenate to.
    Dim lngLen As Long              'Length of string.
    Dim bIsMultiValue As Boolean    'Flag if strField is a multi-valued field.
    
    'Initialize to Null
    ConcatRelated = Null
    
    'Build SQL string, and get the records.
    strSQL = "SELECT " & strField & " FROM " & strTable
    If strWhere <> vbNullString Then
        strSQL = strSQL & " WHERE " & strWhere
    End If
    If strOrderBy <> vbNullString Then
        strSQL = strSQL & " ORDER BY " & strOrderBy
    End If
    Set rs = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenDynaset)
    'Determine if the requested field is multi-valued (Type is above 100.)
    bIsMultiValue = (rs(0).Type > 100)
    
    'Loop through the matching records
    Do While Not rs.EOF
        If bIsMultiValue Then
            'For multi-valued field, loop through the values
            Set rsMV = rs(0).Value
            Do While Not rsMV.EOF
                If Not IsNull(rsMV(0)) Then
                    strOut = strOut & rsMV(0) & strSeparator
                End If
                rsMV.MoveNext
            Loop
            Set rsMV = Nothing
        ElseIf Not IsNull(rs(0)) Then
            strOut = strOut & rs(0) & strSeparator
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    'Return the string without the trailing separator.
    lngLen = Len(strOut) - Len(strSeparator)
    If lngLen > 0 Then
        ConcatRelated = Left(strOut, lngLen)
    End If

Exit_Function:
    'Clean up
    Set rsMV = Nothing
    Set rs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ConcatRelated[mod_SQL])"
    End Select
    Resume Exit_Function
End Function



' ---------------------------------
' SUB:          ConcatRelated
' Description:  Used in SQL queries to generate concatenated string of related records
' Assumptions:  used in Access SQL or control
' Parameters:   strSQL - field to retrieve results from & concatenate (string)
'               NameList() - limiting WHERE clause (string)
'               strDelim - character to use between concatenated values (string)
' Returns:      SQL (string, variant, or NULL if no matches)
' Usage:        SQL string:
'               SELECT documents.MembersOnly, Coalsce("SELECT FName From Persons WHERE Member=True",":") AS Who,
'               Coalsce("", ":", "Mary", "Joe", "Pat?") As Others FROM documents;
' Throws:       none
' References:   none
' Source/date:
' Fionuala, September 18, 2008
' http://stackoverflow.com/questions/92698/combine-rows-concatenate-rows?lq=1
' Adapted:      Bonnie Campbell, April 8, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/8/2015  - initial version
' ---------------------------------
Function Coalsce(strSQL As String, strDelim, ParamArray NameList() As Variant)
Dim db As Database
Dim rs As DAO.Recordset
Dim strList As String

    Set db = CurrentDb

    If strSQL <> "" Then
        Set rs = db.OpenRecordset(strSQL)

        Do While Not rs.EOF
            strList = strList & strDelim & rs.Fields(0)
            rs.MoveNext
        Loop

        strList = Mid(strList, Len(strDelim))
    Else

        strList = Join(NameList, strDelim)
    End If

    Coalsce = strList

Exit_Function:
    'Clean up
    Set rs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Coalesce[mod_SQL])"
    End Select
    Resume Exit_Function
End Function