Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Data
' Description:  data functions & procedures
'
' Source/date:  Bonnie Campbell, 2/9/2015
' Revisions:    BLC - 2/9/2015 - initial version
' =================================

' ---------------------------------
' SUB:          fillList
' Description:  Fill a list (or listbox like subform) from specific queries for datasheets, species or other items
' Assumptions:  Either a listbox or subform control is being populated
' Parameters:   frm - main form object
'               ctrl - either:
'                      lbx - main form listbox object (for filling a listbox control)
'                      sfrm - subform object (for populating a subform control)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015  - initial version
'   BLC - 2/18/2015 - adapted to include subform as well as listbox controls
' ---------------------------------
Public Sub fillList(frm As Form, ctrlSource As Control, Optional ctrlDest As Control)

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strQuery As String, strSQL As String
    
    'output to form or listbox control?
   
    'determine data source
    Select Case ctrlSource.name
    
        Case "lbxDataSheets", "sfrmDatasheets" 'Datasheets
            strQuery = "qryActiveDatasheets"
            strSQL = CurrentDb.QueryDefs(strQuery).sql
            
        Case "lbxSpecies", "lbxTgtSpecies", "sfrmSpeciesListbox" 'Species
            strQuery = "qryPlantSpecies"
            strSQL = CurrentDb.QueryDefs(strQuery).sql
            
    End Select

    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)

    'set TempVars
    TempVars.Add "strSQL", strSQL

    If Not ctrlDest Is Nothing Then
        'populate list & headers
        PopulateList ctrlSource, rs, ctrlDest
    Else
        'populate only ctrlSource headers
        PopulateListHeaders ctrlSource, rs
    End If
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fillList[mod_Data])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' FUNCTION:     getParkState
' Description:  Retrieve the state associated with a park (via tlu_Parks)
' Assumptions:  Park state is properly identified in tlu_Parks
' Parameters:   parkCode - 4 character park designator
' Returns:      ParkState - 2 character state abbreviation
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 19, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/19/2015  - initial version
' ---------------------------------
Public Function getParkState(parkCode As String) As String

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim state As String, strSQL As String
   
    'handle only appropriate park codes
    If Len(parkCode) <> 4 Then
        GoTo Exit_Function
    End If
    
    'generate SQL ==> NOTE: LIMIT 1; syntax not viable for Access, use SELECT TOP x instead
    strSQL = "SELECT TOP 1 ParkState FROM tlu_Parks WHERE ParkCode LIKE '" & parkCode & "';"
            
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)
    
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        state = rs.Fields("ParkState").Value
    End If
   
    'return value
    getParkState = state
    
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getParkState[mod_Data])"
    End Select
    Resume Exit_Function
End Function

'Public param As String

' ---------------------------------
' FUNCTION:     SetParam
' Description:  Set a parameter value (useful for parameter queries)
' Assumptions:  Companion GetParam() function exists & param is publicly defined
' Parameters:   paramValue - parameter name (string)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 24, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/24/2015  - initial version
' ---------------------------------
Public Function SetParam(paramValue As Variant)

On Error GoTo Err_Handler
    
    param = paramValue
    
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetParam[mod_Data])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     SetParam
' Description:  Set a parameter value (useful for parameter queries)
' Assumptions:  Companion GetParam() function exists & param is publicly defined
' Parameters:   paramValue - parameter name (string)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 24, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/24/2015  - initial version
' ---------------------------------
Public Function GetParam()

On Error GoTo Err_Handler
    
    GetParam = param
    
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetParam[mod_Data])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     DisconnectRecordset
' Description:  Create a disconnected ADO (in-memory) recordset for manipulation
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   ADO required
' Source/date:
' Danny Lesandrini, November 2, 2009
' http://www.databasejournal.com/features/msaccess/article.php/3846361/Create-In-Memory-ADO-Recordsets.htm
' Fionnuala, October 12, 2011
' http://stackoverflow.com/questions/7738811/access-listbox-based-on-value-list-sorting-on-column
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015 - initial version
' ---------------------------------
Public Function DisconnectRecordset(valueList As Variant) As Variant

On Error GoTo Err_Handler
    
'Dim rs As New adodb.Recordset
Dim rs As Recordset

'slist = "0,Standard price,1650," _
'& "14,Bookings made during Oct 2011,3770," _
'& "15,Minimum Stay 4 Nights - Special Price,2460"

With rs
 ' .ActiveConnection = Nothing
 ' .CursorLocation = adUseClient
 ' .CursorType = adOpenStatic
 ' .LockType = adLockBatchOptimistic
 ' With .Fields
 '   .Append "Field1", adInteger
 '   .Append "Field2", adVarChar, 200
 '   .Append "Field3", adInteger
 ' End With
 ' .Open

Dim Ary As Variant
Dim j As Integer, i As Integer

  Ary = Split(valueList, ",")

  For j = 0 To UBound(Ary)
      .AddNew
      For i = 0 To 2
  '        .Fields(i).Value = Ary(j)
          j = j + 1
      Next
      j = j - 1

  Next

  '.Sort = "Field3"
End With

'slist = rs.GetString(, , ",", ",")
'slist = Left(slist, Len(slist) - 1)

Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CountArrayValues[mod_Lists])"
    End Select
    Resume Exit_Function
End Function

Public Function CreateDisconnectedRecordset()
Dim rstADO As ADODB.Recordset
Dim fld As ADODB.Field

Set rstADO = New ADODB.Recordset
With rstADO
    .Fields.Append "EmployeeID", adInteger, , adFldKeyColumn
    .Fields.Append "FirstName", adVarChar, 10, adFldMayBeNull
    .Fields.Append "LastName", adVarChar, 20, adFldMayBeNull
    .Fields.Append "Email", adVarChar, 64, adFldMayBeNull
    .Fields.Append "Include", adInteger, , adFldMayBeNull
    .Fields.Append "Selected", adBoolean, , adFldMayBeNull

    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    .LockType = adLockPessimistic
    .Open
End With
End Function

' ---------------------------------
' FUNCTION:     GetQuerySQL
' Description:  Get SQL for a query
' Assumptions:  -
' Parameters:   qryName - Name of query to fetch SQL for (string)
' Returns:      qrySQL - full SQL for the query (string)
' Throws:       none
' References:   -
' Source/date:
' S. Phinney, July 13, 2009
' http://bytes.com/topic/access/answers/871500-getting-sql-string-query
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015 - initial version
' ---------------------------------
Private Function GetQuerySQL(qryName As String) As String
Dim qdf As DAO.QueryDef
 
    'fetch query
    Set qdf = CurrentDb.QueryDefs(qryName)
    
    'return SQL
    GetQuerySQL = qdf.sql
 
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetQuerySQL[mod_Data])"
    End Select
    Resume Exit_Function
End Function