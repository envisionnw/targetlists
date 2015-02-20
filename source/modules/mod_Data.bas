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
Public Sub fillList(frm As Form, ctrlSource As Control, ctrlDest As Control)

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
            'Set lbxDest = frm.Controls("lbxPrintSheets")
            
        Case "lbxSpecies", "sfrmSpeciesListbox" 'Species
            strQuery = "qryPlantSpecies"
            strSQL = CurrentDb.QueryDefs(strQuery).sql
 '           MsgBox strSQL & "species"
            'Set lbxDest = frm.Controls("lbxTgtSpecies")
            
    End Select

    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)

    'set TempVars
    TempVars.Add "strSQL", strSQL

    'PopulateList frm.Controls(lbx), rs
    PopulateList ctrlSource, rs, ctrlDest
    
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

' ---------------------------------
' SUB:          XfillList
' Description:  Fill a list (or listbox like subform) from specific queries for datasheets, species or other items
' Assumptions:  Either a listbox or subform control is being populated
' Parameters:   frm - main form object
'               lbx - main form listbox object (for filling a listbox control)
'               sfrm - subform object (for populating a subform control)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015  - initial version
'   BLC - 2/18/2015 - adapted to include subform as well as listbox controls
' ---------------------------------
Public Sub XfillList(frm As Form, Optional lbx As ListBox, Optional sfrm As Form)

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strQuery As String, strSQL As String
    Dim lbxDest As ListBox
    Dim dataSource As String
    
    'output to form or listbox control?
    If Not lbx Is Nothing Then
        dataSource = lbx.name
    ElseIf Not sfrm Is Nothing Then
        dataSource = sfrm.name
    Else
        'no other options
        GoTo Exit_Sub
    End If
    
    
    Select Case dataSource
    
        Case "lbxDataSheets", "sfrmDatasheets" 'Datasheets
            strQuery = "qryActiveDatasheets"
            strSQL = CurrentDb.QueryDefs(strQuery).sql
            Set lbxDest = frm.Controls("lbxPrintSheets")
            
        Case "lbxSpecies", "sfrmSpeciesListbox" 'Species
            strQuery = "qryPlantSpecies"
            strSQL = CurrentDb.QueryDefs(strQuery).sql
 '           MsgBox strSQL & "species"
            Set lbxDest = frm.Controls("lbxTgtSpecies")
            
    End Select

    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)

    'set TempVars
    TempVars.Add "strSQL", strSQL

    'PopulateList frm.Controls(lbx), rs
    PopulateList lbx, rs, lbxDest

    'Enable move items lbls (or not)
    If lbx.ListCount > 0 Then
        frm.Controls("lblAddAll").Visible = True
        frm.Controls("lblRemoveAll").Visible = True
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