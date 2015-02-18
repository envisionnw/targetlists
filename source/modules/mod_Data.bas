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
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
' ---------------------------------
Public Sub fillList(frm As Form, lbx As ListBox)

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strQuery As String, strSQL As String
    Dim lbxDest As ListBox

    Select Case lbx.name
    
        Case "lbxDataSheets" 'Datasheets
            strQuery = "qryActiveDatasheets"
            strSQL = CurrentDb.QueryDefs(strQuery).sql
            Set lbxDest = frm.Controls("lbxPrintSheets")
            
        Case "lbxSpecies" 'Species
            strQuery = "qryPlantSpecies"
            strSQL = CurrentDb.QueryDefs(strQuery).sql
            MsgBox strSQL & "species"
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