Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Forms
' Description:  generic form functions & procedures
'
' Source/date:  Bonnie Campbell, 2/19/2015
' Revisions:    BLC - 2/19/2015 - initial version
' =================================

'=================================================================
'  References
'=================================================================
' ---------------------------------
'  Access Control Types
' ---------------------------------
' dbtech1, March 13, 2008
' http://www.utteraccess.com/forum/control-type-vba-t1609220.html
' 126 - acAttachment         119 - acCustomControl  114 - acObjectFrame    101 - acRectangle
' 108 - acBoundObjectFrame   103 - acImage          105 - acOptionButton   112 - acSubform
' 106 - acCheckBox           100 - acLabel          107 - acOptionGroup    123 - acTabCtl
' 111 - acComboBox           102 - acLine           124 - acPage           109 - acTextBox
' 104 - acCommandButton      110 - acListBox        118 - acPageBreak      122 - acToggleButton
' ---------------------------------

' ---------------------------------
'  Access Form Sections
' ---------------------------------
'   acDetail        0   (Default) Detail section    acGroupLevel1Footer 6   Group-level 1 footer (reports only)
'   acFooter        2   Form or report footer       acGroupLevel1Header 5   Group-level 1 header (reports only)
'   acHeader        1   Form or report header       acGroupLevel2Footer 8   Group-level 2 footer (reports only)
'   acPageFooter    4   Page footer                 acGroupLevel2Header 7   Group-level 2 header (reports only)
'   acPageHeader    3   Page header
' ---------------------------------

' ---------------------------------
'  Access Backstyle Property
' ---------------------------------
'  Transparent  0           Normal  1
' ---------------------------------

'=================================================================
'  Declarations
'=================================================================
Declare Function IsZoomed Lib "User32" (ByVal hWnd As Long) As _
     Integer
Declare Function IsIconic Lib "User32" (ByVal hWnd As Long) As _
     Integer

'=================================================================
'  Properties
'=================================================================


'=================================================================
'  Subroutines & Functions
'=================================================================

' ---------------------------------
' SUB:          ClearFields
' Description:  initialize application values
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 20, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/20/2015  - initial version
' ---------------------------------
Public Sub ClearFields(frm As Form)
On Error GoTo Err_Handler

    Select Case frm.name
    
        Case "frmSpeciesSearch"
            frm.Controls("cbxCO").DefaultValue = False
            frm.Controls("cbxUT").DefaultValue = False
            frm.Controls("cbxWY").DefaultValue = False
            frm.Controls("cbxITIS").DefaultValue = False
            frm.Controls("cbxCommon").DefaultValue = False
            frm.Controls("tbxSearchFor").Value = ""
    End Select
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxITIS_Click[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          ResetHeaders
' Description:  reset header fields to their
' Assumptions:  if only a subset of form controls are to be reset, these controls should have the same Tag property value
' Parameters:   frm - form to reset headers on
'               allCtrls - if all form controls should be reset (boolean) (true = reset all controls,
'                           false = reset one control [requires oCtrl to be populated])
'               ctrlTag - control's tag string if resetting only a subset of forms controls (string)
'               fontBold - whether text should be bold (boolean) (true = make font bold, false not bold),  (optional)
'               backstyle - if back control back color is normal or transparent (integer) (1-normal 0-transparent) (optional)
'               forecolor - text color (long) (optional)
'               backcolor - backgound color of control (long) (optional)
'               oCtrl - control to change, if only one control is to be changed (optional)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Fionnuala January 20, 2013
' http://stackoverflow.com/questions/3344649/how-to-loop-through-all-controls-in-a-form-including-controls-in-a-subform-ac
' Adapted:      Bonnie Campbell, February 20, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/20/2015  - initial version
' ---------------------------------
Public Sub ResetHeaders(frm As Form, _
                        allCtrls As Boolean, _
                        ctrlTag As String, _
                        Optional fontBold As Boolean = True, _
                        Optional backstyle As Integer = 1, _
                        Optional forecolor As Long, _
                        Optional backcolor As Long, _
                        Optional oCtrl As Control)
On Error GoTo Err_Handler

Dim ctrl As Control

    If allCtrls = True Then
    
        'iterate through all form controls
        For Each ctrl In frm
            
            'check control type
             If ctrl.ControlType = acTextBox Or _
                ctrl.ControlType = acComboBox Or _
                ctrl.ControlType = acListBox Or _
                ctrl.ControlType = acLabel _
             Then
             
                'check tag
                If ctrl.tag = ctrlTag Then
                    If VarType(fontBold) = vbBoolean Then ctrl.fontBold = fontBold
                    If VarType(backstyle) = vbInteger Then ctrl.backstyle = backstyle
                    If VarType(backcolor) = vbLong Then ctrl.backcolor = backcolor
                    If VarType(forecolor) = vbLong Then ctrl.forecolor = forecolor
                End If
                
          End If
          
        Next
    Else
        'reset only oCtrl

        'check tag
        If oCtrl.tag = ctrlTag Then
        
            'check control type
            If oCtrl.ControlType = acTextBox Or _
                oCtrl.ControlType = acComboBox Or _
                oCtrl.ControlType = acListBox Or _
                oCtrl.ControlType = acLabel _
            Then
          
                If VarType(fontBold) = vbBoolean Then oCtrl.fontBold = fontBold
                If VarType(backstyle) = vbInteger Then oCtrl.backstyle = backstyle
                If VarType(backcolor) = vbLong Then oCtrl.backcolor = backcolor
                If VarType(forecolor) = vbLong Then oCtrl.forecolor = forecolor
             
            End If
            
        End If

    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ResetHeaders[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          ResetHeaders
' Description:  reset header fields to their
' Assumptions:  if only a subset of form controls are to be reset, these controls should have the same Tag property value
' Parameters:   frm - form to reset headers on
'               allCtrls - if all form controls should be reset (boolean) (true = reset all controls,
'                           false = reset one control [requires oCtrl to be populated])
'               ctrlTag - control's tag string if resetting only a subset of forms controls (string)
'               visibility - whether control should be visible or not (boolean) (true = make font bold, false not bold),  (optional)
'               oCtrl - control to change, if only one control is to be changed (optional)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Fionnuala January 20, 2013
' http://stackoverflow.com/questions/3344649/how-to-loop-through-all-controls-in-a-form-including-controls-in-a-subform-ac
' Adapted:      Bonnie Campbell, February 20, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/20/2015  - initial version
' ---------------------------------
Public Sub ShowControls(frm As Form, _
                        allCtrls As Boolean, _
                        ctrlTag As String, _
                        visibility As Boolean, _
                        Optional oCtrl As Control)
On Error GoTo Err_Handler

Dim ctrl As Control

    If allCtrls = True Then
    
        'iterate through all form controls
        For Each ctrl In frm

            'check tag
            If ctrl.tag = ctrlTag Then
                ctrl.Visible = visibility
            End If

        Next
    Else
        'reset only oCtrl

        'check tag
        If oCtrl.tag = ctrlTag Then
                oCtrl.Visible = visibility
        End If

    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ShowControls[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          AddControl
' Description:  initialize application values
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' meloncolly, October 27, 2006
' http://forums.aspfree.com/microsoft-access-help-18/add-controls-form-dynamically-139627.html
' https://msdn.microsoft.com/en-us/library/bb237827(office.12).aspx
' Adapted:      Bonnie Campbell, February 19, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/19/2015  - initial version
' ---------------------------------
Public Sub AddControl(frm As Form, ctrl As Control, ctrlName As String, _
                        xPos As Integer, yPos As Integer)
On Error GoTo Err_Handler

    ' Create ctrl
    Set ctrl = CreateControl(frm.name, ctrl.ControlType, , "", "", xPos, yPos)
    
    ' Restore form
    DoCmd.Restore

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxITIS_Click[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub