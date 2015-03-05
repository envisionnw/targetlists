Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Init
' Description:  initialize functions & procedures
'
' Source/date:  Bonnie Campbell, 2/9/2015
' Revisions:    BLC - 2/9/2015 - initial version
' =================================

' ---------------------------------
' SUB:          Initialize
' Description:  initialize application values
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015  - initial version
'   BLC - 2/19/2015 - added dynamic getParkState() & standard error handling
'   BLC - 3/4/2015  - shifted colors to mod_Color, removed setting of park, state, tgtYear TempVars
' ---------------------------------
Public Sub Initialize()
On Error GoTo Err_Handler

    'TempVars.item("park") = "ARCH"
    'TempVars.item("state") = getParkState(TempVars.item("park"))
    'TempVars.item("tgtYear") = 2013

    '------------------------
    'set standard variables
    '------------------------
    'std control colors
    TempVars.Add "ctrlDisabled", lngLtGray
    TempVars.Add "ctrlAddEnabled", lngLime
    TempVars.Add "ctrlRemoveEnabled", lngLtOrange
    TempVars.Add "textEnabled", lngBlue
    TempVars.Add "textDisabled", lngGray

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Initialize[mod_Init])"
    End Select
    Resume Exit_Sub
End Sub