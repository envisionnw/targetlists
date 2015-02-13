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
'  Properties
' ---------------------------------
' http://cloford.com/resources/colours/500col.htm
' vbGrayText            &H80000011  Grayed (disabled) text
' vbInactiveTitleBar    &H80000003  Color of the title bar for the inactive window
' Andy Pope, March 7, 2003
' http://www.ozgrid.com/forum/showthread.php?t=49072
' Microsoft
' https://msdn.microsoft.com/en-us/library/office/aa195896%28v=office.11%29.aspx
Public Const lngGray As Long = 8224125      '?RGB(125, 125, 125)
Public Const lngLtGray As Long = 13882323   '?RGB(211, 211, 211)
Public Const lngLime As Long = 6750105      '?RGB(153, 255, 102) #99FF66
Public Const lngBlue As Long = 16711680     '?RGB(0, 0, 255) #0000FF
Public Const lngLtOrange As Long = 52479    '?RGB(255,204,0) #FFCC00
Public Const lngLtLime As Long = 6750156    '?RGB(204,255,102) #CCFF66
Public Const lngDkLime As Long = 52377      '?RGB(153,204,0) #99CC00
Public Const lngBrtLime As Long = 3407769   '?RGB(153,255,51) #99FF33
Public Const lngLtGreen As Long = 52224     '?RGB(0,204,0) #00CC00
Public Const lngDkGray As Long = 2375487      '?RGB(63,63,63) #3F3F3F



' ---------------------------------
' SUB:          Initialize
' Description:  initialize application values
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
Public Sub Initialize()

    TempVars.item("Park") = "ARCH"
    TempVars.item("state") = "UT"

    '------------------------
    'set standard variables
    '------------------------
    'std control colors
    TempVars.Add "ctrlDisabled", lngLtGray
    TempVars.Add "ctrlAddEnabled", lngLime
    TempVars.Add "ctrlRemoveEnabled", lngLtOrange
    TempVars.Add "textEnabled", lngBlue
    TempVars.Add "textDisabled", lngGray

End Sub