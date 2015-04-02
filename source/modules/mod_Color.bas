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
' http://rainbowprod.com/english/powerbuilder/colors.html
' long value = (65536*Blue) + (256*Green) + (Red)
' Anonymous, March 9, 1999
' http://www.vbcode.com/asp/showsn.asp?theID=311
' Convert RGB to LONG:      LONG = B * 65536 + G * 256 + R
' ---------------------------------
Public Const lngGray As Long = 8224125      '?RGB(125, 125, 125)
Public Const lngLtGray As Long = 13882323   '?RGB(211, 211, 211)
Public Const lngLime As Long = 6750105      '?RGB(153, 255, 102) #99FF66
Public Const lngBlue As Long = 16711680     '?RGB(0, 0, 255) #0000FF
Public Const lngLtOrange As Long = 52479    '?RGB(255,204,0) #FFCC00
Public Const lngLtLime As Long = 6750156    '?RGB(204,255,102) #CCFF66
Public Const lngDkLime As Long = 52377      '?RGB(153,204,0) #99CC00
Public Const lngBrtLime As Long = 3407769   '?RGB(153,255,51) #99FF33
Public Const lngLtGreen As Long = 52224     '?RGB(0,204,0) #00CC00
Public Const lngDkGray As Long = 2375487    '?RGB(63,63,63) #3F3F3F
Public Const lngYelLime As Long = 9699294   '?RGB(222,255,147) #DEFF93
Public Const lngWhite As Long = 16777215    '?RGB(255,255,255) #FFFFFF
Public Const lngBlack As Long = 0           '?RGB(0,0,0) #000000
Public Const lngYellow As Long = 65535      '?RGB(255,255,0) #FFFF00
Public Const lngLtYellow As Long = 14745599 '?RGB(255,255,224) #FFFFE0



' ---------------------------------
' SUB:          ConvertLongToRGB
' Description:  Convert long color value to RGB component values
' Assumptions:  User will call specific color values via dict("R"), dict("G"), dict("B") as needed
' Parameters:   lngValue - long color value
' Returns:      RGB - as dicitionary object (R, G, B components)
' Throws:       none
' References:   none
' Source/date:
' Chirag, March 12, 2008
' http://www.pcreview.co.uk/threads/convert-long-integer-color-value-to-rgb.3447564/
' JTolle, August 21, 2009
' http://stackoverflow.com/questions/1309689/hash-table-associative-array-in-vba
' Adapted:      Bonnie Campbell, March 4, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/4/2015  - initial version
' ---------------------------------
Public Function ConvertLongToRGB(ByVal lngValue As Long) As Dictionary
On Error GoTo Err_Handler
    Dim dRGB As Dictionary
    Set dRGB = New Dictionary
       
    dRGB("R") = lngValue Mod 256
    dRGB("G") = Int(lngValue / 256) Mod 256
    dRGB("B") = Int(lngValue / 256 / 256) Mod 256
    
    Set ConvertLongToRGB = dRGB
    
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ConvertLongToRGB[mod_Color])"
    End Select
    Resume Exit_Function
End Function