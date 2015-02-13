﻿Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =11
    ItemSuffix =5
    Right =20208
    Bottom =9660
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xc1f3db6ed487e440
    End
    RecordSource ="tbl_Target_Areas"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    SplitFormOrientation =1
    SplitFormSplitterBar =0
    SplitFormOrientation =1
    SplitFormSplitterBar =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =2400
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =1512
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblTgtAreaHdr"
                    Caption ="Target Areas"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1572
                    LayoutCachedHeight =432
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =540
                    Top =1080
                    Width =4140
                    Height =360
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxTgtArea"
                    DefaultValue ="\"\""
                    OnKeyUp ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="Enter new target area"
                    GridlineColor =10921638

                    LayoutCachedLeft =540
                    LayoutCachedTop =1080
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =1440
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =600
                            Width =4476
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblTgtArea"
                            Caption ="Enter the target area name."
                            GridlineColor =10921638
                            LayoutCachedLeft =240
                            LayoutCachedTop =600
                            LayoutCachedWidth =4716
                            LayoutCachedHeight =900
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4200
                    Top =1620
                    Width =2220
                    TabIndex =1
                    ForeColor =16711680
                    Name ="btnAddTgtArea"
                    Caption ="Add new Target Area"
                    StatusBarText ="Add new target area"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4200
                    LayoutCachedTop =1620
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =1980
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =6750156
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =52377
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =3407769
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =52224
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =2375487
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =6750156
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                    Overlaps =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =7560
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =2448
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblTgtAreaList"
                    Caption ="Current Target Areas"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =2508
                    LayoutCachedHeight =432
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =180
                    Top =540
                    Width =420
                    Height =300
                    ForeColor =4210752
                    Name ="btnDeleteTgtArea"
                    Caption ="Delete Target Area"
                    ControlTipText ="Delete Record"
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Action ="OnError"
                            Argument ="0"
                        End
                        Begin
                            Action ="GoToControl"
                            Argument ="=[Screen].[PreviousControl].[Name]"
                        End
                        Begin
                            Action ="ClearMacroError"
                        End
                        Begin
                            Condition ="Not [Form].[NewRecord]"
                            Action ="RunCommand"
                            Argument ="223"
                        End
                        Begin
                            Condition ="[Form].[NewRecord] And Not [Form].[Dirty]"
                            Action ="Beep"
                        End
                        Begin
                            Condition ="[Form].[NewRecord] And [Form].[Dirty]"
                            Action ="RunCommand"
                            Argument ="292"
                        End
                        Begin
                            Condition ="[MacroError]<>0"
                            Action ="MsgBox"
                            Argument ="=[MacroError].[Description]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"btnDeleteTgtArea\" xmlns=\"http://schemas.microsoft.com/offi"
                                "ce/accessservices/2009/11/application\" xmlns:a=\"http://schemas.microsoft.com/o"
                                "ffice/accessservices/2009/11/f"
                        End
                        Begin
                            Comment ="_AXL:orms\"><Statements><Action Name=\"OnError\"/><Action Name=\"GoToControl\"><"
                                "Argument Name=\"ControlName\">=[Screen].[PreviousControl].[Name]</Argument></Act"
                                "ion><Action Name=\"ClearMacroError\"/><ConditionalBlock><If><Condition>Not [Form"
                                "].[NewRecord]</Conditi"
                        End
                        Begin
                            Comment ="_AXL:on><Statements><Action Name=\"DeleteRecord\"/></Statements></If></Condition"
                                "alBlock><ConditionalBlock><If><Condition>[Form].[NewRecord] And Not [Form].[Dirt"
                                "y]</Condition><Statements><Action Name=\"Beep\"/></Statements></If></Conditional"
                                "Block><Conditiona"
                        End
                        Begin
                            Comment ="_AXL:lBlock><If><Condition>[Form].[NewRecord] And [Form].[Dirty]</Condition><Sta"
                                "tements><Action Name=\"UndoRecord\"/></Statements></If></ConditionalBlock><Condi"
                                "tionalBlock><If><Condition>[MacroError]&lt;&gt;0</Condition><Statements><Action "
                                "Name=\"MessageBo"
                        End
                        Begin
                            Comment ="_AXL:x\"><Argument Name=\"Message\">=[MacroError].[Description]</Argument></Acti"
                                "on></Statements></If></ConditionalBlock></Statements></UserInterfaceMacro>"
                        End
                    End
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b0a090ff302010ff302010ff302010ff302010ff302010ff ,
                        0x302010ff302010ff302010ff302010ff302010ff302010ff302010ff00000000 ,
                        0x0000000000000000b0a090fffff8f0fffff0f0ffffe8e0fff0e8e0fff0e0d0ff ,
                        0xf0d8d0fff0d8c0fff0d8c0fff0d8c0fff0d8c0fff0d8c0ff302010ff00000000 ,
                        0x0000000000000000b0a090ffffffffffe06830ffe06830ffe06830ffd06830ff ,
                        0xd06830ffd06830ffd06030ffc06030ff904820ffffe0d0ff302010ff00000000 ,
                        0x0000000000000000b0a090ffffffffffd06830ffffb080ffffa880ffffa070ff ,
                        0xf09870fff09060ffa0b0f0ff1020e0ffc0c8f0ffffe0d0ff302010ff00000000 ,
                        0x00000000a0a8f0ffb0a090ffffffffffe06830ffe06830ffe06830ffd06830ff ,
                        0xd06830ffe0e0f0ff0028ffff1028f0ff4050d0ffffe0d0ff302010ff00000000 ,
                        0x4050e0ff0010b0ffb0a090ffffffffffffffffffffffffffffffffffffffffff ,
                        0xfff8f0ffffe8e0ff2048ffff1038ffff1028ffffe0e8f0ff302010ff7088f0ff ,
                        0x0018c0ff6078f0ffb0a090ffb0a090ffb0a090ffb0a090ffb0a090ffb0a090ff ,
                        0xb0a090ffb0a090ffe0e0f0ff3050ffff2040ffff8090f0ffb0b8f0ff0028f0ff ,
                        0x4058f0ff00000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000d0d8f0ff4060ffff3050ffff2040ffff3050ffff ,
                        0xe0e8f0ff00000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000c0d0f0ff4068ffff4060ffffc0c8f0ff ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000c0c8f0ff6078ffff6078ffff6080ffff5070ffff ,
                        0xe0e0f0ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b0b8f0ff6078ffff6078ffffb0c0f0fff0f0f0ff7088ffff ,
                        0x6078ffffc0d0f0ff000000000000000000000000000000000000000000000000 ,
                        0x0000000090a0ffff6078ffff6078ffffd0d8f0ff000000000000000000000000 ,
                        0xb0b8f0ff8098ffff000000000000000000000000000000000000000000000000 ,
                        0x000000008098ffff6080ffffd0d8f0ff00000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =180
                    LayoutCachedTop =540
                    LayoutCachedWidth =600
                    LayoutCachedHeight =840
                    Gradient =0
                    BackThemeColorIndex =1
                    BackTint =100.0
                    OldBorderStyle =0
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =3
                    WebImagePaddingBottom =3
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =720
                    Top =540
                    Width =2640
                    Height =300
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxTgtAreaName"
                    ControlSource ="Target_Area"
                    GridlineColor =10921638

                    LayoutCachedLeft =720
                    LayoutCachedTop =540
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =840
                End
            End
        End
        Begin FormFooter
            Height =360
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' MODULE:       Form_frmTgtSpeciesList
' Description:  Target species functions & procedures
'
' Source/date:  Bonnie Campbell, 2/11/2015
' Revisions:    BLC - 2/11/2015 - initial version
' =================================

' ---------------------------------
' SUB:          Form_Load
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler
    
    Initialize
       
    If Len(tbxTgtArea.Value) = 0 Then
        'disable search until something is entered
        DisableControl lblAddTgtArea
        btnAddTgtArea.Enabled = False
        DisableControl btnAddTgtArea
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[form_frmTgtAreas])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxTgtArea_Change
' Description:  Actions to take when new target area textbox is not empty
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Private Sub tbxTgtArea_Change()
On Error GoTo Err_Handler
    
    If Len(tbxTgtArea.Value) > 0 Then
        'enable the search "button"
        EnableControl lblAddTgtArea, TempVars.item("ctrlAddEnabled"), TempVars.item("textEnabled")
        btnAddTgtArea.Enabled = True
    Else
        'disable the search "button"
        DisableControl lblAddTgtArea
        btnAddTgtArea.Enabled = False
        DisableControl btnAddTgtArea
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxTgtArea_Change[Form_frmTgtAreas])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxTgtArea_LostFocus
' Description:  Actions to take when new target area textbox is not empty
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 10, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/10/2015 - initial version
' ---------------------------------
Private Sub tbxTgtArea_LostFocus()
On Error GoTo Err_Handler
    
    If Len(tbxTgtArea.Value) > 0 Then
        'enable the search "button"
        EnableControl lblAddTgtArea, TempVars.item("ctrlAddEnabled"), TempVars.item("textEnabled")
        btnAddTgtArea.Enabled = True
    Else
        'disable the search "button"
        DisableControl lblAddTgtArea
        btnAddTgtArea.Enabled = False
        DisableControl btnAddTgtArea
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxTgtArea_LostFocus[Form_frmTgtAreas])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxTgtArea_KeyUp
' Description:  Actions to take when new target area textbox is not empty
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Private Sub tbxTgtArea_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler
    
    If Len(tbxTgtArea.Value) > 0 Then
        'enable the search "button"
        EnableControl lblAddTgtArea, TempVars.item("ctrlAddEnabled"), TempVars.item("textEnabled")
        btnAddTgtArea.Enabled = True
        EnableControl btnAddTgtArea, lngLtLime, lngBlue, lngDkLime, lngBrtLime, lngLtGreen, lngDkGray, lngLtLime
    Else
        'disable the search "button"
        DisableControl lblAddTgtArea
        btnAddTgtArea.Enabled = False
        DisableControl btnAddTgtArea
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxTgtArea_LostFocus[Form_frmTgtAreas])"
    End Select
    Resume Exit_Sub
End Sub



' ---------------------------------
' SUB:          lblAddTgtArea_Click
' Description:  Open add target area form
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Private Sub lblAddTgtArea_Click()
On Error GoTo Err_Handler
    Dim strTgtArea As String
    'Dim strSQL As String
    
    'strSQL = "INSERT INTO tbl_Target_Areas(Target_Area) VALUES "
    
    If ValidateString(tbxTgtArea.Value, "alphaspace") = True Then
        strTgtArea = Trim(tbxTgtArea.Value)
        
        'strSQL = strSQL & "('" & strTgtArea & "')"
        
        Dim rs As Recordset
    
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM [tbl_Target_Areas]")
        rs.AddNew
        
        rs![Target_Area] = strTgtArea
        rs.Update
        rs.Close
        Set rs = Nothing
        DoCmd.Close
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblAddTgtArea_Click[Form_frmTgtAreas])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnAddTgtArea_Click
' Description:  Add target area
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Private Sub btnAddTgtArea_Click()
On Error GoTo Err_Handler
    Dim strTgtArea As String
    'Dim strSQL As String
    
    'strSQL = "INSERT INTO tbl_Target_Areas(Target_Area) VALUES "
    
    If ValidateString(tbxTgtArea.Value, "alphaspace") = True Then
        strTgtArea = Trim(tbxTgtArea.Value)
        
        'strSQL = strSQL & "('" & strTgtArea & "')"
        
        Dim rs As Recordset
    
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM [tbl_Target_Areas]")
        rs.AddNew
        
        rs![Target_Area] = strTgtArea
        rs.Update
        rs.Close
        Set rs = Nothing
        DoCmd.Close
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddTgtArea_Click[Form_frmTgtAreas])"
    End Select
    Resume Exit_Sub
End Sub
