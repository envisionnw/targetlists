Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9000
    DatasheetFontHeight =11
    ItemSuffix =47
    Right =20460
    Bottom =9660
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x72574db34b86e440
    End
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
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
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
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin ListBox
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
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =5172
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =4020
                    Width =9000
                    Height =1140
                    BackColor =14806254
                    BorderColor =10921638
                    Name ="boxCurrTgtArea"
                    GridlineColor =10921638
                    LayoutCachedTop =4020
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =5160
                    BackThemeColorIndex =3
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =840
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSearchHdr"
                    Caption ="Search"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =900
                    LayoutCachedHeight =432
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =300
                    Top =720
                    Width =8154
                    Height =1380
                    ColumnOrder =0
                    BorderColor =10921638
                    Name ="optgSpeciesType"
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =720
                    LayoutCachedWidth =8454
                    LayoutCachedHeight =2100
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =420
                            Top =600
                            Width =1596
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblSpeciesTypes"
                            Caption ="What to Search..."
                            GridlineColor =10921638
                            LayoutCachedLeft =420
                            LayoutCachedTop =600
                            LayoutCachedWidth =2016
                            LayoutCachedHeight =900
                            BackThemeColorIndex =-1
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =720
                    Top =1620
                    Width =240
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =10921638
                    Name ="cbxUT"
                    DefaultValue ="False"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Utah species"
                    GridlineColor =10921638

                    LayoutCachedLeft =720
                    LayoutCachedTop =1620
                    LayoutCachedWidth =960
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =1020
                            Top =1560
                            Width =525
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblUtah"
                            Caption ="Utah"
                            ControlTipText ="Utah species"
                            GridlineColor =10921638
                            LayoutCachedLeft =1020
                            LayoutCachedTop =1560
                            LayoutCachedWidth =1545
                            LayoutCachedHeight =1875
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =1980
                    Top =1620
                    Width =240
                    ColumnOrder =2
                    TabIndex =2
                    BorderColor =10921638
                    Name ="cbxCO"
                    DefaultValue ="False"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Utah species"
                    GridlineColor =10921638

                    LayoutCachedLeft =1980
                    LayoutCachedTop =1620
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =2280
                            Top =1560
                            Width =900
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblCO"
                            Caption ="Colorado"
                            ControlTipText ="Colorado species"
                            GridlineColor =10921638
                            LayoutCachedLeft =2280
                            LayoutCachedTop =1560
                            LayoutCachedWidth =3180
                            LayoutCachedHeight =1860
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =3600
                    Top =1620
                    Width =240
                    ColumnOrder =3
                    TabIndex =3
                    BorderColor =10921638
                    Name ="cbxWY"
                    DefaultValue ="False"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Wyoming species"
                    GridlineColor =10921638

                    LayoutCachedLeft =3600
                    LayoutCachedTop =1620
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3960
                            Top =1560
                            Width =936
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblWY"
                            Caption ="Wyoming"
                            ControlTipText ="Wyoming species"
                            GridlineColor =10921638
                            LayoutCachedLeft =3960
                            LayoutCachedTop =1560
                            LayoutCachedWidth =4896
                            LayoutCachedHeight =1860
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =5280
                    Top =1620
                    Width =240
                    ColumnOrder =4
                    TabIndex =4
                    BorderColor =10921638
                    Name ="cbxITIS"
                    DefaultValue ="False"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="ITIS species"
                    GridlineColor =10921638

                    LayoutCachedLeft =5280
                    LayoutCachedTop =1620
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5640
                            Top =1560
                            Width =405
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblITIS"
                            Caption ="ITIS"
                            GridlineColor =10921638
                            LayoutCachedLeft =5640
                            LayoutCachedTop =1560
                            LayoutCachedWidth =6045
                            LayoutCachedHeight =1875
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =6540
                    Top =1620
                    Width =240
                    ColumnOrder =5
                    TabIndex =5
                    BorderColor =10921638
                    Name ="cbxCommon"
                    DefaultValue ="False"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Common name"
                    GridlineColor =10921638

                    LayoutCachedLeft =6540
                    LayoutCachedTop =1620
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6900
                            Top =1560
                            Width =900
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblCommon"
                            Caption ="Common"
                            ControlTipText ="Common name"
                            GridlineColor =10921638
                            LayoutCachedLeft =6900
                            LayoutCachedTop =1560
                            LayoutCachedWidth =7800
                            LayoutCachedHeight =1860
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =660
                    Top =1020
                    Width =5700
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblChooseSpeciesType"
                    Caption ="Choose at least one species type or common name to search."
                    GridlineColor =10921638
                    LayoutCachedLeft =660
                    LayoutCachedTop =1020
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =1335
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =660
                    Top =2760
                    Width =6540
                    Height =360
                    ColumnOrder =6
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSearchFor"
                    DefaultValue ="\"\""
                    OnLostFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =660
                    LayoutCachedTop =2760
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =3120
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2280
                            Width =4476
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblSearchFor"
                            Caption ="Enter the name or portion of name to search for."
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =2280
                            LayoutCachedWidth =4836
                            LayoutCachedHeight =2580
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =7080
                    Top =3300
                    Width =1618
                    Height =373
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =29
                    RightMargin =29
                    BottomMargin =29
                    BackColor =13882323
                    BorderColor =8355711
                    ForeColor =8224125
                    Name ="lblSearch"
                    Caption ="Search!"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Search for species"
                    GridlineColor =10921638
                    LayoutCachedLeft =7080
                    LayoutCachedTop =3300
                    LayoutCachedWidth =8698
                    LayoutCachedHeight =3673
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =120
                    Top =4140
                    Width =2448
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSearchResults"
                    Caption ="Search Results"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =4140
                    LayoutCachedWidth =2568
                    LayoutCachedHeight =4512
                End
                Begin Label
                    OverlapFlags =215
                    Left =300
                    Top =4620
                    Width =6576
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSearchResultInstructions"
                    Caption ="Double click the species you'd like to add to your target species listing. "
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =4620
                    LayoutCachedWidth =6876
                    LayoutCachedHeight =4920
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =87
                    Top =4020
                    Width =9000
                    BorderColor =8355711
                    Name ="lineCurrTgtAreaTop"
                    GridlineColor =10921638
                    LayoutCachedTop =4020
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =4020
                    BorderTint =50.0
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =87
                    Top =5160
                    Width =9000
                    BorderColor =8355711
                    Name ="lineCurrTgtAreaBtm"
                    GridlineColor =10921638
                    LayoutCachedTop =5160
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =5160
                    BorderTint =50.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4740
                    Top =3300
                    Width =2220
                    TabIndex =7
                    ForeColor =16711680
                    Name ="btnSearch"
                    Caption ="Search!"
                    StatusBarText ="Add new target area"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4740
                    LayoutCachedTop =3300
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =3660
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
                End
            End
        End
        Begin Section
            Height =1020
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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
' MODULE:       Form_frmSpeciesSearch
' Description:  Species search functions & procedures
'
' Source/date:  Bonnie Campbell, 2/9/2015
' Revisions:    BLC - 2/9/2015 - initial version
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
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler
    
    Initialize
       
    'species type selections
    TempVars.Add "speciestype", ""
    
    'disable search until something is entered
    btnSearch.Enabled = False
    DisableControl btnSearch

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxSearchFor_LostFocus
' Description:  Actions to take when search for textbox is not empty
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
Private Sub tbxSearchFor_LostFocus()
On Error GoTo Err_Handler
    
    If Len(tbxSearchFor.Value) > 0 Then
        'check if species list is identified
        If Len(TempVars.item("speciestype")) > 0 Then
            'enable the search "button"
            btnSearch.Enabled = True
            EnableControl btnSearch, TempVars.item("ctrlAddEnabled"), TempVars.item("textEnabled")
        Else
            MsgBox "Please choose at least one species list to search.", vbOKOnly, "Oops! Missing Species List to Search"
        End If
    Else
        'disable the search "button"
        btnSearch.Enabled = False
        DisableControl btnSearch
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxSearchFor_LostFocus[Form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxCO_Click
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub cbxCO_Click()
On Error GoTo Err_Handler

If cbxCO = True Then
    'TempVars.Item("speciestype") = TempVars.Item("speciestype") & ";CO"
    cbxAddToList "speciestype", "CO", ";"

Else
    'TempVars.Item("speciestype") = Replace(Replace(TempVars.Item("speciestype"), "CO", ""), ";;", ";")
    cbxRemoveFromList "speciestype", "CO", ";"

End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxCO_Click[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxUT_Click
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub cbxUT_Click()
On Error GoTo Err_Handler

If cbxUT = True Then
    'TempVars.Item("speciestype") = TempVars.Item("speciestype") & ";UT"
    cbxAddToList "speciestype", "UT", ";"

Else
    'TempVars.Item("speciestype") = Replace(Replace(TempVars.Item("speciestype"), "UT", ""), ";;", ";")
    cbxRemoveFromList "speciestype", "UT", ";"

End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxUT_Click[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxWY_Click
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub cbxWY_Click()
On Error GoTo Err_Handler

If cbxWY = True Then
    'TempVars.Item("speciestype") = TempVars.Item("speciestype") & ";WY"
    cbxAddToList "speciestype", "WY", ";"

Else
    'TempVars.Item("speciestype") = Replace(Replace(TempVars.Item("speciestype"), "WY", ""), ";;", ";")
    cbxRemoveFromList "speciestype", "WY", ";"

End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxWY_Click[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxITIS_Click
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub cbxITIS_Click()
On Error GoTo Err_Handler

If cbxITIS = True Then
    'TempVars.Item("speciestype") = TempVars.Item("speciestype") & ";ITIS"
    cbxAddToList "speciestype", "ITIS", ";"

Else
    'TempVars.Item("speciestype") = Replace(Replace(TempVars.Item("speciestype"), "ITIS", ""), ";;", ";")
    cbxRemoveFromList "speciestype", "ITIS", ";"

End If

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
' SUB:          cbxCommon_Click
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub cbxCommon_Click()
On Error GoTo Err_Handler

If cbxCommon = True Then
    'TempVars.Item("speciestype") = TempVars.Item("speciestype") & ";CMN"
    cbxAddToList "speciestype", "CMN", ";"
Else
    'TempVars.Item("speciestype") = Replace(Replace(TempVars.Item("speciestype"), "CMN", ""), ";;", ";")
    cbxRemoveFromList "speciestype", "CMN", ";"
End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxCO_Click[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxAddToList
' Description:  Add an item to a list
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub cbxAddToList(list As String, cbxValue As String, separator As String)
On Error GoTo Err_Handler
    
    'if list exists and item is in it, exit
    If Len(TempVars.item(list)) > 0 Then
        If CountInString(TempVars.item(list), cbxValue) > 0 Then
            GoTo Exit_Sub
        End If
    End If
        
    'add item if it's not already in list
    TempVars.item(list) = TempVars.item(list) & cbxValue & separator
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxAddToList[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxRemoveFromList
' Description:  Remove an item from a list
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub cbxRemoveFromList(list As String, cbxValue As String, separator As String)
On Error GoTo Err_Handler
    
    TempVars.item(list) = Replace(Replace(TempVars.item(list), cbxValue, ""), separator & separator, separator)
    
    'clear if only = separator
    If Len(TempVars.item(list)) = 1 And TempVars.item(list) = separator Then TempVars.item(list) = ""

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxRemoveFromList[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lblSearch_Click
' Description:  Search for the name or portion of a name in the species/common names listed & return a result list
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' http://codevba.com/msaccess/status_bar_and_progress_meter.htm#.VNb9X_lM4_4
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015 - initial version
' ---------------------------------
Private Sub lblSearch_Click()
On Error GoTo Err_Handler
    
    Dim speciestype As Variant
    Dim strSearch As String, strSpecies As String, strWHERE As String, strSQL As String
    Dim i As Integer

    'ignore if disabled
    'If lblSearch.backColor = TempVars.item("ctrlDisabled") Then GoTo Exit_Sub
    If btnSearch.Enabled = False Then GoTo Exit_Sub

    strSearch = Trim(tbxSearchFor.Value)
            
    'check strSearch is alpha numeric
    
    'determine which areas are to be searched
    MsgBox TempVars.item("speciestype"), vbOKOnly, "speciestype"

    'perform search
    strWHERE = "WHERE "
    
    'determine which species names to check
    For Each speciestype In Split(TempVars.item("speciestype"), ";")
        
        'If CountInString(speciestype, ";") > 1 Then
        i = i + 1
        If i > 1 Then
            strWHERE = strWHERE & " OR "
        End If
    
        Select Case speciestype
            Case "CO"   'Colorado
                strSpecies = "CO_Species"
            Case "UT"   'Utah
                strSpecies = "UT_Species"
            Case "WY"   'Wyoming
                strSpecies = "WY_Species"
            Case "ITIS" 'Master
                strSpecies = "Master_Species"
            Case "CMN"  'Common
                strSpecies = "Master_Common_Name"
        End Select
                
        strWHERE = strWHERE & " " & strSpecies & " LIKE '*" & strSearch & "*'"

    Next
    
    MsgBox strWHERE, vbOKOnly, "strWHERE"
    'prep WHERE clause
    If Len(Replace(strWHERE, "WHERE", "")) = 0 Then strWHERE = ""
    
    'build SQL statement
    strSQL = "SELECT LU_Code, Master_Species, Utah_Species, CO_Species, WY_Species, " _
            & "Master_Common_Name" _
            & "FROM tlu_NCPN_Plants" _
            & strWHERE & ";"
            
    MsgBox strSQL, vbOKOnly, "strSQL"
            
            'set statusbar notice
            Dim varReturn As Variant
            varReturn = SysCmd(acSysCmdSetStatus, "Searching for " & strSearch & "...")
            
            'print it
            MsgBox "Results for search: " & vbCrLf & strSearch, vbOKOnly, "Search Results"

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblSearch_Click[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnSearch_Click
' Description:  Search for the name or portion of a name in the species/common names listed & return a result list
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' http://codevba.com/msaccess/status_bar_and_progress_meter.htm#.VNb9X_lM4_4
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015 - initial version
' ---------------------------------
Private Sub btnSearch_Click()
On Error GoTo Err_Handler
    
    Dim speciestype As Variant
    Dim strSearch As String, strSpecies As String, strWHERE As String, strSQL As String
    Dim i As Integer

    'ignore if disabled
    If btnSearch.Enabled = False Then GoTo Exit_Sub

    strSearch = Trim(tbxSearchFor.Value)
            
    'check strSearch is alpha numeric
    
    'check if species list is selected
    If Len(TempVars.item("speciestype")) > 0 Then
        'enable the search "button"
        btnSearch.Enabled = True
        EnableControl btnSearch, TempVars.item("ctrlAddEnabled"), TempVars.item("textEnabled")
    Else
        MsgBox "Please choose at least one species list to search.", vbOKOnly, "Oops! Missing Species List to Search"
        GoTo Exit_Sub
    End If
    
    
    'determine which areas are to be searched
    MsgBox TempVars.item("speciestype"), vbOKOnly, "speciestype"

    'perform search
    strWHERE = " WHERE "
    
    'determine which species names to check
    For Each speciestype In Split(TempVars.item("speciestype"), ";")
        
        If Len(speciestype) > 0 Then
            
            'If CountInString(speciestype, ";") > 1 Then
            i = i + 1
            If i > 1 Then
                strWHERE = strWHERE & " OR "
            End If
        
            Select Case speciestype
                Case "CO"   'Colorado
                    strSpecies = "CO_Species"
                Case "UT"   'Utah
                    strSpecies = "UT_Species"
                Case "WY"   'Wyoming
                    strSpecies = "WY_Species"
                Case "ITIS" 'Master
                    strSpecies = "Master_Species"
                Case "CMN"  'Common
                    strSpecies = "Master_Common_Name"
            End Select
                    
            strWHERE = strWHERE & " " & strSpecies & " LIKE '*" & strSearch & "*'"
            
        End If
    Next
    
    MsgBox strWHERE, vbOKOnly, "strWHERE"
    'prep WHERE clause
    If Len(Replace(strWHERE, "WHERE", "")) = 0 Then strWHERE = ""
    
    'build SQL statement
    strSQL = "SELECT LU_Code, Master_Species, Utah_Species, CO_Species, WY_Species, " _
            & "Master_Common_Name" _
            & "FROM tlu_NCPN_Plants" _
            & strWHERE & ";"
            
    MsgBox strSQL, vbOKOnly, "strSQL"
            
            'set statusbar notice
            Dim varReturn As Variant
            varReturn = SysCmd(acSysCmdSetStatus, "Searching for " & strSearch & "...")
            
            'print it
            MsgBox "Results for search: " & vbCrLf & strSearch, vbOKOnly, "Search Results"

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSearch_Click[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub
