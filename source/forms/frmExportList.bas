Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7215
    DatasheetFontHeight =11
    ItemSuffix =33
    Right =15720
    Bottom =11760
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xcecbc43e9089e440
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
    SplitFormSplitterBar =0
    SplitFormSplitterBar =0
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
            Height =4320
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =1776
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblChooseListHdr"
                    Caption ="Export List"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1836
                    LayoutCachedHeight =432
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4620
                    Top =3720
                    Width =2220
                    ForeColor =16711680
                    Name ="btnContinue"
                    Caption ="Continue >>"
                    StatusBarText ="Continue to create target list"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4620
                    LayoutCachedTop =3720
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =4080
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
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =2880
                    Left =2100
                    Top =1920
                    Height =300
                    ColumnOrder =2
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4138256
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="cbxYear"
                    ControlSource ="Target_Year"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tbl_Target_Species.Target_Year FROM tbl_Target_Species ORDER BY "
                        "tbl_Target_Species.Target_Year DESC; "
                    ColumnWidths ="1440"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =2100
                    LayoutCachedTop =1920
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =2220
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =840
                            Top =1920
                            Width =825
                            Height =315
                            FontWeight =500
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblAction"
                            Caption ="Year(s)"
                            GridlineColor =10921638
                            LayoutCachedLeft =840
                            LayoutCachedTop =1920
                            LayoutCachedWidth =1665
                            LayoutCachedHeight =2235
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =2880
                    Left =2100
                    Top =1440
                    Height =300
                    ColumnOrder =0
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4138256
                    ColumnInfo ="\"\";\"\";\"10\";\"8\""
                    Name ="cbxPark"
                    ControlSource ="Park_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tbl_Target_Species.Park_Code FROM tbl_Target_Species ORDER BY tb"
                        "l_Target_Species.Park_Code; "
                    ColumnWidths ="1440"
                    GridlineColor =10921638

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =2100
                    LayoutCachedTop =1440
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =1740
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =840
                            Top =1440
                            Width =825
                            Height =315
                            FontWeight =500
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblPark"
                            Caption ="Park(s)"
                            GridlineColor =10921638
                            LayoutCachedLeft =840
                            LayoutCachedTop =1440
                            LayoutCachedWidth =1665
                            LayoutCachedHeight =1755
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =540
                    Top =660
                    Width =4680
                    Height =585
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label21"
                    Caption ="Choose the park(s), year(s), and export format. Then click continue to export yo"
                        "ur desired list."
                    GridlineColor =10921638
                    LayoutCachedLeft =540
                    LayoutCachedTop =660
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =1245
                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =600
                    Top =2520
                    Width =4686
                    Height =943
                    ColumnOrder =1
                    TabIndex =3
                    BorderColor =10921638
                    Name ="Frame22"
                    DefaultValue ="1"
                    GridlineColor =10921638

                    LayoutCachedLeft =600
                    LayoutCachedTop =2520
                    LayoutCachedWidth =5286
                    LayoutCachedHeight =3463
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =720
                            Top =2400
                            Width =1575
                            Height =315
                            FontWeight =500
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblFormats"
                            Caption ="Export Format"
                            GridlineColor =10921638
                            LayoutCachedLeft =720
                            LayoutCachedTop =2400
                            LayoutCachedWidth =2295
                            LayoutCachedHeight =2715
                            BackThemeColorIndex =-1
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =780
                            Top =2848
                            OptionValue =1
                            BorderColor =10921638
                            Name ="optPDF"
                            GridlineColor =10921638

                            LayoutCachedLeft =780
                            LayoutCachedTop =2848
                            LayoutCachedWidth =1040
                            LayoutCachedHeight =3088
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1010
                                    Top =2820
                                    Width =435
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="lblPDF"
                                    Caption ="PDF"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =1010
                                    LayoutCachedTop =2820
                                    LayoutCachedWidth =1445
                                    LayoutCachedHeight =3135
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =1920
                            Top =2848
                            OptionValue =2
                            BorderColor =10921638
                            Name ="optXLS"
                            GridlineColor =10921638

                            LayoutCachedLeft =1920
                            LayoutCachedTop =2848
                            LayoutCachedWidth =2180
                            LayoutCachedHeight =3088
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =2150
                                    Top =2820
                                    Width =390
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="lblXLS"
                                    Caption ="XLS"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =2150
                                    LayoutCachedTop =2820
                                    LayoutCachedWidth =2540
                                    LayoutCachedHeight =3135
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =3060
                            Top =2850
                            OptionValue =3
                            BorderColor =10921638
                            Name ="optCSV"
                            GridlineColor =10921638

                            LayoutCachedLeft =3060
                            LayoutCachedTop =2850
                            LayoutCachedWidth =3320
                            LayoutCachedHeight =3090
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =3290
                                    Top =2820
                                    Width =435
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="lblCSV"
                                    Caption ="CSV"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =3290
                                    LayoutCachedTop =2820
                                    LayoutCachedWidth =3725
                                    LayoutCachedHeight =3135
                                End
                            End
                        End
                    End
                End
            End
        End
        Begin Section
            Height =20
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
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
' MODULE:       Form_frmSelectYear
' Description:  Action functions & procedures
'
' Source/date:  Bonnie Campbell, 2/23/2015
' Revisions:    BLC - 2/23/2015 - initial version
' =================================

' ---------------------------------
' SUB:          Form_Load
' Description:  Select year form loading processes
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015 - initial version
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler
    
    ' close select action form
    DoCmd.Close acForm, "frmActions"
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[form_frmSelectYear])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          cbxYear_Change
' Description:  Actions to take when a task action is selected
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015 - initial version
' ---------------------------------
Private Sub cbxYear_Change()
On Error GoTo Err_Handler
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxYear_Change[Form_frmSelectYear])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          btnContinue_Click
' Description:  Continue to specific actions
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015 - initial version
' ---------------------------------
Private Sub btnContinue_Click()
On Error GoTo Err_Handler
    
    Dim frm As String
    
    If Len(cbxYear.Value) > 0 Then
        
        'determine the selected action
        Select Case cbxYear.Value
            
            Case "SEL"  'default (non-select option)
                MsgBox "Please select a task.", vbOKOnly, "Oops! Missing Action"
                
                GoTo Exit_Sub
                
            Case "2013", "2014", "2015", "2016" 'year options
                'reference via AllForms to include closed forms
                frm = "frmTgtSpecies"
        
        End Select
        
        'open form
        DoCmd.OpenForm frm, acNormal, , , , , cbxYear.Value
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnContinue_Click[Form_frmSelectPark])"
    End Select
    Resume Exit_Sub
End Sub
