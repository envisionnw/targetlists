Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11400
    DatasheetFontHeight =11
    ItemSuffix =30
    Right =20136
    Bottom =9660
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xbf59acd5ff8de440
    End
    RecordSource ="qryParkTgtSpeciesLists"
    Caption ="INVASIVE LIST"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6a01000068010000680100006801000000000000fe2c0000b001000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    FitToPage =1
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
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            ControlSource ="tbl_Target_Species.Target_Year"
        End
        Begin BreakLevel
            ControlSource ="tbl_Target_Species.Park_Code"
        End
        Begin BreakLevel
            ControlSource ="Species_Name"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =948
            BackColor =15849926
            Name ="ReportHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =2892
                    Height =528
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblReportHdr"
                    Caption ="INVASIVES LIST"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =2952
                    LayoutCachedHeight =588
                End
            End
        End
        Begin PageHeader
            Height =1332
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    OldBorderStyle =0
                    Width =11400
                    Height =480
                    BackColor =15849926
                    BorderColor =10921638
                    Name ="rectPageHdr"
                    GridlineColor =10921638
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =480
                    BackThemeColorIndex =2
                    BackTint =20.0
                End
                Begin Label
                    TextAlign =2
                    Left =180
                    Top =960
                    Width =1800
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesNameUT"
                    Caption ="UT"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =960
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =2
                    Left =2220
                    Top =960
                    Width =1680
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesNameCO"
                    Caption ="CO"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =2220
                    LayoutCachedTop =960
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =4200
                    Top =960
                    Width =1380
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPlantCode"
                    Caption ="Plant Code"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4200
                    LayoutCachedTop =960
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =5880
                    Top =960
                    Width =1980
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblFamily"
                    Caption ="Family"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5880
                    LayoutCachedTop =960
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =8160
                    Top =960
                    Width =1680
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCommonName"
                    Caption ="Common Name"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =8160
                    LayoutCachedTop =960
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =10080
                    Top =960
                    Width =1200
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPriority"
                    Caption ="Priority"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10080
                    LayoutCachedTop =960
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =2
                    Left =180
                    Top =600
                    Width =3720
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesNames"
                    Caption ="Species Names"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =600
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =900
                End
                Begin Line
                    Left =180
                    Top =924
                    Width =3720
                    Name ="lnSpecies"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =924
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =924
                End
                Begin Line
                    BorderWidth =2
                    Left =180
                    Top =1320
                    Width =11100
                    Name ="lnHeader"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =1320
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =1320
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =5040
                    Height =312
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDate"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6240
                    Top =60
                    Width =5040
                    Height =312
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPage"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6240
                    LayoutCachedTop =60
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =372
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =432
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    OldBorderStyle =0
                    Width =11400
                    Height =420
                    BorderColor =10921638
                    Name ="rectDetail"
                    GridlineColor =10921638
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =60
                    Width =1800
                    Height =312
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbl_Target_Species.Target_Year"
                    ControlSource ="utah_species"
                    StatusBarText ="Year (4-digit)"
                    EventProcPrefix ="tbl_Target_Species_Target_Year"
                    GridlineColor =10921638

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2220
                    Top =60
                    Width =1680
                    Height =312
                    ColumnWidth =1170
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbl_Target_Species.Park_Code"
                    ControlSource ="Co_Species"
                    StatusBarText ="Standard park code (CANY, FOBU, etc.)"
                    EventProcPrefix ="tbl_Target_Species_Park_Code"
                    GridlineColor =10921638

                    LayoutCachedLeft =2220
                    LayoutCachedTop =60
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4200
                    Top =60
                    Width =1380
                    Height =312
                    ColumnWidth =2655
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Species_Name"
                    ControlSource ="Master_Plant_Code_FK"
                    StatusBarText ="Target Species name (ITIS species name from tlu_NCPN_Plants.Master_Species)"
                    GridlineColor =10921638

                    LayoutCachedLeft =4200
                    LayoutCachedTop =60
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8160
                    Top =60
                    Width =1680
                    Height =312
                    ColumnWidth =2400
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCommon"
                    ControlSource ="Master_Common_Name"
                    StatusBarText ="FK to plant master code (tlu_NCPN_Plants.Master_Species)"
                    GridlineColor =10921638

                    LayoutCachedLeft =8160
                    LayoutCachedTop =60
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10080
                    Top =60
                    Width =1140
                    Height =300
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPriority"
                    ControlSource ="=Switch([Transect_Only]=1,\"Transect Only\",Len([Tgt_Area])>0,[Tgt_Area],[Priori"
                        "ty]>0,[Priority])"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10080
                    LayoutCachedTop =60
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5880
                    Top =60
                    Width =1980
                    Height =312
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFamily"
                    ControlSource ="Family"
                    StatusBarText ="Target Species name (ITIS species name from tlu_NCPN_Plants.Master_Species)"
                    GridlineColor =10921638

                    LayoutCachedLeft =5880
                    LayoutCachedTop =60
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =372
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =540
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10020
                    Top =120
                    Width =1140
                    Height =312
                    FontSize =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumPriority1"
                    ControlSource ="=CDbl(Nz(Count([Priority]),0))"
                    StatusBarText ="Standard park code (CANY, FOBU, etc.)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10020
                    LayoutCachedTop =120
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =432
                End
                Begin Label
                    TextAlign =3
                    Left =7320
                    Top =120
                    Width =2700
                    Height =324
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblTotalNum"
                    Caption ="Total # Priority 1 Species ="
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =7320
                    LayoutCachedTop =120
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =444
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Width =11100
                    Name ="lnPageFooter"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedWidth =11160
                End
            End
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
' MODULE:       Form_frmLoadList
' Description:  Load species list to target species list functions and routines
'
' Source/date:  Bonnie Campbell, 3/5/2015
' Revisions:    BLC - 3/5/2015 - initial version
' =================================

' ---------------------------------
' SUB:          Report_Open
' Description:  Actions for when reports open
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/1/2015 - initial version
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)

On Error GoTo Err_Handler
'http://stackoverflow.com/questions/11477297/giving-an-alias-to-a-subquery-containing-a-join-in-access

    If Len(Me.OpenArgs) > 0 Then
        ' Bob Larsen, January 28, 2012
        ' https://social.msdn.microsoft.com/Forums/office/en-US/3e126484-112f-4854-a5c0-2e9ef48e02bc/how-to-change-recordsource-for-a-report-with-vba?forum=accessdev
        'set recordset to passed in SQL via OpenArgs
        'If Me.OpenArgs <> vbNullString Then
        'Me.Recordset = Me.OpenArgs
        ' dyDMA, Sept 8, 2008
        ' http://www.utteraccess.com/forum/Run-time-error-32585-t1710296.html
        '==> Run-time Error 32585: This feature is only available in an ADP
        '==> Only Access ADP's can use this method (assign report recordset @ run-time)
        '==> Not available for *.mdb or *.accdb's
        
        'set orderby
        Me.OrderBy = Me.OpenArgs
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[Report_rptTgtSpeciesList])"
    End Select
    Resume Exit_Sub
End Sub


' ---------------------------------
' SUB:          Report_Load
' Description:  Report loading actions
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/1/2015 - initial version
' ---------------------------------
Private Sub Report_Load()

On Error GoTo Err_Handler
Dim iNumRecords As Integer, i As Integer

iNumRecords = DCount("*", Me.RecordSource)
'MsgBox iNumRecords & " records.", vbInformation, "System Error"

    For i = 1 To iNumRecords

    'set the background color if tbxPriority = "Transect Only" or a Target_Area vs. Priority #
    If Not IsNumeric(Me.tbxPriority) Then
        If Me.tbxPriority = "Transect Only" Then
            'Me.Detail.backcolor = lngLtYellow 'light yellow
            Me.rectDetail.backcolor = lngLtYellow 'light yellow
        Else
            'Me.Detail.backcolor = lngLtLime   'light lime
            Me.rectDetail.backcolor = lngLtLime   'light lime
        End If
    End If
    
    Next
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Load[Report_rptTgtSpeciesList])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          Detail_Format
' Description:  Report detail format actions
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/1/2015 - initial version
' ---------------------------------
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

On Error GoTo Err_Handler

    'set the background color if tbxPriority = "Transect Only" or a Target_Area vs. Priority #
    If Not IsNumeric(Me!tbxPriority) Then
        If Me!tbxPriority = "Transect Only" Then
            'Me.Detail.backcolor = lngLtYellow 'light yellow
        Else
            'Me.Detail.backcolor = lngLtLime   'light lime
        End If
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Format[Report_rptTgtSpeciesList])"
    End Select
    Resume Exit_Sub
End Sub
