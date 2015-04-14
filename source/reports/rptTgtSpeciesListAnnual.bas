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
    Width =25860
    DatasheetFontHeight =11
    ItemSuffix =119
    Right =12390
    Bottom =9600
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x0f001b43ee8ee440
    End
    RecordSource ="Query2"
    Caption ="INVASIVE LIST"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x690100006801000068010000680100000000000004650000a201000001000000 ,
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
            Height =1380
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    OldBorderStyle =0
                    Width =15660
                    Height =480
                    BackColor =15849926
                    BorderColor =10921638
                    Name ="rectPageHdr"
                    GridlineColor =10921638
                    LayoutCachedWidth =15660
                    LayoutCachedHeight =480
                    BackThemeColorIndex =2
                    BackTint =20.0
                End
                Begin Label
                    TextAlign =2
                    Left =1620
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
                    LayoutCachedLeft =1620
                    LayoutCachedTop =960
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =2
                    Left =3660
                    Top =960
                    Width =1980
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesNameCO"
                    Caption ="CO"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =3660
                    LayoutCachedTop =960
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =5880
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
                    LayoutCachedLeft =5880
                    LayoutCachedTop =960
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =240
                    Top =960
                    Width =1200
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblFamily"
                    Caption ="Family"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =960
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =7560
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
                    LayoutCachedLeft =7560
                    LayoutCachedTop =960
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =2
                    Left =1620
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
                    LayoutCachedLeft =1620
                    LayoutCachedTop =600
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =900
                End
                Begin Line
                    Left =1620
                    Top =924
                    Width =3720
                    Name ="lnSpecies"
                    GridlineColor =10921638
                    LayoutCachedLeft =1620
                    LayoutCachedTop =924
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =924
                End
                Begin Line
                    BorderWidth =2
                    Left =180
                    Top =1320
                    Width =15420
                    Name ="lnHeader"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =1320
                    LayoutCachedWidth =15600
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
                    Left =10500
                    Top =60
                    Width =5040
                    Height =312
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPage"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    GridlineColor =10921638

                    LayoutCachedLeft =10500
                    LayoutCachedTop =60
                    LayoutCachedWidth =15540
                    LayoutCachedHeight =372
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =9480
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblArch"
                    Caption ="ARCH"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =9480
                    LayoutCachedTop =600
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =9840
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblBLCA"
                    Caption ="BLCA"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =9840
                    LayoutCachedTop =600
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =10200
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblBRCA"
                    Caption ="BRCA"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10200
                    LayoutCachedTop =600
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =10560
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCANY"
                    Caption ="CANY"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10560
                    LayoutCachedTop =600
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =10920
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCARE"
                    Caption ="CARE"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10920
                    LayoutCachedTop =600
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =11280
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCEBR"
                    Caption ="CEBR"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =11280
                    LayoutCachedTop =600
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =11640
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCOLM"
                    Caption ="COLM"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =11640
                    LayoutCachedTop =600
                    LayoutCachedWidth =11940
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =12000
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCURE"
                    Caption ="CURE"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =12000
                    LayoutCachedTop =600
                    LayoutCachedWidth =12300
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =12360
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblDINO"
                    Caption ="DINO"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =12360
                    LayoutCachedTop =600
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =12720
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblFOBU"
                    Caption ="FOBU"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =12720
                    LayoutCachedTop =600
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =13080
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblGOSP"
                    Caption ="GOSP"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =13080
                    LayoutCachedTop =600
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =13440
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblHOVE"
                    Caption ="HOVE"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =13440
                    LayoutCachedTop =600
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =13800
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblNABR"
                    Caption ="NABR"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =13800
                    LayoutCachedTop =600
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =14160
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPISP"
                    Caption ="PISP"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =14160
                    LayoutCachedTop =600
                    LayoutCachedWidth =14460
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =14520
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblTICA"
                    Caption ="TICA"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =14520
                    LayoutCachedTop =600
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =14880
                    Top =600
                    Width =300
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblZION"
                    Caption ="ZION"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =14880
                    LayoutCachedTop =600
                    LayoutCachedWidth =15180
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =2
                    Left =15240
                    Top =600
                    Width =1380
                    Height =660
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label84"
                    Caption ="# Parks Where Priority 1"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =15240
                    LayoutCachedTop =600
                    LayoutCachedWidth =16620
                    LayoutCachedHeight =1260
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =418
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8520
                    Top =60
                    Width =5280
                    Height =300
                    FontSize =9
                    TabIndex =21
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxAll"
                    ControlSource ="ParkPriorities"
                    StatusBarText ="FK to plant master code (tlu_NCPN_Plants.Master_Species)"
                    GridlineColor =10921638

                    LayoutCachedLeft =8520
                    LayoutCachedTop =60
                    LayoutCachedWidth =13800
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Width =15660
                    Height =418
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDetail"
                    ConditionalFormat = Begin
                        0x0100000028010000020000000100000000000000000000001e00000001000000 ,
                        0x00000000ccff660001000000000000001f000000630000000100000000000000 ,
                        0xffff990000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078005000720069006f0072006900740079005d003d0022005400 ,
                        0x720061006e00730065006300740020004f006e006c0079002200000000002800 ,
                        0x4e006f0074002000490073004e0075006d00650072006900630028005b007400 ,
                        0x620078005000720069006f0072006900740079005d0029002900200041006e00 ,
                        0x6400200028005b007400620078005000720069006f0072006900740079005d00 ,
                        0x3c003e0022005400720061006e00730065006300740020004f006e006c007900 ,
                        0x2200290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedWidth =15660
                    LayoutCachedHeight =418
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000ccff66001d0000005b00 ,
                        0x7400620078005000720069006f0072006900740079005d003d00220054007200 ,
                        0x61006e00730065006300740020004f006e006c00790022000000000000000000 ,
                        0x0000000000000000000000000001000000000000000100000000000000ffff99 ,
                        0x004300000028004e006f0074002000490073004e0075006d0065007200690063 ,
                        0x0028005b007400620078005000720069006f0072006900740079005d00290029 ,
                        0x00200041006e006400200028005b007400620078005000720069006f00720069 ,
                        0x00740079005d003c003e0022005400720061006e00730065006300740020004f ,
                        0x006e006c00790022002900000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            Width =705
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label34"
                            Caption ="Text33"
                            GridlineColor =10921638
                            LayoutCachedWidth =705
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1740
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

                    LayoutCachedLeft =1740
                    LayoutCachedTop =60
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3780
                    Top =60
                    Width =1980
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

                    LayoutCachedLeft =3780
                    LayoutCachedTop =60
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5880
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

                    LayoutCachedLeft =5880
                    LayoutCachedTop =60
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7500
                    Top =60
                    Width =1800
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

                    LayoutCachedLeft =7500
                    LayoutCachedTop =60
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Width =1320
                    Height =312
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFamily"
                    ControlSource ="Family"
                    StatusBarText ="Target Species name (ITIS species name from tlu_NCPN_Plants.Master_Species)"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9480
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxARCHPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"ARCH\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220041005200430048002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =9480
                    LayoutCachedTop =60
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x41005200430048002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9840
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxBLCAPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"BLCA\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220042004c00430041002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =9840
                    LayoutCachedTop =60
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x42004c00430041002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxBRCAPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"BRCA\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220042005200430041002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =10200
                    LayoutCachedTop =60
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x42005200430041002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10560
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCANYPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"CANY\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x2200430041004e0059002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =10560
                    LayoutCachedTop =60
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x430041004e0059002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10920
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCAREPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"CARE\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220043004100520045002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =10920
                    LayoutCachedTop =60
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x43004100520045002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11280
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCEBRPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"CEBR\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220043004500420052002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =11280
                    LayoutCachedTop =60
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x43004500420052002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11640
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCOLMPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"COLM\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220043004f004c004d002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =11640
                    LayoutCachedTop =60
                    LayoutCachedWidth =11940
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x43004f004c004d002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12000
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCUREPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"CURE\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220043005500520045002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =12000
                    LayoutCachedTop =60
                    LayoutCachedWidth =12300
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x43005500520045002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12360
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDINOPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"DINO\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x2200440049004e004f002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =12360
                    LayoutCachedTop =60
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x440049004e004f002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12720
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFOBUPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"FOBU\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220046004f00420055002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =12720
                    LayoutCachedTop =60
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x46004f00420055002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13080
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxGOSPPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"GOSP\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220047004f00530050002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =13080
                    LayoutCachedTop =60
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x47004f00530050002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13440
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =17
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxHOVEPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"HOVE\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220048004f00560045002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =13440
                    LayoutCachedTop =60
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x48004f00560045002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13800
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =18
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxNABRPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"NABR\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x22004e004100420052002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =13800
                    LayoutCachedTop =60
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x4e004100420052002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14160
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =19
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPISPPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"PISP\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220050004900530050002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =14160
                    LayoutCachedTop =60
                    LayoutCachedWidth =14460
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x50004900530050002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14520
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =20
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxTICAPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"TICA\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x220054004900430041002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =14520
                    LayoutCachedTop =60
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x54004900430041002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14880
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxZIONPriority"
                    ControlSource ="=PopulateSpeciesPriorities(\"ZION\",[tbxAll])"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x010000001a010000010000000100000000000000000000005c00000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x280043006f0075006e00740049006e0053007400720069006e00670028005b00 ,
                        0x5000610072006b005000720069006f007200690074006900650073005d002c00 ,
                        0x22005a0049004f004e002d003100220029002b00490049006600280043006f00 ,
                        0x75006e00740049006e0053007400720069006e00670028005b00500061007200 ,
                        0x6b005000720069006f007200690074006900650073005d002c0022007c002200 ,
                        0x29003e0030002c0032002c003000290029003d00310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =14880
                    LayoutCachedTop =60
                    LayoutCachedWidth =15180
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ccffff005b0000002800 ,
                        0x43006f0075006e00740049006e0053007400720069006e00670028005b005000 ,
                        0x610072006b005000720069006f007200690074006900650073005d002c002200 ,
                        0x5a0049004f004e002d003100220029002b00490049006600280043006f007500 ,
                        0x6e00740049006e0053007400720069006e00670028005b005000610072006b00 ,
                        0x5000720069006f007200690074006900650073005d002c0022007c0022002900 ,
                        0x3e0030002c0032002c003000290029003d003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =16500
                    Top =60
                    Width =660
                    Height =300
                    FontSize =9
                    TabIndex =23
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxRunSumPri1"
                    ControlSource ="=CountInString([ParkPriorities],1)"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =16500
                    LayoutCachedTop =60
                    LayoutCachedWidth =17160
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15480
                    Top =60
                    Width =660
                    Height =300
                    FontSize =9
                    TabIndex =24
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumSpeciesPri1"
                    ControlSource ="=CountInString([ParkPriorities],1)"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =15480
                    LayoutCachedTop =60
                    LayoutCachedWidth =16140
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =17880
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =25
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxARCH"
                    ControlSource ="=CountInString([ParkPriorities],\"ARCH-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =17880
                    LayoutCachedWidth =18120
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =18180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =26
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxBLCA"
                    ControlSource ="=CountInString([ParkPriorities],\"BLCA-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =18180
                    LayoutCachedWidth =18420
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =18480
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =27
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxBRCA"
                    ControlSource ="=CountInString([ParkPriorities],\"BRCA-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =18480
                    LayoutCachedWidth =18720
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =18780
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =28
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCANY"
                    ControlSource ="=CountInString([ParkPriorities],\"CANY-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =18780
                    LayoutCachedWidth =19020
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19080
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =29
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCARE"
                    ControlSource ="=CountInString([ParkPriorities],\"CARE-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =19080
                    LayoutCachedWidth =19320
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19380
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =30
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCEBR"
                    ControlSource ="=CountInString([ParkPriorities],\"CEBR-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =19380
                    LayoutCachedWidth =19620
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19680
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =31
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCOLM"
                    ControlSource ="=CountInString([ParkPriorities],\"COLM-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =19680
                    LayoutCachedWidth =19920
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =19980
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =32
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCURE"
                    ControlSource ="=CountInString([ParkPriorities],\"CURE-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =19980
                    LayoutCachedWidth =20220
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =20280
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =33
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDINO"
                    ControlSource ="=CountInString([ParkPriorities],\"DINO-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =20280
                    LayoutCachedWidth =20520
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =20580
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =34
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFOBU"
                    ControlSource ="=CountInString([ParkPriorities],\"FOBU-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =20580
                    LayoutCachedWidth =20820
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =20880
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =35
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxGOSP"
                    ControlSource ="=CountInString([ParkPriorities],\"GOSP-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =20880
                    LayoutCachedWidth =21120
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =21180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =36
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxHOVE"
                    ControlSource ="=CountInString([ParkPriorities],\"HOVE-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =21180
                    LayoutCachedWidth =21420
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =21480
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =37
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxNABR"
                    ControlSource ="=CountInString([ParkPriorities],\"NABR-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =21480
                    LayoutCachedWidth =21720
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =21780
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =38
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPISP"
                    ControlSource ="=CountInString([ParkPriorities],\"PISP-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =21780
                    LayoutCachedWidth =22020
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =22080
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =39
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxTICA"
                    ControlSource ="=CountInString([ParkPriorities],\"TICA-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =22080
                    LayoutCachedWidth =22320
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =22380
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =40
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxZION"
                    ControlSource ="=CountInString([ParkPriorities],\"ZION-1\")"
                    StatusBarText ="Park priority"
                    GridlineColor =10921638

                    LayoutCachedLeft =22380
                    LayoutCachedWidth =22620
                    LayoutCachedHeight =300
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
                    Top =960
                    Width =1140
                    Height =312
                    FontSize =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumPriority1"
                    ControlSource ="=[tbxRunSumPri1]"
                    StatusBarText ="Standard park code (CANY, FOBU, etc.)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10020
                    LayoutCachedTop =960
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =1272
                End
                Begin Label
                    TextAlign =3
                    Left =7200
                    Top =960
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
                    LayoutCachedLeft =7200
                    LayoutCachedTop =960
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =1284
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Width =15480
                    Name ="lnPageFooter"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedWidth =15540
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9480
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumARCH"
                    ControlSource ="=[tbxARCH]"
                    StatusBarText ="Total # priority 1 (ARCH)"
                    GridlineColor =10921638

                    LayoutCachedLeft =9480
                    LayoutCachedTop =60
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9840
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumBLCA"
                    ControlSource ="=[tbxBLCA]"
                    StatusBarText ="Total # priority 1 (BLCA)"
                    GridlineColor =10921638

                    LayoutCachedLeft =9840
                    LayoutCachedTop =60
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10200
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumBRCA"
                    ControlSource ="=[tbxBRCA]"
                    StatusBarText ="Total # priority 1 (BRCA)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10200
                    LayoutCachedTop =60
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10560
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumCANY"
                    ControlSource ="=[tbxCANY]"
                    StatusBarText ="Total # priority 1 (CANY)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10560
                    LayoutCachedTop =60
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10920
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumCARE"
                    ControlSource ="=[tbxCARE]"
                    StatusBarText ="Total # priority 1 (CARE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10920
                    LayoutCachedTop =60
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11280
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumCEBR"
                    ControlSource ="=[tbxCEBR]"
                    StatusBarText ="Total # priority 1 (CEBR)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11280
                    LayoutCachedTop =60
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11640
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumCOLM"
                    ControlSource ="=[tbxCOLM]"
                    StatusBarText ="Total # priority 1 (COLM)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11640
                    LayoutCachedTop =60
                    LayoutCachedWidth =11940
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12000
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumCURE"
                    ControlSource ="=[tbxCURE]"
                    StatusBarText ="Total # priority 1 (CURE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12000
                    LayoutCachedTop =60
                    LayoutCachedWidth =12300
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12360
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumDINO"
                    ControlSource ="=[tbxDINO]"
                    StatusBarText ="Total # priority 1 (DINO)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12360
                    LayoutCachedTop =60
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12720
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumFOBU"
                    ControlSource ="=[tbxFOBU]"
                    StatusBarText ="Total # priority 1 (FOBU)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12720
                    LayoutCachedTop =60
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13080
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumGOSP"
                    ControlSource ="=[tbxGOSP]"
                    StatusBarText ="Total # priority 1 (GOSP)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13080
                    LayoutCachedTop =60
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13440
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumHOVE"
                    ControlSource ="=[tbxHOVE]"
                    StatusBarText ="Total # priority 1 (HOVE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13440
                    LayoutCachedTop =60
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13800
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumNABR"
                    ControlSource ="=[tbxNABR]"
                    StatusBarText ="Total # priority 1 (NABR)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13800
                    LayoutCachedTop =60
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14160
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumPISP"
                    ControlSource ="=[tbxPISP]"
                    StatusBarText ="Total # priority 1 (PISP)"
                    GridlineColor =10921638

                    LayoutCachedLeft =14160
                    LayoutCachedTop =60
                    LayoutCachedWidth =14460
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14520
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumTICA"
                    ControlSource ="=[tbxTICA]"
                    StatusBarText ="Total # priority 1 (TICA)"
                    GridlineColor =10921638

                    LayoutCachedLeft =14520
                    LayoutCachedTop =60
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14880
                    Top =60
                    Width =300
                    Height =270
                    FontSize =9
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumZION"
                    ControlSource ="=[tbxZION]"
                    StatusBarText ="Total # priority 1 (ZION)"
                    GridlineColor =10921638

                    LayoutCachedLeft =14880
                    LayoutCachedTop =60
                    LayoutCachedWidth =15180
                    LayoutCachedHeight =330
                End
                Begin Label
                    TextAlign =3
                    Left =5880
                    Top =60
                    Width =3480
                    Height =324
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblParkPriorities"
                    Caption ="Total # Priority 1 Species by Park ="
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5880
                    LayoutCachedTop =60
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =384
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9480
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =17
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueARCH"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"ARCH-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (ARCH)"
                    GridlineColor =10921638

                    LayoutCachedLeft =9480
                    LayoutCachedTop =420
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9840
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =18
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueBLCA"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"BLCA-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (BLCA)"
                    GridlineColor =10921638

                    LayoutCachedLeft =9840
                    LayoutCachedTop =420
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10200
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =19
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueBRCA"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"BRCA-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (BRCA)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10200
                    LayoutCachedTop =420
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10560
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =20
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueCANY"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"CANY-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (CANY)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10560
                    LayoutCachedTop =420
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10920
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =21
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueCARE"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"CARE-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (CARE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10920
                    LayoutCachedTop =420
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11280
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueCEBR"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"CEBR-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (CEBR)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11280
                    LayoutCachedTop =420
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11640
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =23
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueCOLM"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"COLM-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (COLM)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11640
                    LayoutCachedTop =420
                    LayoutCachedWidth =11940
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12000
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =24
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueCURE"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"CURE-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (CURE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12000
                    LayoutCachedTop =420
                    LayoutCachedWidth =12300
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12360
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =25
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueDINO"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"DINO-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (DINO)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12360
                    LayoutCachedTop =420
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12720
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =26
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueFOBU"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"FOBU-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (FOBU)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12720
                    LayoutCachedTop =420
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13080
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =27
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueGOSP"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"GOSP-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (GOSP)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13080
                    LayoutCachedTop =420
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13440
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =28
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueHOVE"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"HOVE-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (HOVE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13440
                    LayoutCachedTop =420
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13800
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =29
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueNABR"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"NABR-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (NABR)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13800
                    LayoutCachedTop =420
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14160
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =30
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniquePISP"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"PISP-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (PISP)"
                    GridlineColor =10921638

                    LayoutCachedLeft =14160
                    LayoutCachedTop =420
                    LayoutCachedWidth =14460
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14520
                    Top =420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =31
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueTICA"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"TICA-1\"),0))"
                    StatusBarText ="Total # unqiue priority 1 (TICA)"
                    GridlineColor =10921638

                    LayoutCachedLeft =14520
                    LayoutCachedTop =420
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14880
                    Top =420
                    Width =300
                    Height =270
                    FontSize =9
                    TabIndex =32
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueZION"
                    ControlSource ="=Sum(IIf(CountInString([ParkPriorities],\"1\")=1,CountInString([ParkPriorities],"
                        "\"ZION-1\"),0))"
                    StatusBarText ="Total # unique priority 1 (ZION)"
                    GridlineColor =10921638

                    LayoutCachedLeft =14880
                    LayoutCachedTop =420
                    LayoutCachedWidth =15180
                    LayoutCachedHeight =690
                End
                Begin Label
                    TextAlign =3
                    Left =5880
                    Top =420
                    Width =3480
                    Height =324
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label118"
                    Caption ="Unique Priority 1 Species ="
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5880
                    LayoutCachedTop =420
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =744
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
' MODULE:       Report_rptTgtSpeciesListAnnual
' Description:  Load species list to target species list functions and routines
'
' Source/date:  Bonnie Campbell, 4/7/2015
' Revisions:    BLC - 4/7/2015 - initial version
' =================================

' ---------------------------------
' SUB:          Report_Open
' Description:  Actions for when report opens
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, April 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/7/2015 - initial version
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)

On Error GoTo Err_Handler
'http://stackoverflow.com/questions/11477297/giving-an-alias-to-a-subquery-containing-a-join-in-access

'Dim ParkPriorities As Variant

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
            "Error encountered (#" & Err.Number & " - Report_Open[Report_rptTgtSpeciesListAnnual])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          Report_Load
' Description:  Actions for when report is loaded
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, April 8, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/8/2015 - initial version
' ---------------------------------
Private Sub Report_Load()

On Error GoTo Err_Handler

    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Load[Report_rptTgtSpeciesListAnnual])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          Detail_Format
' Description:  Actions for when report detail is formatted
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, April 8, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/8/2015 - initial version
' ---------------------------------
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

On Error GoTo Err_Handler

    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Format[Report_rptTgtSpeciesListAnnual])"
    End Select
    Resume Exit_Sub
End Sub
