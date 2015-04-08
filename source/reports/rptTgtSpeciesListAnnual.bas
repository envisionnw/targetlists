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
    Width =15660
    DatasheetFontHeight =11
    ItemSuffix =80
    Right =15600
    Bottom =7248
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x41aa626fb98ee440
    End
    RecordSource ="Query2"
    Caption ="INVASIVE LIST"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6a01000068010000680100006801000000000000fe2c0000b001000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            Height =1560
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
                    Left =2100
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
                    LayoutCachedLeft =2100
                    LayoutCachedTop =960
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =2
                    Left =4140
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
                    LayoutCachedLeft =4140
                    LayoutCachedTop =960
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =6300
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
                    LayoutCachedLeft =6300
                    LayoutCachedTop =960
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =240
                    Top =960
                    Width =1800
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
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =8040
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
                    LayoutCachedLeft =8040
                    LayoutCachedTop =960
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =2
                    Left =2100
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
                    LayoutCachedLeft =2100
                    LayoutCachedTop =600
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =900
                End
                Begin Line
                    Left =2100
                    Top =924
                    Width =3720
                    Name ="lnSpecies"
                    GridlineColor =10921638
                    LayoutCachedLeft =2100
                    LayoutCachedTop =924
                    LayoutCachedWidth =5820
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
                    Left =9840
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
                    Name ="lblBLCA"
                    Caption ="BLCA"
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
                    Name ="lblBRCA"
                    Caption ="BRCA"
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
                    Name ="lblCANY"
                    Caption ="CANY"
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
                    Name ="lblCARE"
                    Caption ="CARE"
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
                    Name ="lblCEBR"
                    Caption ="CEBR"
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
                    Name ="lblCOLM"
                    Caption ="COLM"
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
                    Name ="lblCURE"
                    Caption ="CURE"
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
                    Name ="lblDINO"
                    Caption ="DINO"
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
                    Name ="lblFOBU"
                    Caption ="FOBU"
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
                    Name ="lblGOSP"
                    Caption ="GOSP"
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
                    Name ="lblHOVE"
                    Caption ="HOVE"
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
                    Name ="lblNABR"
                    Caption ="NABR"
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
                    Name ="lblPISP"
                    Caption ="PISP"
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
                    Name ="lblTICA"
                    Caption ="TICA"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =14880
                    LayoutCachedTop =600
                    LayoutCachedWidth =15180
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =3
                    Left =15240
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
                    LayoutCachedLeft =15240
                    LayoutCachedTop =600
                    LayoutCachedWidth =15540
                    LayoutCachedHeight =1260
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =418
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Width =15660
                    Height =418
                    TabIndex =5
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
                    Left =2100
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

                    LayoutCachedLeft =2100
                    LayoutCachedTop =60
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4140
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

                    LayoutCachedLeft =4140
                    LayoutCachedTop =60
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6240
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

                    LayoutCachedLeft =6240
                    LayoutCachedTop =60
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7980
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

                    LayoutCachedLeft =7980
                    LayoutCachedTop =60
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Width =1800
                    Height =312
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFamily"
                    ControlSource ="Family"
                    StatusBarText ="Target Species name (ITIS species name from tlu_NCPN_Plants.Master_Species)"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =372
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9840
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxArchPriority"
                    ControlSource ="=ParkPriorities(0)"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =9840
                    LayoutCachedTop =60
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10200
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxBLCAPriority"
                    ControlSource ="=IIf([Park]=\"BLCA\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10200
                    LayoutCachedTop =60
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10560
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxBRCAPriority"
                    ControlSource ="=IIf([Park]=\"BRCA\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10560
                    LayoutCachedTop =60
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10920
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCANYPriority"
                    ControlSource ="=IIf([Park]=\"CANY\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10920
                    LayoutCachedTop =60
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11280
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCAREPriority"
                    ControlSource ="=IIf([Park]=\"CARE\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11280
                    LayoutCachedTop =60
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11640
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCEBRPriority"
                    ControlSource ="=IIf([Park]=\"CEBR\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11640
                    LayoutCachedTop =60
                    LayoutCachedWidth =11940
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12000
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCOLMPriority"
                    ControlSource ="=IIf([Park]=\"COLM\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12000
                    LayoutCachedTop =60
                    LayoutCachedWidth =12300
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12360
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCUREPriority"
                    ControlSource ="=IIf([Park]=\"CURE\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12360
                    LayoutCachedTop =60
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12720
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDINOPriority"
                    ControlSource ="=IIf([Park]=\"DINO\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12720
                    LayoutCachedTop =60
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13080
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFOBUPriority"
                    ControlSource ="=IIf([Park]=\"FOBU\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13080
                    LayoutCachedTop =60
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13440
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxGOSPPriority"
                    ControlSource ="=IIf([Park]=\"GOSP\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13440
                    LayoutCachedTop =60
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13800
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =17
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxHOVEPriority"
                    ControlSource ="=IIf([Park]=\"HOVE\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13800
                    LayoutCachedTop =60
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14160
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =18
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxNABRPriority"
                    ControlSource ="=IIf([Park]=\"NABR\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =14160
                    LayoutCachedTop =60
                    LayoutCachedWidth =14460
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14520
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =19
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPISPPriority"
                    ControlSource ="=IIf([Park]=\"PISP\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =14520
                    LayoutCachedTop =60
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14880
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =20
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxTICAPriority"
                    ControlSource ="=IIf([Park]=\"TICA\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =14880
                    LayoutCachedTop =60
                    LayoutCachedWidth =15180
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15240
                    Top =60
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =21
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxZIONPriority"
                    ControlSource ="=IIf([Park]=\"ZION\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =15240
                    LayoutCachedTop =60
                    LayoutCachedWidth =15540
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7320
                    Top =60
                    Width =600
                    Height =300
                    FontSize =9
                    TabIndex =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPriority"
                    ControlSource ="ParkPriorities"
                    StatusBarText ="FK to plant master code (tlu_NCPN_Plants.Master_Species)"
                    GridlineColor =10921638

                    LayoutCachedLeft =7320
                    LayoutCachedTop =60
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =360
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
            Height =1320
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
                    Top =720
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
                    LayoutCachedTop =720
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =1032
                End
                Begin Label
                    TextAlign =3
                    Left =7200
                    Top =720
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
                    LayoutCachedTop =720
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =1044
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
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9840
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text63"
                    ControlSource ="=IIf([Park]=\"ARCH\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =9840
                    LayoutCachedTop =120
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10200
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text64"
                    ControlSource ="=IIf([Park]=\"BLCA\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10200
                    LayoutCachedTop =120
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10560
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text65"
                    ControlSource ="=IIf([Park]=\"BRCA\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10560
                    LayoutCachedTop =120
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10920
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text66"
                    ControlSource ="=IIf([Park]=\"CANY\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10920
                    LayoutCachedTop =120
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11280
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text67"
                    ControlSource ="=IIf([Park]=\"CARE\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11280
                    LayoutCachedTop =120
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11640
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text68"
                    ControlSource ="=IIf([Park]=\"CEBR\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11640
                    LayoutCachedTop =120
                    LayoutCachedWidth =11940
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12000
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text69"
                    ControlSource ="=IIf([Park]=\"COLM\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12000
                    LayoutCachedTop =120
                    LayoutCachedWidth =12300
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12360
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text70"
                    ControlSource ="=IIf([Park]=\"CURE\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12360
                    LayoutCachedTop =120
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12720
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text71"
                    ControlSource ="=IIf([Park]=\"DINO\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12720
                    LayoutCachedTop =120
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13080
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text72"
                    ControlSource ="=IIf([Park]=\"FOBU\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13080
                    LayoutCachedTop =120
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13440
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text73"
                    ControlSource ="=IIf([Park]=\"GOSP\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13440
                    LayoutCachedTop =120
                    LayoutCachedWidth =13740
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13800
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text74"
                    ControlSource ="=IIf([Park]=\"HOVE\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13800
                    LayoutCachedTop =120
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14160
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text75"
                    ControlSource ="=IIf([Park]=\"NABR\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =14160
                    LayoutCachedTop =120
                    LayoutCachedWidth =14460
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14520
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text76"
                    ControlSource ="=IIf([Park]=\"PISP\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =14520
                    LayoutCachedTop =120
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =14880
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text77"
                    ControlSource ="=IIf([Park]=\"TICA\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =14880
                    LayoutCachedTop =120
                    LayoutCachedWidth =15180
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15240
                    Top =120
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text78"
                    ControlSource ="=IIf([Park]=\"ZION\",Switch([Transect_Only]=1,\"Transect\",Len([Tgt_Area])>0,[Tg"
                        "t_Area],[Priority]>0,[Priority],[Priority]=0,\"X\"),\"X\")"
                    StatusBarText ="Park priority (1 - , 2- , 3- , 4- , 5-)"
                    GridlineColor =10921638

                    LayoutCachedLeft =15240
                    LayoutCachedTop =120
                    LayoutCachedWidth =15540
                    LayoutCachedHeight =420
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

Dim ParkPriorities As Variant

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
    
    'set the background color if tbxPriority = "Transect Only" or a Target_Area vs. Priority #
    'use conditional formatting for tbxDetail:
    '   [tbxPriority] = "Transect Only"  >>  ltLime
    '   (Not IsNumeric[tbxPriority])) And ([tbxPriority] <> "Transect Only") >> ltYellow
    
    'prepare the Park Priority values
    ParkPriorities = Split([ParkPriority], "|")
    
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
