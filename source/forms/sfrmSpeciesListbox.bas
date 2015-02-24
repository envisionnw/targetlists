Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    ScrollBars =2
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9480
    DatasheetFontHeight =11
    ItemSuffix =11
    Left =420
    Top =1080
    Right =4116
    Bottom =5100
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x37771cae5d89e440
    End
    RecordSource ="SELECT Switch(tlu_NCPN_Plants.LU_Code Is Null,\" \",tlu_NCPN_Plants.LU_Code<>\"\""
        ",tlu_NCPN_Plants.LU_Code) AS Code, tlu_NCPN_Plants.Master_Species AS Species, tl"
        "u_NCPN_Plants.Master_PLANT_Code\015\012FROM tlu_NCPN_Plants;\015\012"
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
    OrderByOnLoad =0
    OrderByOnLoad =0
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
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin PageHeader
            DisplayWhen =1
            Height =1320
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Width =840
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblResultsHdr"
                    Caption ="Results"
                    GridlineColor =10921638
                    LayoutCachedWidth =840
                    LayoutCachedHeight =372
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =600
                    Width =4800
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblChooseSpeciesType"
                    Caption ="Double click the species to add it to your target list."
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =600
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =915
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =1020
                    Width =1440
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCodeHdr"
                    Caption ="Code"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =1020
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =1320
                End
                Begin Label
                    OverlapFlags =85
                    Left =1680
                    Top =1020
                    Width =2520
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesHdr"
                    Caption ="Species"
                    GridlineColor =10921638
                    LayoutCachedLeft =1680
                    LayoutCachedTop =1020
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =1320
                End
                Begin Line
                    OverlapFlags =85
                    Left =1620
                    Top =1020
                    Width =0
                    Height =299
                    Name ="lineHdrSeparator"
                    GridlineColor =10921638
                    LayoutCachedLeft =1620
                    LayoutCachedTop =1020
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =1319
                End
            End
        End
        Begin Section
            Height =300
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Height =300
                    FontSize =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCode"
                    ControlSource ="Code"
                    OnDblClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Width =2520
                    Height =300
                    FontSize =10
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSpecies"
                    ControlSource ="Species"
                    OnDblClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1680
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =300
                End
                Begin Line
                    OverlapFlags =85
                    Left =1620
                    Width =0
                    Height =299
                    Name ="lineListSeparator"
                    GridlineColor =10921638
                    LayoutCachedLeft =1620
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =299
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4320
                    Width =5160
                    Height =300
                    FontSize =10
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxMasterCode"
                    ControlSource ="Master_PLANT_Code"
                    OnDblClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4320
                    LayoutCachedWidth =9480
                    LayoutCachedHeight =300
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =360
            Name ="PageFooterSection"
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
' MODULE:       Form_sfrmSpeciesListbox
' Description:  Species selction functions & procedures
'               and for lists which exceed standard listbox capacity
'
' Source/date:  Bonnie Campbell, 2/18/2015
' Revisions:    BLC - 2/18/2015 - initial version
' =================================

' ---------------------------------
' SUB:          Form_Load
' Description:  Form loading routine
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 18, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/18/2015 - initial version
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler

    'initial data fill
    'fillList Forms("frmTgtSpecies"), Me, Forms("frmTgtSpecies")!lbxTgtSpecies
   ' fillList Me.Parent, Me.Parent!sfrmSpeciesListbox, Forms("frmTgtSpecies")!lbxTgtSpecies
    fillList Me.Parent, Me.Parent.Controls("sfrmSpeciesListbox"), Forms("frmTgtSpecies")!lbxTgtSpecies

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_sfrmSpeciesListbox])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxCode_DblClick
' Description:  Actions for clicking tbxCode
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 19, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/19/2015 - initial version
'   BLC - 2/23/2015 - added lblTgtSpeciesCount update
' ---------------------------------
Private Sub tbxCode_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    Dim item As String
    Dim lbx As ListBox
    
    'add components of item (code, species (UT or whatever), & ITIS) to listbox

    'prepare item for listbox value
    item = tbxCode & ";" & tbxSpecies & ";" & tbxMasterCode
    
    'check listbox for duplicate & skip if already present
    If IsListDuplicate(Forms("frmTgtSpecies").Controls("lbxTgtSpecies"), 2, tbxMasterCode) Then
        'duplicate, so exit
        GoTo Exit_Sub
    End If

    Set lbx = Forms("frmTgtSpecies").Controls("lbxTgtSpecies")
    
    With lbx
        'add item if not duplicate
        .AddItem item
    
        'update target species count
        Forms("frmTgtSpecies").Controls("lblTgtSpeciesCount").Caption = .ListCount - 1 & " species"
        
        'return to the species list
        DoCmd.Minimize
        Forms("frmTgtSpecies").SetFocus
    End With
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxCode_DblClick[Form_sfrmSpeciesListbox])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxSpecies_DblClick
' Description:  Actions for clicking tbxSpecies
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 19, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/19/2015 - initial version
'   BLC - 2/23/2015 - added lblTgtSpeciesCount update
' ---------------------------------
Private Sub tbxSpecies_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    Dim item As String
    Dim lbx As ListBox
    
    'add components of item (code, species (UT or whatever), & ITIS) to listbox

    'prepare item for listbox value
    item = tbxCode & ";" & tbxSpecies & ";" & tbxMasterCode
    
    'check listbox for duplicate & skip if already present
    If IsListDuplicate(Forms("frmTgtSpecies").Controls("lbxTgtSpecies"), 2, tbxMasterCode) Then
        'duplicate, so exit
        GoTo Exit_Sub
    End If

    Set lbx = Forms("frmTgtSpecies").Controls("lbxTgtSpecies")
    
    With lbx
        'add item if not duplicate
        .AddItem item
    
        'update target species count
        Forms("frmTgtSpecies").Controls("lblTgtSpeciesCount").Caption = .ListCount - 1 & " species"
    End With
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxSpecies_DblClick[Form_sfrmSpeciesListbox])"
    End Select
    Resume Exit_Sub
End Sub
