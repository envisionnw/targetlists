﻿Version =20
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
    ItemSuffix =15
    Left =420
    Top =1080
    Right =4110
    Bottom =5100
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x6036d90f9c8ae440
    End
    RecordSource ="SELECT Switch(tlu_NCPN_Plants.LU_Code Is Null,\" \",tlu_NCPN_Plants.LU_Code<>\"\""
        ",tlu_NCPN_Plants.LU_Code) AS Code, tlu_NCPN_Plants.Master_Species AS Species, tl"
        "u_NCPN_Plants.Master_PLANT_Code\015\012FROM tlu_NCPN_Plants;\015\012"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnKeyDown ="[Event Procedure]"
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
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7200
                    Top =960
                    Width =1800
                    Height =300
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCurrentRecord"
                    GridlineColor =10921638

                    LayoutCachedLeft =7200
                    LayoutCachedTop =960
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =1260
                End
            End
        End
        Begin Section
            Height =300
            Name ="Detail"
            OnPaint ="[Event Procedure]"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Height =300
                    FontSize =10
                    BackColor =9699294
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCode"
                    ControlSource ="Code"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =300
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
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
                    OnClick ="[Event Procedure]"
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
                    BackColor =9699294
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxMasterCode"
                    ControlSource ="Master_PLANT_Code"
                    OnDblClick ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4320
                    LayoutCachedWidth =9480
                    LayoutCachedHeight =300
                    BackThemeColorIndex =-1
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

'=================================================================
'  Declarations
'=================================================================
Dim curID As String 'Integer

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
' SUB:          Form_Current
' Description:  Actions for current detail record
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Rabbit July 11, 2011
' http://bytes.com/topic/access/answers/914781-set-colour-current-record
' March 6, 2010
' http://www.upsizing.co.uk/Art53_Highlight.aspx
' Adapted:      Bonnie Campbell, March 4, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/4/2015 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

    'set selected record ID
    curID = Me.tbxMasterCode 'Nz(Me.tbxMasterCode, 0)
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[Form_sfrmSpeciesListbox])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          Detail_Paint
' Description:  Actions for clicking tbxCode
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Rabbit July 11, 2011
' http://bytes.com/topic/access/answers/914781-set-colour-current-record
' Adapted:      Bonnie Campbell, March 4, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/4/2015 - initial version
' ---------------------------------
Private Sub Detail_Paint()
On Error GoTo Err_Handler

    'set selected record backcolor
    If Me.tbxMasterCode = curID Then
        Me.Detail.backcolor = lngYelLime
        Me.tbxCode.backcolor = lngYelLime
        'Me.tbxSpecies.backcolor = lngYelLime
        Me.tbxMasterCode.backcolor = lngYelLime
        
    Else
        Me.Detail.backcolor = lngWhite
        'Me.tbxCode.backcolor = lngWhite
    End If
       
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Paint[Form_sfrmSpeciesListbox])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxCode_Click
' Description:  Actions for clicking tbxCode
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 4, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/19/2015 - initial version
' ---------------------------------
Private Sub tbxCode_Click()
On Error GoTo Err_Handler

    'set selected record ID
    curID = Me.tbxMasterCode
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxCode_Click[Form_sfrmSpeciesListbox])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxSpecies_Click
' Description:  Actions for clicking tbxSpecies
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 4, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/4/2015 - initial version
' ---------------------------------
Private Sub tbxSpecies_Click()
On Error GoTo Err_Handler

    'set selected record ID
    curID = Me.tbxMasterCode 'Nz(Me.tbxMasterCode, 0)
       
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxSpecies_Click[Form_sfrmSpeciesListbox])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          tbxMasterCode_Click
' Description:  Actions for clicking tbxMasterCode
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015 - initial version
' ---------------------------------
Private Sub tbxMasterCode_Click()
On Error GoTo Err_Handler

    'set selected record ID
    curID = Me.tbxMasterCode

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxMasterCode_Click[Form_sfrmSpeciesListbox])"
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
    
    'check listbox for duplicate & skip if already present (col 0 vs 2)
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

' ---------------------------------
' SUB:          Form_KeyDown
' Description:  Respond to Up/Down in a continuous form by moving to next record
' Assumptions:  Active control's EnterKeyBehaviro is OFF
' Parameters:   frm - form for key behavior
'               KeyCode - code for key being pressed (integer)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Allen Browne via Jeanette Cunningham, Apr 13, 2010
' http://www.pcreview.co.uk/threads/need-to-get-the-up-down-arrow-keys-working.3995845/
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015  - initial version
' ---------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

    Call ContinuousUpDown(Me, KeyCode)
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_KeyDown[Form_sfrmSpeciesListbox])"
    End Select
    Resume Exit_Sub
End Sub
