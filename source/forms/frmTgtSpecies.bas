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
    Width =10935
    DatasheetFontHeight =11
    ItemSuffix =20
    Right =14508
    Bottom =9408
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x72574db34b86e440
    End
    OnClose ="[Event Procedure]"
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =11220
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =8340
                    Top =5220
                    Width =1618
                    Height =373
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =29
                    RightMargin =29
                    BottomMargin =29
                    BackColor =6750105
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblSaveList"
                    Caption ="Save List"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Save list"
                    GridlineColor =10921638
                    LayoutCachedLeft =8340
                    LayoutCachedTop =5220
                    LayoutCachedWidth =9958
                    LayoutCachedHeight =5593
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =5760
                    Top =1080
                    Width =3960
                    Height =4032
                    FontSize =10
                    BoundColumn =2
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxTgtSpecies"
                    RowSourceType ="Value List"
                    RowSource ="Code;Species;Master_PLANT_Code;''"
                    ColumnWidths ="1440;2520;720;14"
                    OnDblClick ="[Event Procedure]"
                    OnKeyUp ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Target species"
                    GridlineColor =10921638

                    LayoutCachedLeft =5760
                    LayoutCachedTop =1080
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =5112
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5580
                            Top =720
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblTgtSpecies"
                            Caption ="Target Species"
                            GridlineColor =10921638
                            LayoutCachedLeft =5580
                            LayoutCachedTop =720
                            LayoutCachedWidth =7020
                            LayoutCachedHeight =1035
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4757
                    Top =3780
                    Width =493
                    Height =463
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =101
                    RightMargin =29
                    BottomMargin =29
                    BackColor =13882323
                    BorderColor =8355711
                    ForeColor =8224125
                    Name ="lblRemove"
                    Caption ="<"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Remove selected"
                    GridlineColor =10921638
                    LayoutCachedLeft =4757
                    LayoutCachedTop =3780
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =4243
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4757
                    Top =1200
                    Width =493
                    Height =463
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =101
                    RightMargin =29
                    BottomMargin =29
                    BackColor =6750105
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblAddAll"
                    Caption =">>"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add all"
                    GridlineColor =10921638
                    LayoutCachedLeft =4757
                    LayoutCachedTop =1200
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =1663
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4757
                    Top =4560
                    Width =493
                    Height =463
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =101
                    RightMargin =29
                    BottomMargin =29
                    BackColor =52479
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblRemoveAll"
                    Caption ="<<"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Remove all"
                    GridlineColor =10921638
                    LayoutCachedLeft =4757
                    LayoutCachedTop =4560
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =5023
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4757
                    Top =2040
                    Width =493
                    Height =463
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =101
                    RightMargin =29
                    BottomMargin =29
                    BackColor =13882323
                    BorderColor =8355711
                    ForeColor =8224125
                    Name ="lblAdd"
                    Caption =">"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add selected"
                    GridlineColor =10921638
                    LayoutCachedLeft =4757
                    LayoutCachedTop =2040
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =2503
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4410
                    Top =5220
                    Width =1264
                    Height =358
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =29
                    RightMargin =29
                    BottomMargin =29
                    BackColor =6750105
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblReset"
                    Caption ="Reset Lists"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Reset lists to their original state"
                    GridlineColor =10921638
                    LayoutCachedLeft =4410
                    LayoutCachedTop =5220
                    LayoutCachedWidth =5674
                    LayoutCachedHeight =5578
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =1425
                    Top =5220
                    Width =1468
                    Height =373
                    FontWeight =600
                    LeftMargin =29
                    TopMargin =29
                    RightMargin =29
                    BottomMargin =29
                    BackColor =6750105
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblSearch"
                    Caption ="Find Species"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find a species..."
                    GridlineColor =10921638
                    LayoutCachedLeft =1425
                    LayoutCachedTop =5220
                    LayoutCachedWidth =2893
                    LayoutCachedHeight =5593
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
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
                    Name ="lblParkHdr"
                    Caption ="ARCH"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =900
                    LayoutCachedHeight =432
                End
                Begin Subform
                    OverlapFlags =85
                    Left =420
                    Top =1080
                    Width =3960
                    Height =4032
                    TabIndex =1
                    BorderColor =10921638
                    Name ="sfrmSpeciesListbox"
                    SourceObject ="Form.sfrmSpeciesListbox"
                    GridlineColor =10921638

                    LayoutCachedLeft =420
                    LayoutCachedTop =1080
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =5112
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =660
                            Width =1200
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblSpeciesListbox"
                            Caption ="Species"
                            GridlineColor =10921638
                            LayoutCachedLeft =180
                            LayoutCachedTop =660
                            LayoutCachedWidth =1380
                            LayoutCachedHeight =975
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =8340
                    Top =780
                    Width =1440
                    Height =276
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblTgtSpeciesCount"
                    Caption ="0 species"
                    ControlTipText ="Number of species in the current list"
                    GridlineColor =10921638
                    LayoutCachedLeft =8340
                    LayoutCachedTop =780
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =1056
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =3000
                    Top =780
                    Width =1440
                    Height =276
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSfrmSpeciesCount"
                    Caption ="3195 species"
                    ControlTipText ="Number of species in the current list"
                    GridlineColor =10921638
                    LayoutCachedLeft =3000
                    LayoutCachedTop =780
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =1056
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =2880
                    Top =120
                    Width =4320
                    Height =315
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear"
                    Caption =" "
                    GridlineColor =10921638
                    LayoutCachedLeft =2880
                    LayoutCachedTop =120
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =435
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
' MODULE:       Form_frmTgtSpecies
' Description:  Species selction functions & procedures
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
        
    'set state
    TempVars.item("state") = getParkState(TempVars.item("park"))
    
    'prep headers
    lblParkHdr.Caption = TempVars.item("park")
    'lblSpecies.Caption = TempVars.item("state") & " Species"
    lblYear.Caption = "Target Species List for " & Form.OpenArgs
    lblSpeciesListbox.Caption = TempVars.item("state") & " Species"
    
    'clear headers
    'lbxSpecies.RowSource = ""
    lbxTgtSpecies.RowSource = ""
    
    'initial listbox fill
'    fillList Me, lbxSpecies, lbxTgtSpecies  '.RowSourceType = "Value List"
     fillList Me, lbxTgtSpecies

    'Enable move items lbls (or not)
    'If lbxSpecies.ListCount > 0 Then
        lblAddAll.Visible = True
        lblRemoveAll.Visible = True
    'End If
    
    'Set counts
    'lblSpeciesCount.Caption = lbxSpecies.ListCount & " species"
    lblTgtSpeciesCount.Caption = lbxTgtSpecies.ListCount - 1 & " species"
'    lblLbxSpeciesCount.Caption = Count(sfrm.ListCount) & " species"
    
    DisableControl lblAdd
    DisableControl lblRemove
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lblReset_Click
' Description:  XX
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
Private Sub lblReset_Click()
On Error GoTo Err_Handler

    'go back to initial state
    Form_Load

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblReset_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxTgtSpecies_Click
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Allen Browne, June 2006
' http://allenbrowne.com/func-12.html
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
' ---------------------------------
Private Sub lbxTgtSpecies_Click()
On Error GoTo Err_Handler
    
    Dim varItem As Variant
    
    'deselect items in source control (lbxSpecies)
    'For Each varItem In lbxSpecies.ItemsSelected
    '    lbxSpecies.Selected(varItem) = False
    'Next

    'check for selected items --> if present, enable lblRemove
    If lbxTgtSpecies.ItemsSelected.Count > 0 Then
        If lblRemove.backcolor <> TempVars.item("ctrlRemoveEnabled") Then
            EnableControl lblRemove, TempVars.item("ctrlRemoveEnabled"), TempVars.item("textEnabled")
        End If
    Else
        DisableControl lblRemove
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxTgtSpecies_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxTgtSpecies_DblClick
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
Private Sub lbxTgtSpecies_DblClick(Cancel As Integer)
    
On Error GoTo Err_Handler

    MoveSingleItem Me, "lbxTgtSpecies", "lbxSpecies"

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxTgtSpecies_DblClick[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxTgtSpecies_KeyUp
' Description:  Deselects items after control update
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
Private Sub lbxTgtSpecies_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

'    If lbxSpecies.ItemsSelected.Count > 0 And lblRemove.backcolor <> TempVars.item("ctrlRemoveEnabled") Then
    If lblRemove.backcolor <> TempVars.item("ctrlRemoveEnabled") Then
        EnableControl lblRemove, TempVars.item("ctrlRemoveEnabled"), TempVars.item("textEnabled")
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxTgtSpecies_KeyUp[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lblAdd_Click
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
Private Sub lblAdd_Click()
On Error GoTo Err_Handler
    
    'ignore if 'disabled'
    If lblAdd.backcolor = lngGray Then GoTo Exit_Sub
    
    MoveSingleItem Me, "lbxSpecies", "lbxTgtSpecies"

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblAdd_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lblRemove_Click
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
Private Sub lblRemove_Click()
On Error GoTo Err_Handler
    
    'ignore if 'disabled'
    If lblRemove.backcolor = TempVars.item("ctrlDisabled") Then GoTo Exit_Sub
    
    MoveSingleItem Me, "lbxTgtSpecies", "lbxSpecies"
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblRemove_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lblAddAll_Click
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
Private Sub lblAddAll_Click()
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'fetch recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset(TempVars.item("strSQL"))
    
    MoveAllItems Me, "lbxSpecies", "lbxTgtSpecies"

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblAddAll_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lblRemoveAll_Click
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
Private Sub lblRemoveAll_Click()
On Error GoTo Err_Handler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'fetch recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset(TempVars.item("strSQL"))
    
    MoveAllItems Me, "lbxTgtSpecies", "lbxSpecies"

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblRemoveAll_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lblSaveList_Click
' Description:  Save list items
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015 - initial version
' ---------------------------------
Private Sub lblSaveList_Click()
On Error GoTo Err_Handler

    Dim iRow As Integer, i As Integer
    Dim strMasterCode As String, strSpecies As String, strSQL As String, strInsert As String
    Dim varReturn As Variant
    
    'start @ row 1 (headers = row 0)
    For iRow = 1 To lbxTgtSpecies.ListCount - 1
       
       ' ---------------------------------------------------
       '  NOTE: listbox column MUST have a non-zero width to retrieve its value
       ' ---------------------------------------------------
        strMasterCode = lbxTgtSpecies.Column(0, iRow) 'column 0 = Master_PLANT_Code
        strSpecies = lbxTgtSpecies.Column(1, iRow) 'column 1 = Species name
        
       ' ---------------------------------------------------
       '  Check if item exists in tbl_TgtSpecies for Park, Year, Species combo
       ' ---------------------------------------------------
        strSQL = "SELECT * FROM tbl_Target_Species " & _
                 "WHERE Master_PLANT_Code_FK = '" & strMasterCode & _
                 " ' AND Park_Code = '" & TempVars.item("park") & _
                 " ' AND Target_Year = " & TempVars.item("TgtYear") & ";"
        
        Dim db As DAO.Database
        Dim rs As DAO.Recordset

        Set rs = CurrentDb.OpenRecordset(strSQL) 'CurrentDb.Execute(strSQL, dbFailOnError) >> doesn't compile expected function or variable
        
        'check if there are no records (rs.BOF & rs.EOF are both true)
        If rs.BOF And rs.EOF Then
            
            'set statusbar notice
            varReturn = SysCmd(acSysCmdSetStatus, "Saving " & strSpecies & "...")
            
            'prepare SQL
            strSQL = "INSERT INTO tbl_Target_Species" _
                    & "(Master_Plant_Code_FK, Park_Code, Target_Year, Species_Name)" _
                    & "VALUES "
    
            'prepare insert value
            strInsert = "('" & strMasterCode & "','" & TempVars.item("park") & "'," & TempVars.item("tgtYear") & ",'" & strSpecies & "');"
            
            'add comma if more than one row to insert
            'If (lbxTgtSpecies.ListCount - 1) > 1 And iRow < (lbxTgtSpecies.ListCount - 1) Then strInsert = strInsert & ","
            
            'finalize SQL
            strSQL = strSQL & strInsert
            
            'save full target list (insert value) [NOTE: MS Access does not support multiple insert statements, must go 1 @ a time]
            CurrentDb.Execute strSQL, dbFailOnError
            
        End If
    Next
    
    'end SQL statement
    'strSQL = strSQL & ";"
    
    'save full target list (insert values)
    'CurrentDb.Execute strSQL, dbFailOnError
    
    'set statusbar notice
    varReturn = SysCmd(acSysCmdSetStatus, "Targetlist save complete.")
    
    'pause to view status bar
    For i = 0 To 10000
        i = i
    Next i
    'reset status bar
    varReturn = SysCmd(acSysCmdSetStatus, " ")

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblSaveList_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lblSearch_Click
' Description:  Opens species search to find species for populating target list
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
Private Sub lblSearch_Click()
On Error GoTo Err_Handler
    Dim originForm As String
    
    originForm = Me.name
    
    'open species search form
    DoCmd.OpenForm "frmSpeciesSearch", acNormal, , , , acWindowNormal, originForm
    If Forms("frmSpeciesSearch").Minimized Then DoCmd.Restore

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblSearch_Click[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub


' ---------------------------------
' SUB:          Form_Close
' Description:  Actions for closing form
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
Private Sub Form_Close()
On Error GoTo Err_Handler

    'clear tempvars
    TempVars.Remove ("park")
    TempVars.Remove ("state")

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[Form_frmTgtSpecies])"
    End Select
    Resume Exit_Sub
End Sub
