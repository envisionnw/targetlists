Version =20
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
    Width =10080
    DatasheetFontHeight =11
    ItemSuffix =13
    Right =14508
    Bottom =9408
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
        Begin Section
            Height =5760
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
                    Left =8220
                    Top =4920
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
                    Name ="lblPrintSheets"
                    Caption ="Print Sheet(s)"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Print sheets"
                    GridlineColor =10921638
                    LayoutCachedLeft =8220
                    LayoutCachedTop =4920
                    LayoutCachedWidth =9838
                    LayoutCachedHeight =5293
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
                    ColumnCount =5
                    Left =360
                    Top =660
                    Width =3960
                    Height =4032
                    FontSize =10
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxDataSheets"
                    RowSourceType ="Value List"
                    RowSource ="Park;File_Code;Datasheet;File_Description;Sort_Order;CANY;Photo;CANY_Big_Rivers_"
                        "Photo_Data_Sheet_bc20150128a.pdf;photo data sheet;1;CANY;VegPlot;CANY_Big_Rivers"
                        "_Veg_Plots_Data_Sheet_bc20150128a.pdf;veg plots data sheet;2;CANY;VegCont;CANY_B"
                        "ig_Rivers_Veg_Plots_Continuation_Data_Sheet_bc20150128a.pdf;veg continuation dat"
                        "a sheet;3;CANY;VegWalk;CANY_Big_Rivers_Veg_Walk_Data_Sheet_bc20150128a.pdf;veg w"
                        "alk data sheet;4"
                    ColumnWidths ="720;3240;14;14;14"
                    OnDblClick ="[Event Procedure]"
                    OnKeyUp ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =360
                    LayoutCachedTop =660
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =4692
                    DatasheetCaption ="Datasheets"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =300
                            Width =1110
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblDatasheets"
                            Caption ="Datasheets"
                            ControlTipText ="Documents available for printing"
                            GridlineColor =10921638
                            LayoutCachedLeft =180
                            LayoutCachedTop =300
                            LayoutCachedWidth =1290
                            LayoutCachedHeight =615
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    MultiSelect =2
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =5760
                    Top =660
                    Width =3960
                    Height =4032
                    FontSize =10
                    TabIndex =1
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxPrintSheets"
                    RowSourceType ="Value List"
                    RowSource ="Park;File_Code;Datasheet;File_Description;Sort_Order"
                    ColumnWidths ="720;3240;14;14;14"
                    OnDblClick ="[Event Procedure]"
                    OnKeyUp ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Documents selected for printing"
                    GridlineColor =10921638

                    LayoutCachedLeft =5760
                    LayoutCachedTop =660
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =4692
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5580
                            Top =300
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblSheetsToPrint"
                            Caption ="Sheets to Print"
                            GridlineColor =10921638
                            LayoutCachedLeft =5580
                            LayoutCachedTop =300
                            LayoutCachedWidth =7020
                            LayoutCachedHeight =615
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4757
                    Top =3360
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
                    LayoutCachedTop =3360
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =3823
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4757
                    Top =780
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
                    LayoutCachedTop =780
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =1243
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4757
                    Top =4140
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
                    LayoutCachedTop =4140
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =4603
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4757
                    Top =1620
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
                    LayoutCachedTop =1620
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =2083
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4440
                    Top =4800
                    Width =1174
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
                    LayoutCachedLeft =4440
                    LayoutCachedTop =4800
                    LayoutCachedWidth =5614
                    LayoutCachedHeight =5158
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
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
' MODULE:       Form_frmPrintDocs
' Description:  Document printing functions & procedures
'
' Source/date:  Bonnie Campbell, 1/30/2015
' Revisions:    BLC - 1/30/2015 - initial version
' =================================

' ---------------------------------
' SUB:          Form_Load
' Description:  form loading routine
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, January 30, 2015 - for NCPN tools
' Revisions:
'   BLC - 1/30/2015 - initial version
'   BLC - 2/19/2015 - update fillList parameters & documentation
' ---------------------------------
Private Sub Form_Load()

On Error GoTo Err_Handler
       
    'file directory path
    TempVars.Add "FileDir", "U:\NCPN_WORK\___PARTS\DataSheets\"
    
    'clear headers
    lbxDataSheets.RowSource = ""
    lbxPrintSheets.RowSource = ""
    
    'initial listbox fill
    fillList Me, lbxDataSheets, lbxPrintSheets  '.RowSourceType = "Value List"

    'enable > or < *only* if at least one item selected
    'check background / text color, if gray then exit_sub
'Forms!MasterForm!Label1550.SpecialEffect = vbraised
'Forms!MasterForm!Label1550.SpecialEffect = vbetched
'Forms!MasterForm!Label1550.SpecialEffect = vbsunken
    
    DisableControl lblAdd
    DisableControl lblRemove
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[Form_frmPrintDocs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          fillList
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
Public Sub XfillList(frm As Form, lbx As ListBox)

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strQuery As String, strSQL As String
    Dim lbxDest As ListBox

    Select Case lbx.name
        Case "lbxDataSheets"
            strQuery = "qryActiveDatasheets"
            strSQL = CurrentDb.QueryDefs(strQuery).sql
            Set lbxDest = frm.Controls("lbxPrintSheets")
    End Select

    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)

    'set TempVars
    TempVars.Add "strSQL", strSQL

    'PopulateList frm.Controls(lbx), rs
    PopulateList lbx, rs, lbxDest

    'Enable move items lbls (or not)
    If lbx.ListCount > 0 Then
        lblAddAll.Visible = True
        lblRemoveAll.Visible = True
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fillList[form_frmPrintDocs])"
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
            "Error encountered (#" & Err.Number & " - lblReset_Click[Form_frmPrintDocs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxDataSheets_Click
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
Private Sub lbxDataSheets_Click()
On Error GoTo Err_Handler

    Dim varItem As Variant

    'deselect items in target control (lbxPrintSheets)
    For Each varItem In lbxPrintSheets.ItemsSelected
        lbxPrintSheets.Selected(varItem) = False
    Next

    'check for selected items --> if present, enable lblAdd
    If lbxDataSheets.ItemsSelected.count > 0 Then
        If lblAdd.backColor <> TempVars.item("ctrlAddEnabled") Then
            EnableControl lblAdd, TempVars.item("ctrlAddEnabled"), TempVars.item("textEnabled")
        End If
    Else
        DisableControl lblAdd
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxDataSheets_Click[Form_frmPrintDocs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxDataSheets_DblClick
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, January 30, 2015 - for NCPN tools
' Revisions:
'   BLC - 1/30/2015 - initial version
' ---------------------------------
Public Sub lbxDataSheets_DblClick(Cancel As Integer)
On Error GoTo Err_Handler:

    MoveSingleItem Me, "lbxDataSheets", "lbxPrintSheets"

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxDataSheets_DblClick[Form_frmPrintDocs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxDataSheets_KeyUp
' Description:  Deselects items after control update
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
Private Sub lbxDataSheets_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

    If lbxDataSheets.ItemsSelected.count > 0 And lblAdd.backColor <> TempVars.item("ctrlAddEnabled") Then
        EnableControl lblAdd, TempVars.item("ctrlAddEnabled"), TempVars.item("textEnabled")
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxDataSheets_KeyUp[Form_frmPrintDocs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxPrintSheets_Click
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
Private Sub lbxPrintSheets_Click()
On Error GoTo Err_Handler
    
    Dim varItem As Variant
    
    'deselect items in source control (lbxDataSheets)
    For Each varItem In lbxDataSheets.ItemsSelected
        lbxDataSheets.Selected(varItem) = False
    Next

    'check for selected items --> if present, enable lblRemove
    If lbxPrintSheets.ItemsSelected.count > 0 Then
        If lblRemove.backColor <> TempVars.item("ctrlRemoveEnabled") Then
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
            "Error encountered (#" & Err.Number & " - lbxPrintSheets_Click[Form_frmPrintDocs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxPrintSheets_DblClick
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, January 30, 2015 - for NCPN tools
' Revisions:
'   BLC - 1/30/2015 - initial version
' ---------------------------------
Private Sub lbxPrintSheets_DblClick(Cancel As Integer)
    
On Error GoTo Err_Handler

    MoveSingleItem Me, "lbxPrintSheets", "lbxDataSheets"

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxPrintSheets_DblClick[Form_frmPrintDocs])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          lbxPrintSheets_KeyUp
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
Private Sub lbxPrintSheets_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Handler

    If lbxDataSheets.ItemsSelected.count > 0 And lblRemove.backColor <> TempVars.item("ctrlRemoveEnabled") Then
        EnableControl lblRemove, TempVars.item("ctrlRemoveEnabled"), TempVars.item("textEnabled")
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxPrintSheets_KeyUp[Form_frmPrintDocs])"
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
' Adapted:      Bonnie Campbell, January 30, 2015 - for NCPN tools
' Revisions:
'   BLC - 1/30/2015 - initial version
' ---------------------------------
Private Sub lblAdd_Click()
On Error GoTo Err_Handler
    
    'ignore if 'disabled'
    If lblAdd.backColor = lngGray Then GoTo Exit_Sub
    
    MoveSingleItem Me, "lbxDataSheets", "lbxPrintSheets"

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblAdd_Click[Form_frmPrintDocs])"
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
' Adapted:      Bonnie Campbell, January 30, 2015 - for NCPN tools
' Revisions:
'   BLC - 1/30/2015 - initial version
' ---------------------------------
Private Sub lblRemove_Click()
On Error GoTo Err_Handler
    
    'ignore if 'disabled'
    If lblRemove.backColor = TempVars.item("ctrlDisabled") Then GoTo Exit_Sub
    
    MoveSingleItem Me, "lbxPrintSheets", "lbxDataSheets"
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblRemove_Click[Form_frmPrintDocs])"
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
' Adapted:      Bonnie Campbell, January 30, 2015 - for NCPN tools
' Revisions:
'   BLC - 1/30/2015 - initial version
' ---------------------------------
Private Sub lblAddAll_Click()
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'fetch recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset(TempVars.item("strSQL"))
    
    MoveAllItems Me, "lbxDataSheets", "lbxPrintSheets"

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblAddAll_Click[Form_frmPrintDocs])"
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
' Adapted:      Bonnie Campbell, January 30, 2015 - for NCPN tools
' Revisions:
'   BLC - 1/30/2015 - initial version
' ---------------------------------
Private Sub lblRemoveAll_Click()
On Error GoTo Err_Handler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    'fetch recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset(TempVars.item("strSQL"))
    
    MoveAllItems Me, "lbxPrintSheets", "lbxDataSheets"

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblRemoveAll_Click[Form_frmPrintDocs])"
    End Select
    Resume Exit_Sub
End Sub


' ---------------------------------
' SUB:          lblPrintSheets_Click
' Description:  Print all selected items
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
Private Sub lblPrintSheets_Click()
On Error GoTo Err_Handler

    Dim iFileCol As Integer, iRow As Integer
    Dim strPath As String, strFilePathName As String
    
    'fetch the file path
    strPath = TempVars.item("FileDir")
    
    'column for file name & path
    iFileCol = 3 - 1
    
    For iRow = 0 To lbxPrintSheets.ListCount - 1
       
       ' ---------------------------------------------------
       '  NOTE: listbox column MUST have a non-zero width to retrieve its value
       ' ---------------------------------------------------
        If lbxPrintSheets.Selected(iRow) = True Then
            
            'get the full file path
            strFilePathName = strPath & lbxPrintSheets.Column(iFileCol, iRow)
            
            'set statusbar notice
            Dim varReturn As Variant
            varReturn = SysCmd(acSysCmdSetStatus, "Printing " & strFilePathName & "...")
            
            'print it
            MsgBox "Results for printing file: " & vbCrLf & strFilePathName & vbCrLf & vbCrLf & PrintFile(strFilePathName), vbOKOnly, "Print Results"
        End If
    
    Next

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblPrintSheets_Click[Form_frmPrintDocs])"
    End Select
    Resume Exit_Sub
End Sub
