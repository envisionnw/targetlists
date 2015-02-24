Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Lists
' Description:  listbox functions & procedures
'
' Source/date:  Bonnie Campbell, 1/30/2015
' Revisions:    BLC - 1/30/2015 - initial version
' =================================

' ---------------------------------
'  Properties
' ---------------------------------

' ---------------------------------
' SUB:          MoveSingleItem
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
Public Sub MoveSingleItem(frm As Form, strSourceControl As String, strTargetControl As String)
    
On Error GoTo Err_Handler
    
    Dim strItem As String
    Dim intColumnCount As Integer
    
    'check for at *least* one selected item
    If frm.Controls(strSourceControl).ItemsSelected.Count = 0 Then
        MsgBox "Please select at least one item.", vbExclamation, "Oops!"
        GoTo Exit_Sub
    End If
    
    If frm.Controls(strSourceControl).ItemsSelected.Count > 1 Then
        MoveSelectedItems frm, strSourceControl, strTargetControl
        GoTo Exit_Sub
    End If
    
    For intColumnCount = 0 To frm.Controls(strSourceControl).ColumnCount - 1
        strItem = strItem & frm.Controls(strSourceControl).Column(intColumnCount) & ";"
    Next
    
    'remove extra semi-colon (;)
    strItem = Left(strItem, Len(strItem) - 1)

    'Check the length to make sure something is selected
    ' -------------------------------------------------------------------------
    '  NOTE: ListIndex is zero based, so add 1 to remove proper item
    ' -------------------------------------------------------------------------
    If Len(strItem) > 0 Then
        frm.Controls(strTargetControl).AddItem strItem
        frm.Controls(strSourceControl).RemoveItem frm.Controls(strSourceControl).ListIndex + 1
    Else
        MsgBox "Please select an item to move."
    End If


Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MoveSingleItem[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          MoveAllItems
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
Public Sub MoveAllItems(frm As Form, strSourceControl As String, strTargetControl As String)
    
On Error GoTo Err_Handler
    
    Dim strItem As String
    Dim intColumnCount As Integer, startRow As Integer
    Dim lngRowCount As Long
    
    'check for at *least* one item
    If frm.Controls(strSourceControl).ListCount = 0 Then
        MsgBox "Your list needs at least one item to move.", vbExclamation, "Oops!"
        GoTo Exit_Sub
    End If
    
    startRow = 0 'default
    'set start row
    If frm.Controls(strSourceControl).ColumnHeads = True Then
        startRow = 1
    End If
    
    For lngRowCount = startRow To frm.Controls(strSourceControl).ListCount - 1
        For intColumnCount = 0 To frm.Controls(strSourceControl).ColumnCount - 1
            strItem = strItem & frm.Controls(strSourceControl).Column(intColumnCount, lngRowCount) & ";"
        Next
        strItem = Left(strItem, Len(strItem) - 1)
        frm.Controls(strTargetControl).AddItem strItem
        strItem = ""
    Next
        
    'clear the list
    frm.Controls(strSourceControl).RowSource = ""
    
    'add back the headers
    ' -------------------------------------------------------------------------
    ' NOTE: target lbx will already have headers, so only add back to source
    ' -------------------------------------------------------------------------
    If frm.Controls(strSourceControl).ColumnHeads = True Then
        frm.Controls(strSourceControl).AddItem TempVars("lbxHdr")
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MoveAllItems[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          MoveSelectedItems
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' ManningFan, January 30,2015
' http://bytes.com/topic/access/answers/765291-populating-1-listbox-another-listbox
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
' ---------------------------------
Public Sub MoveSelectedItems(frm As Form, strSourceControl As String, strTargetControl As String)
    
On Error GoTo Err_Handler
    
    Dim iRow As Integer, startRow As Integer, i As Integer, x As Integer, iRemovedItems As Integer
    Dim arySelectedItems() As Integer
    Dim blnDimensioned As Boolean
    Dim strItem As String
    
    'check for at *least* one selected item
    If frm.Controls(strSourceControl).ItemsSelected.Count = 0 Then
        MsgBox "Please select at least one item.", vbExclamation, "Oops!"
        GoTo Exit_Sub
    End If
    
    startRow = 0 'default
    'set start row
    If frm.Controls(strSourceControl).ColumnHeads = True Then
        startRow = 1
    End If
    
    'add back the header if it doesn't exist
    If frm.Controls(strTargetControl).ColumnHeads = True And frm.Controls(strTargetControl).ListCount = 0 Then
       strItem = TempVars.item("lbxHdr") & strItem
       frm.Controls(strTargetControl).AddItem strItem
    End If
    
    'generate array of selected items
    For iRow = startRow To frm.Controls(strSourceControl).ListCount - 1
    
        'fetch array of selected items
        '--------------------------------------------------
        ' if > 1 item selected, other selected items
        ' deselected when first source item removed
        '--------------------------------------------------
        If frm.Controls(strSourceControl).Selected(iRow) Then
            
            'Array dimensioned?
            If blnDimensioned = True Then
                      
                'Yes ==> extend array 1 element largee than current upper bound
                '        w/o "Preserve" keyword previous elements erased w/ resizing
                ReDim Preserve arySelectedItems(0 To UBound(arySelectedItems) + 1) As Integer
                      
            Else
                      
                'No ==> dimension it and flag as dimensioned
                ReDim arySelectedItems(0 To 0) As Integer
                blnDimensioned = True
                          
            End If
                  
            'Add to last element in the array.
            arySelectedItems(UBound(arySelectedItems)) = iRow
        End If
    
    Next
    
    'set default
    iRemovedItems = 0
    
    'iterate through selected items
    For x = LBound(arySelectedItems) To UBound(arySelectedItems)
                        
        iRow = arySelectedItems(x) - iRemovedItems
            
        'clear string
        strItem = ""
        
        'add all columns
        For i = 0 To frm.Controls(strSourceControl).ColumnCount
            strItem = strItem & frm.Controls(strSourceControl).Column(i, iRow) & ";"
        Next i
        
        'add to target
        frm.Controls(strTargetControl).AddItem strItem
        
        'remove from source
        frm.Controls(strSourceControl).RemoveItem iRow
            
        'adjust list after removal
        If UBound(arySelectedItems) > 0 Then
            iRemovedItems = iRemovedItems + 1
        End If
    
    Next x

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MoveSelectedItems[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          PopulateListHeaders
' Description:  Populate the headers for listbox controls
' Assumptions:  headers are the same as recordset field names
'               sfrms acting as listboxes have static headers already present
' Parameters:   ctrl - listbox control
'               rs   - recordset containing list headers
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015  - initial version
'   BLC - 2/19/2015 - converted to generic to handle listbox-like controls & documentation update
' ---------------------------------
Public Sub PopulateListHeaders(ctrl As Control, rs As Recordset)

On Error GoTo Err_Handler

    Dim rows As Integer, cols As Integer, i As Integer, j As Integer, matches As Integer
    Dim frm As Form
    Dim strItem As String, strColHeads As String, aryColWidths() As String

    'exit if subform control (hdrs are static & present on sfrm)
    If ctrl.ControlType = 112 Then
        GoTo Exit_Sub
    End If

    Set frm = ctrl.Parent
    
    rows = rs.RecordCount
    cols = rs.Fields.Count
    
    If Nz(rows, 0) = 0 Then
        MsgBox "Sorry, no records found..."
        GoTo Exit_Sub
    End If
    
    'fetch column widths
    aryColWidths = Split(ctrl.ColumnWidths, ";")
    
    'populate column names (if desired)
    If ctrl.ColumnHeads = True Then
        strColHeads = ""
        For i = 0 To cols - 1
            If CInt(aryColWidths(i)) > 0 Then
                strColHeads = strColHeads & rs.Fields(i).name & ";"
            End If
        Next i
        ctrl.AddItem strColHeads
    End If

    'save headers
    TempVars.Add "lbxHdr", strColHeads

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateListHeaders[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          PopulateList
' Description:  Populate listbox and similar controls from recordset
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' krish KM, Aug. 27, 2014
' http://stackoverflow.com/questions/25526904/populate-listbox-using-ado-recordset
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
' ---------------------------------
Public Sub PopulateList(ctrlSource As Control, rs As Recordset, ctrlDest As Control)

On Error GoTo Err_Handler

    Dim frm As Form
    Dim rows As Integer, cols As Integer, i As Integer, j As Integer, matches As Integer, iZeroes As Integer
    Dim strItem As String, strColHeads As String, aryColWidths() As String

    Set frm = ctrlSource.Parent
    
    rows = rs.RecordCount
    cols = rs.Fields.Count
    
    'address no records
    If Nz(rows, 0) = 0 Then
        MsgBox "Sorry, no records found..."
        GoTo Exit_Sub
    End If
    
    'handle sfrm controls (acSubform = 112)
    If ctrlSource.ControlType = acSubform Then
        'ctrlSource.Form.Section("detail").Properties("Record Source") = rs
        'ctrlSource.Parent.Forms(ctrlSource).Form.Recordset = rs 'sfrmSpeciesListbox
        'ctrlSource.Parent.Controls(ctrlSource).Form.Recordset = rs 'sfrmSpeciesListbox
        Set ctrlSource.Form.Recordset = rs
        'Set ctrlSource.Form.Controls("tbxCode") = rs.Fields("Code")
        'Set ctrlSource.Form.Controls("tbxSpecies") = rs.Fields("Species")
        'Set ctrlSource.Form.Controls("tbxCode").RowSource = rs.Fields("Code")
        'Set ctrlSource.Form.Controls("tbxSpecies").RowSource = rs.Fields("Species")
        'Set ctrlSource.Form.Controls("tbxCode").ControlSource = rs.Fields("Code")
        'Set ctrlSource.Form.Controls("tbxSpecies").ControlSource = rs.Fields("Species")
        ctrlSource.Form.Controls("tbxCode").ControlSource = "Code"
        ctrlSource.Form.Controls("tbxSpecies").ControlSource = "Species"
        ctrlSource.Form.Controls("tbxMasterCode").ControlSource = "Master_PLANT_Code"
        
        'set the initial record count (MoveLast to get full count, MoveFirst to set display to first)
        rs.MoveLast
        ctrlSource.Parent.Form.Controls("lblSfrmSpeciesCount").Caption = rs.RecordCount & " species"
        rs.MoveFirst
        
        GoTo Exit_Sub
    End If
    
    'fetch column widths array
    aryColWidths = Split(ctrlSource.ColumnWidths, ";")
    
    'count number of 0 width elements
    iZeroes = CountArrayValues(aryColWidths, "0")
    
'    If Nz(rows, 0) = 0 Then
'        MsgBox "Sorry, no records found..."
'        GoTo Exit_Sub
'    End If
    
    'clear out existing values
    ClearList ctrlSource
    
    'populate column names (if desired)
    If ctrlSource.ColumnHeads = True Then
        PopulateListHeaders ctrlSource, rs
        
        'populate second listbox headers if present
        If ctrlDest.ColumnHeads = True Then
            ClearList ctrlDest
            PopulateListHeaders ctrlDest, rs
        End If
    End If
    
    'populate data
    Select Case ctrlSource.RowSourceType
        Case "Table/Query"
            Set ctrlSource.Recordset = rs
        Case "Value List"
            
            'initialize
            i = 0
            
            Do Until rs.EOF
            
                'initialize item
                strItem = ""
                    
                'generate item
                For j = 0 To cols - 1
                    'check if column is displayed width > 0
                    If CInt(aryColWidths(j)) > 0 Then
                    
                        strItem = strItem & rs.Fields(j).Value & ";"
                    
                        'determine how many separators there are (";") --> should equal # cols
                        matches = (Len(strItem) - Len(Replace$(strItem, ";", ""))) / Len(";")
                        
                        'add item if not already in list --> # of ; should equal cols - 1
                        'but # in list should only be # of non-zero columns --> cols - iZeroes
                        If matches = cols - iZeroes Then
                            ctrlSource.AddItem strItem
                            'reset the string
                            strItem = ""
                        End If
                    
                    End If
                
                Next
                
                i = i + 1
                
                rs.MoveNext
            Loop
        Case "Field List"
    End Select

     'MsgBox ctrlSource.ListCount & " in list" & vbCrLf & rs.RecordCount & " in rs", vbOKOnly, "Num in list"
    'refresh control
    'lbx.Requery

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateList[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          XPopulateList
' Description:  XX
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' krish KM, Aug. 27, 2014
' http://stackoverflow.com/questions/25526904/populate-listbox-using-ado-recordset
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
' ---------------------------------
Public Sub XPopulateList(lbx As ListBox, rs As Recordset, lbxDest As ListBox)

On Error GoTo Err_Handler

    Dim frm As Form
    Dim rows As Integer, cols As Integer, i As Integer, j As Integer, matches As Integer, iZeroes As Integer
    Dim strItem As String, strColHeads As String, aryColWidths() As String

    Set frm = lbx.Parent
    
    rows = rs.RecordCount
    cols = rs.Fields.Count
    
    'fetch column widths array
    aryColWidths = Split(lbx.ColumnWidths, ";")
    
    'count number of 0 width elements
    iZeroes = CountArrayValues(aryColWidths, "0")
    
    If Nz(rows, 0) = 0 Then
        MsgBox "Sorry, no records found..."
        GoTo Exit_Sub
    End If
    
    'clear out existing values
    ClearList lbx
    
    'populate column names (if desired)
    If lbx.ColumnHeads = True Then
        PopulateListHeaders lbx, rs
        
        'populate second listbox headers if present
        If lbxDest.ColumnHeads = True Then
            ClearList lbxDest
            PopulateListHeaders lbxDest, rs
        End If
    End If
    
    'populate data
    Select Case lbx.RowSourceType
        Case "Table/Query"
            Set lbx.Recordset = rs
        Case "Value List"
            
            'initialize
            i = 0
            
            Do Until rs.EOF
            
                'initialize item
                strItem = ""
                    
                'generate item
                For j = 0 To cols - 1
                    'check if column is displayed width > 0
                    If CInt(aryColWidths(j)) > 0 Then
                    
                        strItem = strItem & rs.Fields(j).Value & ";"
                    
                        'determine how many separators there are (";") --> should equal # cols
                        matches = (Len(strItem) - Len(Replace$(strItem, ";", ""))) / Len(";")
                        
                        'add item if not already in list --> # of ; should equal cols - 1
                        'but # in list should only be # of non-zero columns --> cols - iZeroes
                        If matches = cols - iZeroes Then
                            lbx.AddItem strItem
                            'reset the string
                            strItem = ""
                        End If
                    
                    End If
                
                Next
                
                i = i + 1
                
                rs.MoveNext
            Loop
        Case "Field List"
    End Select

     MsgBox lbx.ListCount & " in list" & vbCrLf & rs.RecordCount & " in rs", vbOKOnly, "Num in list"
    'refresh control
    'lbx.Requery

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateList[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          SortList
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
Public Sub SortList(lbx As ListBox, rs As Recordset, orderCol As Integer)

On Error GoTo Err_Handler
    
    Dim propValues As Property
    Dim rows As Integer, cols As Integer, i As Integer, j As Integer
    Dim frm As Form
    Dim strItem As String

    Set frm = lbx.Parent
    
    'sort data
    Dim aryValues As Variant
    
    Set propValues = frm.Controls(lbx).Properties("RowSource")
    aryValues = Split(propValues, ";")
    
    'iterate
    Dim a As Integer
    For a = LBound(aryValues) To UBound(aryValues) - 1
        
    
    Next
'----
Dim prp As Property, Ary, hld As String, Pak As String
Dim x As Integer, y As Integer
   
   Set prp = lbx.Properties("RowSource") 'Me!ListBoxName.Properties("RowSource")
   Ary = Split(prp, ";")
   
   For x = LBound(Ary) To UBound(Ary) - 1
      For y = x + 1 To UBound(Ary)
         If Ary(y) < Ary(x) Then
            hld = Ary(x)
            Ary(x) = Ary(y)
            Ary(y) = hld
         End If
      Next
   Next

   For x = LBound(Ary) To UBound(Ary)
      If Pak <> "" Then
         Pak = Pak & ";" & Ary(x)
      Else
         Pak = Ary(x)
      End If
   Next

   prp = Pak
   
   Set prp = Nothing
'----

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SortList[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' FUNCTION:     IsListDuplicate
' Description:  Check if item is already on the list
' Assumptions:  -
' Parameters:   lbx - listbox control to check (listbox object)
'               col - column which would hold the item being checked (integer)
'               item - name of item to be checked (string)
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
' ---------------------------------
Public Function IsListDuplicate(lbx As ListBox, col As Integer, item As String) As Boolean
On Error GoTo Err_Handler
    
    Dim isDupe As Boolean
    Dim i As Integer
    
    'set default
    isDupe = False
    
    'iterate through listbox (use .Column(col,i) vs .ListIndex(i) which results in error 451 property let not defined, property get...)
    For i = 0 To lbx.ListCount
        'check if item exists in listbox
        If lbx.Column(col, i) = item Then
            'duplicate, so exit
            isDupe = True
            GoTo Exit_Function
        End If
    Next

Exit_Function:
    IsListDuplicate = isDupe
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsListDuplicate[mod_Lists])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' SUB:          ClearList
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
Public Sub ClearList(lbx As ListBox)

On Error GoTo Err_Handler

    'clear listbox items
    lbx.RowSource = ""

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ClearList[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          DisableControl
' Description:  Set color scheme for labels so they appear disabled
' Assumptions:  Assumes control has BackColor and ForeColor properties
' Parameters:   ctrl - control to set color scheme for
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015 - initial version
' ---------------------------------
Public Sub DisableControl(ctrl As Control)

On Error GoTo Err_Handler
    
    ctrl.backcolor = lngLtGray
    ctrl.forecolor = lngGray
    
    If ctrl.ControlType = acCommandButton Then
        ctrl.borderColor = lngGray
    End If

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisableControl[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          EnableControl
' Description:  Set color scheme for labels so they appear enabled
' Assumptions:  Assumes control has BackColor and ForeColor properties
' Parameters:   ctrl - control to set color scheme for
'               backColor - long value for desired back color
'               foreColor - long value for desired fore (text) color
'               optionally for command buttons:
'               borderColor - long value for desired border color
'               hoverColor - long value for desired hover color
'               pressedColor - long value for desired pressed button color
'               hoverForeColor - long value for desired hover fore (text) color
'               pressedForeColor - long value for desired pressed button fore (text) color
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015 - initial version
' ---------------------------------
Public Function EnableControl(ctrl As Control, backcolor As Long, forecolor As Long, _
                                Optional borderColor As Long, _
                                Optional hoverColor As Long, _
                                Optional pressColor As Long, _
                                Optional hoverForeColor As Long, _
                                Optional pressedForeColor As Long)
On Error GoTo Err_Handler
    
    ctrl.backcolor = backcolor
    ctrl.forecolor = forecolor
    
    If ctrl.ControlType = acCommandButton Then
        ctrl.borderColor = borderColor
        ctrl.hoverColor = hoverColor
        ctrl.pressedColor = pressColor
        ctrl.hoverForeColor = hoverForeColor
        ctrl.pressedForeColor = pressedForeColor
    End If

Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisableControl[mod_Lists])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' SUB:          SaveListToTable
' Description:  Save list items to table
' Assumptions:  Assumes control has BackColor and ForeColor properties
' Parameters:   ctrl - control to iterate through
'               tbl - table being
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 8, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/8/2015 - initial version
' ---------------------------------
Public Sub SaveListToTable(ctrl As Control, tbl As String, tblFields As Variant, blnSelectedOnly As Boolean)

On Error GoTo Err_Handler
    
    Dim strSQL As String, strFields As String
    Dim i As Integer, iRow As Integer, jCol As Integer
    
    strSQL = "INSERT INTO " & tbl & " " & tblFields & "VALUES ("
    
    ' prepare fields
    strFields = ""
    For i = 0 To UBound(tblFields)
    
        Select Case tblFields(1, i)
            Case "Integer"
            Case "VarChar"
        End Select
        strFields = strFields
    
    Next

    'iterate through items
    For iRow = 0 To ctrl.ListCount - 1
    
            For jCol = 0 To ctrl.ColumnCount - 1
            
            strSQL = strSQL & "'" & ctrl.Column(jCol, iRow) & "'"
             
            CurrentDb.Execute strSQL, dbFailOnError
            
            Next
    Next 'iRow

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisableControl[mod_Lists])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' FUNCTION:     CountArrayValues
' Description:  XX
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
Public Function CountArrayValues(Ary As Variant, val As Variant) As Integer

On Error GoTo Err_Handler
    
    Dim i As Integer, numItems As Integer

    'default
    numItems = 0
    
    If IsArray(Ary) Then
    
        For i = LBound(Ary) To UBound(Ary)
            If Ary(i) = val Then
                numItems = numItems + 1
            End If
        Next
        
    End If
    
    CountArrayValues = numItems

Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CountArrayValues[mod_Lists])"
    End Select
    Resume Exit_Function
End Function