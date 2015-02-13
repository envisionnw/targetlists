Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Git
' Description:  Git related functions & procedures for version control
'
' Source/date:  Bonnie Campbell, 2/12/2015
' Revisions:    BLC - 2/12/2015 - initial version
' =================================

' ===================================================================================
'  NOTE:
'  To regenerate components backed up w/ functions using SaveAsText
'  Use the following:
'       Application.LoadFromText acForm, "YourFormName", "C:\Temp\Form_frmTest.txt"
' ===================================================================================

' ---------------------------------
' FUNCTION:     ExportVBComponent
' Description:  Export VB components (forms, modules, classes) as text files for later use
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   Requires Microsoft Visual Basic for Applications Extensibility 5.3 library (add via Tools > References)
' Source/date:
' Chip Pearson
' http://www.cpearson.com/excel/vbe.aspx
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Public Function ExportVBComponent(VBComp As VBIDE.VBComponent, _
                FolderName As String, _
                Optional fileName As String, _
                Optional OverwriteExisting As Boolean = True) As Boolean
On Error GoTo Err_Handler

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Exports module code of a VBComponent to a text file.
    ' If FileName is missing, the code will be exported to
    ' a file with the same name as the VBComponent followed by the
    ' appropriate extension.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Extension As String
    Dim FName As String
    Extension = GetFileExtension(VBComp:=VBComp)
    If Trim(fileName) = vbNullString Then
        FName = VBComp.name & Extension
    Else
        FName = fileName
        If InStr(1, FName, ".", vbBinaryCompare) = 0 Then
            FName = FName & Extension
        End If
    End If
    
    If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
        FName = FolderName & FName
    Else
        FName = FolderName & "\" & FName
    End If
    
    If Dir(FName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
        If OverwriteExisting = True Then
            Kill FName
        Else
            ExportVBComponent = False
            Exit Function
        End If
    End If
    
    VBComp.Export fileName:=FName
    ExportVBComponent = True

Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ExportVBComponent[mod_Git])"
    End Select
    Resume Exit_Function
End Function
    
' ---------------------------------
' FUNCTION:     GetFileExtension
' Description:  Return appropriate file extension for VB Components(forms, modules, classes) as text files for later use
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Chip Pearson
' http://www.cpearson.com/excel/vbe.aspx
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Public Function GetFileExtension(VBComp As VBIDE.VBComponent) As String
On Error GoTo Err_Handler
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This returns the appropriate file extension based on the Type of
    ' the VBComponent.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case VBComp.Type
            Case vbext_ct_ClassModule
                GetFileExtension = ".cls"
            Case vbext_ct_Document
                GetFileExtension = ".cls"
            Case vbext_ct_MSForm
                GetFileExtension = ".frm"
            Case vbext_ct_StdModule
                GetFileExtension = ".bas"
            Case Else
                GetFileExtension = ".bas"
        End Select

Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetFileExtension[mod_Git])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' SUB:          DocDatabase
' Description:  Documents the database to a series of text files
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Arvin Meyer, June 2, 1999
' http://www.datastrat.com/Code/DocDatabase.txt
' Renaud Bompius, June 20, 2011
' http://stackoverflow.com/questions/6408951/text-search-in-properties-access-objects/6410015#6410015
' Usage:
' Call DocDatabase from the Access IDE Immediate window.
' Creates a set of directories under and 'Exploded View' folder that will contain all the files.
' Comment: Uses the undocumented [Application.SaveAsText] syntax
'          To reload use the syntax [Application.LoadFromText]
'          Modified to set a reference to DAO 8/22/2005
'          Modified by Renaud Bompuis to export Queries as proper SQL
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Public Sub DocDatabase(Optional Path As String = "")
    
    If IsBlank(Path) Then
        Path = Application.CurrentProject.Path & "\" & Application.CurrentProject.name & " - exploded view\"
    End If

On Error Resume Next
    MkDir Path
    MkDir Path & "\Forms\"
    MkDir Path & "\Queries\"
    MkDir Path & "\Queries(SQL)\"
    MkDir Path & "\Reports\"
    MkDir Path & "\Modules\"
    MkDir Path & "\Scripts\"

On Error GoTo Err_Handler
    Dim dbs As DAO.Database
    Dim cnt As DAO.Container
    Dim doc As DAO.Document
    Dim i As Integer

    Set dbs = CurrentDb() ' use CurrentDb() to refresh Collections

    Set cnt = dbs.Containers("Forms")
    For Each doc In cnt.Documents
        Application.SaveAsText acForm, doc.name, Path & "\Forms\" & doc.name & ".txt"
    Next doc

    Set cnt = dbs.Containers("Reports")
    For Each doc In cnt.Documents
        Application.SaveAsText acReport, doc.name, Path & "\Reports\" & doc.name & ".txt"
    Next doc

    Set cnt = dbs.Containers("Scripts")
    For Each doc In cnt.Documents
        Application.SaveAsText acMacro, doc.name, Path & "\Scripts\" & doc.name & ".txt"
    Next doc

    Set cnt = dbs.Containers("Modules")
    For Each doc In cnt.Documents
        Application.SaveAsText acModule, doc.name, Path & "\Modules\" & doc.name & ".txt"
    Next doc

    Dim intfile As Long
    Dim fileName As String
    For i = 0 To dbs.QueryDefs.Count - 1
         Application.SaveAsText acQuery, dbs.QueryDefs(i).name, Path & "\Queries\" & dbs.QueryDefs(i).name & ".txt"
         fileName = Path & "\Queries(SQL)\" & dbs.QueryDefs(i).name & ".txt"
         intfile = FreeFile()
         Open fileName For Output As #intfile
         Print #intfile, dbs.QueryDefs(i).sql
         Close #intfile
    Next i

    Set doc = Nothing
    Set cnt = Nothing
    Set dbs = Nothing

Exit_Sub:
    Debug.Print "Done."
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DocDatabase[mod_Git])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          RecreateDatabase
' Description:  Recreates the database from series of text files created through SaveAsText
' Assumptions:  -
' Parameters:   XX - XX
' Returns:      XX - XX
' Throws:       none
' References:   none
' Source/date:
' Curtis Inderwiesche, May 13, 2009
' http://stackoverflow.com/questions/859530/alternative-to-application-loadfromtext-for-ms-access-queries
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Public Sub RecreateDatabase()
On Error GoTo Err_Handler
    
    For Each myFile In folder.Files
        objecttype = FSO.GetExtensionName(myFile.name)
        objectname = FSO.GetBaseName(myFile.name)
        WScript.Echo "  " & objectname & " (" & objecttype & ")"
    
        If (objecttype = "form") Then
            oApplication.LoadFromText acForm, objectname, myFile.Path
        ElseIf (objecttype = "bas") Then
            oApplication.LoadFromText acModule, objectname, myFile.Path
        ElseIf (objecttype = "mac") Then
            oApplication.LoadFromText acMacro, objectname, myFile.Path
        ElseIf (objecttype = "report") Then
            oApplication.LoadFromText acReport, objectname, myFile.Path
        ElseIf (objecttype = "sql") Then
            oApplication.LoadFromText acQuery, objectname, myFile.Path
        End If
        
    Next
    
Exit_Sub:
    Debug.Print "Done."
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RecreateDatabase[mod_Git])"
    End Select
    Resume Exit_Sub
End Sub