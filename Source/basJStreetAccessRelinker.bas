'--------------------------------------------------------------------
'
' Copyright 1996-2011 J Street Technology, Inc.
' www.JStreetTech.com
'
' This code may be used and distributed as part of your application
' provided that all comments remain intact.
'
' J Street Technology offers this code "as is" and does not assume
' any liability for bugs or problems with any of the code.  In
' addition, we do not provide free technical support for this code.
'
'--------------------------------------------------------------------
Option Compare Database
Option Explicit

'Revised Type Declare for compatability with NT
Type tagOPENFILENAME
    lStructSize         As Long
    hwndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As Long
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    Flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As Long
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" (OPENFILENAME As tagOPENFILENAME) As Long
  
Private Sub HandleError(strLoc As String, strError As String, intError As Integer)
    MsgBox strLoc & ": " & strError & " (" & intError & ")", 16, "CheckTableLinks"
End Sub

Private Function TableLinkOkay(strTableName As String) As Boolean
'Function accepts a table name and tests first to determine if linked
'table, then tests link by performing refresh link.
'Error causes TableLinkOkay = False, else TableLinkOkay = True
    Dim CurDB As DAO.Database
    Dim tdf As TableDef
    Dim strFieldName As String
    On Error GoTo TableLinkOkayError
    Set CurDB = DBEngine.Workspaces(0).Databases(0)
    Set tdf = CurDB.TableDefs(strTableName)
    TableLinkOkay = True
    If tdf.Connect <> "" Then
        strFieldName = tdf.Fields(0).Name    'Do not test if nonlinked table
    End If
    TableLinkOkay = True
TableLinkOkayExit:
    Exit Function
TableLinkOkayError:
    TableLinkOkay = False
    GoTo TableLinkOkayExit
End Function

'----------------------------------------------------------------
Private Function Relink(tdf As TableDef) As Boolean
'Function accepts a tabledef and tests first to determine if linked
'table, then links table by performing refresh link.
'Error causes Relink = False, else Relink = True
    On Error GoTo RelinkError
    Relink = True
    If tdf.Connect <> "" Then
        tdf.RefreshLink     'Do not test if local or system table
    End If
    Relink = True
RelinkExit:
        Exit Function
RelinkError:
    Relink = False
    GoTo RelinkExit
End Function

'---------------------------------------------------------------------------
Private Sub RelinkTables(strCurConnectProp As String, intResultcode As Integer)
'This subroutine accepts a table connect property and displays a dialog to allow
'modification of table links.  Routine verifies link for each modification.
'intResultcode = 0 if cancel ocx or no link change, 1 if new links OK, and
'2 if link check fails.

    Dim CurDB As DAO.Database
    Dim NewDB As Database
    Dim tdf As TableDef
    Dim strFilter As String
    Dim strDefExt As String
    Dim strTitle As String
    Dim OPENFILENAME As tagOPENFILENAME
    Dim strFileName As String
    Dim strFileTitle As String
    Dim APIResults As Long
    Dim intSlashLoc As Integer
    Dim intConnectCharCt As Integer
    Dim strDBName As String
    Dim strPath As String
    Dim strNewConnectProp As String
    Dim intNumTables As Integer
    Dim intTableIndex As Integer
    Dim strTableName As String
    Dim strSaveCurConnectProp As String
    Dim strMsg As String
    Dim varReturnVal
    Dim strAccExt As String
    
    Const OFN_PATHMUSTEXIST = &H1000
    Const OFN_FILEMUSTEXIST = &H800
    Const OFN_HIDEREADONLY = &H4
    
    On Error GoTo RelinkTablesError
    
    'Returned by GetOpenFileName
    'Revised to handle to the Win32 structure
    'strFileName = Space$(256)
    'strFileTitle = Space$(256)
    strFileName = String(256, 0)
    strFileTitle = String(256, 0)
    
    Set CurDB = DBEngine.Workspaces(0).Databases(0)
    strSaveCurConnectProp = strCurConnectProp
    
    'Parse table connect property to get data base name
        intSlashLoc = 1
        intConnectCharCt = Len(strCurConnectProp)
        Do Until InStr(intSlashLoc, strCurConnectProp, "\") = 0
            intSlashLoc = InStr(intSlashLoc, strCurConnectProp, "\") + 1
        Loop
        strDBName = Right$(strCurConnectProp, intConnectCharCt - intSlashLoc + 1)
        strPath = Right$(strCurConnectProp, intConnectCharCt - 10)
        strPath = Left$(strPath, intSlashLoc - 12)
        
    'Set up display of dialog
    'October 2009 - now handles Access 2007 formats ACCDB and ACCDE
    strAccExt = "*.accdb; *.mdb; *.mda; *.accda; *.mde; *.accde"
    strFilter = "Microsoft Office Access (" & strAccExt & ")" & Chr$(0) & strAccExt & Chr$(0) & _
                "All Files (*.*)" & Chr$(0) & "*.*" & _
                Chr$(0) & Chr$(0)
    strTitle = "Find new location of " & strDBName
    strDefExt = "mdb"
    
    'Revisions to handle to the Win32 structure
    'See changes to type declare
    '-----------------------------------------------------------
    With OPENFILENAME
        .lStructSize = Len(OPENFILENAME)
        .hwndOwner = Application.hWndAccessApp
        .lpstrFilter = strFilter
        .nFilterIndex = 1
        .lpstrFile = strDBName & String(256 - Len(strDBName), 0)
        .nMaxFile = Len(strFileName) - 1
        .lpstrFileTitle = strFileTitle
        .nMaxFileTitle = Len(strFileTitle) - 1
        .lpstrTitle = strTitle
        .Flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
        .lpstrDefExt = strDefExt
        .hInstance = 0
        .lpstrCustomFilter = 0
        .nMaxCustFilter = 0
        .lpstrInitialDir = strPath
        .nFileOffset = 0
        .nFileExtension = 0
        .lCustData = 0
        .lpfnHook = 0
        .lpTemplateName = 0
    End With
    '-----------------------------------------------------------
    APIResults = GetOpenFileName(OPENFILENAME)
    intResultcode = APIResults
    If APIResults = 1 Then      '1 if user selected file
        strNewConnectProp = ";DATABASE=" & OPENFILENAME.lpstrFile
        If Trim(strNewConnectProp) <> Trim(strSaveCurConnectProp) Then
        
    'Open New Database and create New Connect Property
            DoCmd.Hourglass True
            Set NewDB = OpenDatabase(OPENFILENAME.lpstrFile, False, True)
         
    'Set tables connect property to new connect & test
            intNumTables = CurDB.TableDefs.Count
            varReturnVal = SysCmd(acSysCmdInitMeter, "Linking Access Database", intNumTables)
            For intTableIndex = 0 To intNumTables - 1
                DoEvents
                varReturnVal = SysCmd(acSysCmdUpdateMeter, intTableIndex)
                Set tdf = CurDB.TableDefs(intTableIndex)
                If tdf.Connect = strCurConnectProp Then
                    tdf.Connect = strNewConnectProp
                    strTableName = tdf.Name
                    If Not Relink(tdf) Then
                        'Link failed, restore previous connect property and generate msgs
                        tdf.Connect = strCurConnectProp
                        intResultcode = 2       'Link failed
                        strSaveCurConnectProp = Right(strSaveCurConnectProp, intConnectCharCt - 10)
                        strMsg = "Access Table: " & strTableName & " link failed using selected database." & vbCrLf & vbCrLf & "Table is still linked to previous database path: " & strSaveCurConnectProp & "."
                        strTitle = "Failed Access Table Link"
                        MsgBox strMsg, 16, strTitle
                    End If
                End If
            Next intTableIndex
        varReturnVal = SysCmd(acSysCmdRemoveMeter)
        Else
            intResultcode = 0   'No change in Link
        End If
    End If
        
RelinkTablesExit:
    Exit Sub
    
RelinkTablesError:
    HandleError "RelinkTables", Error, Err
    Resume RelinkTablesExit
    Resume
End Sub

'------------------------------------------------------------------
Public Sub jstCheckTableLinks(CheckMode As String, LinksChanged As Boolean, LinksOK As Boolean, Optional CheckAppFolder As Boolean)
'
'INPUT:
'CheckMode = "prompt", Subroutine queries operator for location of
'   each database required by linked tables.  Msgbox for each failed link
'   and summary Msgbox on final link status (success or failure) if any
'   links were changed.  If no links changed, then no summary status.
'
'CheckMode = "full", Subroutine identifies invalid table links
'   and queries operator for location of database(s) required to satisfy
'   failed links.  Msgbox for each failed link and summary Msgbox
'   if link failures.  No Msgbox appears if all links are valid.
'
'CheckMode = "quick", same as "full" except that only the first table for
'   each linked database is checked.  If the link is not valid, the user is
'   is prompted for the location of the database and all tables in that
'   database are relinked.
'
'CheckAppFolder = True, override linked table connections if the same database name
'   exists in the application folder.  If False or not specified, no override occurs.
'
'OUTPUT:
'LinksChanged = true if at least one table link was changed.
'               false if no links where changed.
'LinksOK =      true if all links are OK upon subroutine exit.
'               false if least one table link was not successful.
'--------------------------------------------------------------------

    Dim CurDB As Database
    Dim tdf As TableDef
    Dim TableConnectPropBadArray() As String, intDBBadCount As Integer
    Dim TableConnectPropChkArray() As String, intDBChkCount As Integer
    Dim UniquePathArray() As Variant, intDBCount As Integer, intDBIndex As Integer, intDBOverrideIndex As Integer
    Dim bOverride As Boolean
    Dim bPathFound As Boolean
    Dim strUniqueDBPath As String
    Dim strFileSearch As String
    Dim intTableIndex As Integer
    Dim intNumTables As Integer
    Dim strTableName As String
    Dim strFieldName As String
    Dim intBadIndex As Integer
    Dim intChkIndex As Integer
    Dim fFound As Integer
    Dim fAllFound As Integer
    Dim fLinkGood As Integer
    Dim strCurConnectProp As String
    Dim intResultcode As Integer
    Dim strMsg As String
    Dim strTitle As String
    Dim intNoLinksChanged As Integer
    Dim varReturnVal As Variant
    
    On Error GoTo CheckTableLinksError
    DoCmd.Hourglass True
    varReturnVal = SysCmd(acSysCmdSetStatus, "Checking linked databases.")
    Set CurDB = DBEngine.Workspaces(0).Databases(0)
    
    'Get number of tables.
    intNumTables = CurDB.TableDefs.Count
    ReDim TableConnectPropBadArray(intNumTables)     'Set largest size
    ReDim TableConnectPropChkArray(intNumTables)     'Set largest size
    ReDim UniquePathArray(intNumTables, 1)
        
    'If app configured to first check in applicaiton folder for linked databases
    If CheckAppFolder = True Then
        For intTableIndex = 0 To intNumTables - 1
            Set tdf = CurDB.TableDefs(intTableIndex)
            'If there is a connect string
            If tdf.Connect & "" <> "" Then
                bPathFound = False
                'Loop through the array to check for pre-existence of database to preserve uniqueness of db paths
                For intDBIndex = 0 To (intNumTables - 1)
                    If tdf.Connect = UniquePathArray(intTableIndex, 0) Then
                        bPathFound = True
                        Exit For
                    End If
                Next
                        
                'If the path was not found in the array, add it to the unique array of paths.
                If bPathFound = False Then
                    UniquePathArray(intDBCount, 1) = 0
                    UniquePathArray(intDBCount, 0) = tdf.Connect
                    intDBCount = intDBCount + 1
                End If
            End If
        Next
        
        'Loop through all databases in array; set Override 'flag'(second column of array)
        For intDBIndex = 0 To intDBCount
            strUniqueDBPath = UniquePathArray(intDBIndex, 0)
            UniquePathArray(intDBIndex, 1) = ExistsInAppFolder(strUniqueDBPath)
        Next
        
    End If
    
    'Set up Array of Databases (all if forcelink is true, failed links if
    '   forcelink is false) (local and system tables will pass test).
    varReturnVal = SysCmd(acSysCmdInitMeter, "Checking linked databases.", intNumTables)
    LinksOK = True   'Assume success
    For intTableIndex = 0 To intNumTables - 1
        DoEvents
        varReturnVal = SysCmd(acSysCmdUpdateMeter, intTableIndex)
        Set tdf = CurDB.TableDefs(intTableIndex)
        fFound = False
          
        If Left(tdf.Connect, 10) = ";DATABASE=" Then
            'BGC -- changed from NOT "ODBC" to = ";DATABASE=" explicitly to get Access tables only
          
            strCurConnectProp = tdf.Connect
                            
            If CheckAppFolder = True Then
                bOverride = False
                    For intDBOverrideIndex = 0 To intDBCount
                        If tdf.Connect & "" <> "" And tdf.Connect = UniquePathArray(intDBOverrideIndex, 0) And UniquePathArray(intDBOverrideIndex, 1) = True Then
                            bOverride = True
                            strFileSearch = UniquePathArray(intDBOverrideIndex, 0)
                            tdf.Connect = ";DATABASE=" & PathOnly(CurDB.Name) & FileOnly(strFileSearch)
                            Exit For
                        End If
                    Next
    
            End If
            
            If bOverride = True Then
                If Not Relink(tdf) Then
                    'Link failed, restore previous connect property and generate msgs
                    tdf.Connect = strCurConnectProp
                    'intResultcode = 2       'Link failed
                    strMsg = "Application Folder Table: " & tdf.Name & " link failed." & vbCrLf & vbCrLf & "The current path for this linked table is: " & strCurConnectProp & "."
                    strTitle = "Failed Table Link"
                    MsgBox strMsg, 16, strTitle
                End If
            Else ' regular table, not overridden
            
                Select Case CheckMode
                Case "prompt"
                    ' put each connect string into the Bad array to force prompting later
                    For intBadIndex = 0 To intDBBadCount
                        If tdf.Connect = TableConnectPropBadArray(intBadIndex) Then
                            fFound = True
                            Exit For
                        End If
                    Next intBadIndex
                    If Not fFound Then
                        TableConnectPropBadArray(intDBBadCount) = tdf.Connect
                        intDBBadCount = intDBBadCount + 1
                    End If
                
                Case "full"
                    ' check each link, and put each bad connect string into
                    ' the Bad array to prompt later
                    For intBadIndex = 0 To intDBBadCount
                        If tdf.Connect = TableConnectPropBadArray(intBadIndex) Then
                            fFound = True
                            Exit For
                        End If
                    Next intBadIndex
                    If Not fFound Then
                        If Not TableLinkOkay(tdf.Name) Then
                            TableConnectPropBadArray(intDBBadCount) = tdf.Connect
                            intDBBadCount = intDBBadCount + 1
                            LinksOK = False
                        End If
                    End If
                
                Case "quick"
                    ' for each link, see if it has already been checked.
                    ' if it hasn't, add it to the checked array,
                    ' and check it.  If the link is bad, add it to the bad array to prompt later.
                    For intChkIndex = 0 To intDBChkCount
                        If tdf.Connect = TableConnectPropChkArray(intChkIndex) Then
                            fFound = True
                            Exit For
                        End If
                    Next intChkIndex
                    If Not fFound Then
                        TableConnectPropChkArray(intDBChkCount) = tdf.Connect
                        intDBChkCount = intDBChkCount + 1
                        If Not TableLinkOkay(tdf.Name) Then
                            TableConnectPropBadArray(intDBBadCount) = tdf.Connect
                            intDBBadCount = intDBBadCount + 1
                            LinksOK = False
                        End If
                    End If
            
                Case Else
                    MsgBox "CheckMode parameter """ & CheckMode & """ is not valid.  It must be ""prompt"", ""full"" or ""quick"".", vbCritical + vbOKOnly
                    LinksChanged = False
                    GoTo CheckTableLinksExit
                    
                End Select
            End If ' overridden table
        End If ' an Access linked table
            
        
    Next intTableIndex
    varReturnVal = SysCmd(acSysCmdRemoveMeter)
    
    'Prompt user to locate each database in TableConnectPropBadArray.
    varReturnVal = SysCmd(acSysCmdSetStatus, "Linking databases.")
    fAllFound = True   'Assume success in relinking all tables.
    intNoLinksChanged = 0    'Avoid successful message if no links were changed.
    For intBadIndex = 0 To intDBBadCount - 1
        DoEvents
        strCurConnectProp = TableConnectPropBadArray(intBadIndex)
        RelinkTables strCurConnectProp, intResultcode
        intNoLinksChanged = intNoLinksChanged + intResultcode
        If CheckMode = "prompt" Then
            If intResultcode = 2 Then fAllFound = False   'Failed relink.
        Else
            If Not intResultcode = 1 Then fAllFound = False
        End If
    Next intBadIndex
    
    'Display summary messages based upon forcelink value
    strTitle = "Database Links"
    If fAllFound = False Then
        strMsg = "One or more Access database tables may not be correctly linked."
        MsgBox strMsg, 16, strTitle
        LinksOK = False
    Else
        If CheckMode = "prompt" And intNoLinksChanged <> 0 Then
            strMsg = "All Access databases were linked successfully."
            MsgBox strMsg, 0, strTitle
        End If
        If CheckMode <> "prompt" Then LinksOK = True
    End If
    
    'Setup links changed flag.
    If intNoLinksChanged = 0 Then
        LinksChanged = False
    Else
        LinksChanged = True
    End If

CheckTableLinksExit:
    DoCmd.Hourglass False
    varReturnVal = SysCmd(acSysCmdClearStatus)
    Exit Sub
CheckTableLinksError:
    HandleError "CheckTableLinks", Error, Err
    Resume CheckTableLinksExit
End Sub

Public Function jstCheckTableLinks_Prompt()
    'prompt for new database locations of linked tables
    jstCheckTableLinks CheckMode:="prompt", LinksChanged:=False, LinksOK:=False, CheckAppFolder:=False
End Function

Public Function jstCheckTableLinks_Full()
    'check linked tables
    jstCheckTableLinks CheckMode:="full", LinksChanged:=False, LinksOK:=False, CheckAppFolder:=False
End Function

Public Function jstCheckTableLinks_Quick()
    'check linked tables, only the first per database
    jstCheckTableLinks CheckMode:="quick", LinksChanged:=False, LinksOK:=False, CheckAppFolder:=False
End Function

Private Function ExistsInAppFolder(strPath As String) As Boolean
    On Error GoTo Err_ExistsInAppFolder

    Dim db As Database
    Dim I As Integer
    Dim lngPos As Long
    Dim strDBName As String
    Dim strAppPath As String
    Dim strCurrPath As String
    
    ExistsInAppFolder = False
    
    Set db = CurrentDb
     
    strDBName = FileOnly(strPath)
    strCurrPath = PathOnly(db.Name)
         
    If FileExists(strCurrPath & strDBName) Then
        ExistsInAppFolder = True
    End If
     
Exit_ExistsInAppFolder:
    On Error Resume Next
    db.Close
    Set db = Nothing
    Exit Function

Err_ExistsInAppFolder:
    ExistsInAppFolder = False
    Resume Exit_ExistsInAppFolder
    Resume
End Function

Private Function FileExists(Path As Variant) As Boolean
    On Error GoTo Err_FileExists
    
    Dim varRet As Variant
    
    If IsNull(Path) Then
        FileExists = False
        Exit Function
    End If
    
    varRet = Dir(Path)
    
    If Not IsNull(varRet) And varRet <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If

Exit_FileExists:
    Exit Function

Err_FileExists:
    FileExists = False
    Resume Exit_FileExists

End Function

Private Function FileOnly(WholePath As Variant) As Variant
    On Error GoTo Err_FileOnly
    
    Dim FileOnlyPos
    
    If IsNull(WholePath) Then
        FileOnly = Null
        Exit Function
    End If
    
    FileOnlyPos = InStrRight(WholePath, "\") + 1
    
    FileOnly = Mid(WholePath, FileOnlyPos)
    
Exit_FileOnly:
    Exit Function
Err_FileOnly:
    MsgBox Err.Number & ", " & Err.Description
    Resume Exit_FileOnly
End Function

Private Function PathOnly(WholePath As Variant) As Variant
    On Error GoTo Err_PathOnly
    
    Dim FileOnlyPos
    
    If IsNull(WholePath) Then
        PathOnly = Null
        Exit Function
    End If
    
    FileOnlyPos = InStrRight(WholePath, "\") + 1
    
    PathOnly = Left(WholePath, FileOnlyPos - 1)
    
Exit_PathOnly:
    Exit Function
Err_PathOnly:
    MsgBox Err.Number & ", " & Err.Description
    Resume Exit_PathOnly
End Function

Private Function InStrRight(SearchString As Variant, soughtString As Variant) As Variant
    On Error GoTo Err_InStrRight
    Dim SoughtLen As Integer
    Dim Found As Integer
    Dim Pos As Integer
    
    If IsNull(SearchString) Or IsNull(soughtString) Then
        InStrRight = Null
        Exit Function
    End If
    
    If SearchString = "" Or soughtString = "" Then
        InStrRight = 0
        Exit Function
    End If
    
    SoughtLen = Len(soughtString)
    Found = False
    Pos = Len(SearchString) - SoughtLen + 1
    
    Do While Pos > 0 And Not Found
        If Mid(SearchString, Pos, SoughtLen) = soughtString Then
            Found = True
        Else
            Pos = Pos - 1
        End If
    Loop
    
    InStrRight = Pos
Exit_InStrRight:
    Exit Function
Err_InStrRight:
    MsgBox Err.Number & ", " & Err.Description
    Resume Exit_InStrRight
End Function