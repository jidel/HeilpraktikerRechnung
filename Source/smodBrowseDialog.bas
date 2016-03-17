'======================================================================
' Modul BrowseDialog
'======================================================================

Option Compare Database
Option Explicit

Private Type BROWSEINFO
    hOwner          As Long
    pidlRoot        As Long
    pszDisplayName  As String
    lpszTitle       As String
    ulFlags         As Long
    lpfn            As Long
    lParam          As Long
    iImage          As Long
End Type

'========== Funktionen einbinden ==========

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias _
            "SHGetPathFromIDListA" (ByVal pidl As Long, _
            ByVal pszPath As String) As Long
 
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias _
            "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) _
            As Long
 
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
            (ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, _
            lParam As Any) As Long
            
'========== Private Konstanten ==========

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BFFM_SETSELECTION = &H466
Private Const BFFM_INITIALIZED = 1


Global StartDir As String
 
'======================================================================
' Public Function VerzeichnisSuchen
'======================================================================
 
Public Function VerzeichnisSuchen(szDialogTitle As String, _
                StartVerzeichnis As String) As String
 
  Dim X         As Long
  Dim bi        As BROWSEINFO
  Dim dwIList   As Long
  Dim szPath    As String
  Dim wPos      As Integer
 
  StartDir = StartVerzeichnis
 
  With bi
        .hOwner = hWndAccessApp
        .lpszTitle = szDialogTitle
        .ulFlags = BIF_RETURNONLYFSDIRS
        .lpfn = DummyFunc(AddressOf BrowseCallbackProc)
    End With
 
    dwIList = SHBrowseForFolder(bi)
    szPath = Space$(512)
    X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
 
    If X Then
        wPos = InStr(szPath, Chr(0))
        VerzeichnisSuchen = Left$(szPath, wPos - 1)
    Else
        VerzeichnisSuchen = ""
    End If
End Function

'======================================================================
' Public Function BrowseCallbackProc
'======================================================================

Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                ByVal lParam As Long, ByVal lpData As Long) As Long
 
    Dim pathstring  As String
    Dim retval      As Long
 
    Select Case uMsg
        Case BFFM_INITIALIZED
            pathstring = StartDir
            retval = SendMessage(hWnd, BFFM_SETSELECTION, _
                     ByVal CLng(1), ByVal pathstring)
    End Select
 
    BrowseCallbackProc = 0
 
End Function

'======================================================================
' Public Function DummyFunc
'======================================================================

Public Function DummyFunc(ByVal param As Long) As Long
 
    DummyFunc = param
 
End Function