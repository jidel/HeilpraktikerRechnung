'======================================================================
' Modul BrowseDialog
'======================================================================

Option Compare Database
Option Explicit

#If VBA7 Then
    Private Type BROWSEINFO
        hOwner          As LongPtr        ' Handle, so LongPtr
        pidlRoot        As LongPtr        ' Pointer, so LongPtr
        pszDisplayName  As String
        lpszTitle       As String
        ulFlags         As Long
        lpfn            As LongPtr        ' Callback pointer, LongPtr
        lParam          As LongPtr        ' LPARAM, LongPtr
        iImage          As Long
    End Type
#Else
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
#End If

'========== Funktionen einbinden ==========

#If VBA7 Then
    Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias _
        "SHGetPathFromIDListA" (ByVal pidl As LongPtr, ByVal pszPath As String) As Long

    Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias _
        "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As LongPtr

    Private Declare PtrSafe Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
        ByVal hWnd As LongPtr, ByVal Msg As Long, wParam As Any, lParam As Any) As Long
#Else
    Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias _
        "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

    Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias _
        "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

    Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
        ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long
#End If

            
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
#If VBA7 Then
  Dim dwIList   As LongPtr    ' PIDL handle should be LongPtr in 64-bit
#Else
  Dim dwIList   As Long
#End If
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

#If VBA7 Then
Public Function BrowseCallbackProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, _
                ByVal lParam As LongPtr, ByVal lpData As LongPtr) As Long
#Else
Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                ByVal lParam As Long, ByVal lpData As Long) As Long
#End If
 
    Dim pathstring  As String
    Dim RetVal      As Long
 
    Select Case uMsg
        Case BFFM_INITIALIZED
            pathstring = StartDir
            RetVal = SendMessage(hWnd, BFFM_SETSELECTION, _
                     ByVal CLng(1), ByVal pathstring)
    End Select
 
    BrowseCallbackProc = 0
 
End Function

'======================================================================
' Public Function DummyFunc
'======================================================================

Public Function DummyFunc(ByVal param As LongPtr) As LongPtr
    DummyFunc = param
End Function