'======================================================================
' Modul Main
'======================================================================

Option Compare Database
Option Explicit

'========== Öffentliche Konstanten ==========

Public Const App_Name As String = "ArztpraxisDB"
Public Const App_Title As String = "Heilpraktiker Rechnungsstellung"
Public Const App_Version As String = "v1.5"
Public Const App_Version_Date As String = "20.06.2025"

'======================================================================
' Public Function Main()
'======================================================================

Public Function Main() As Boolean
    DoCmd.OpenForm "frmMain", acNormal, , , , acHidden
End Function