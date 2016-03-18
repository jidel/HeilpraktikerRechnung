'======================================================================
' Modul Main
'======================================================================

Option Compare Database
Option Explicit

'========== Öffentliche Konstanten ==========

Public Const App_Name As String = "ArztpraxisDB"
Public Const App_Title As String = "Arztpraxis Datenbank"
Public Const App_Version As String = "V1.4"
Public Const App_Version_Date As String = "20.03.2016"

'======================================================================
' Public Function Main()
'======================================================================

Public Function Main()
    DoCmd.OpenForm "frmMain", acNormal, , , , acHidden
End Function