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

    LoadBackendDB

    DoCmd.OpenForm "frmMain", acNormal, , , , acHidden

End Function

'======================================================================
' Public Function Main()
'======================================================================

Public Sub LoadBackendDB()

    On Error GoTo MyError

    Dim db As DAO.Database
    Dim strDaten As String
    Dim i As Integer
    
    Set db = CurrentDb()
        
    strDaten = Left(db.Name, Len(db.Name) - Len(Dir(db.Name))) & "Backend/Datenbank.accdb"
    
    For i = 0 To db.TableDefs.Count - 1
      If db.TableDefs(i).Connect <> "" Then
        If Mid(db.TableDefs(i).Connect, 11) <> strDaten Then
          db.TableDefs(i).Connect = ";database=" & strDaten
          db.TableDefs(i).RefreshLink
        End If
      End If
    Next i
    
MyExit:
      Exit Sub
    
MyError:
      MsgBox "Bei der Installation ist eine Ausnahme aufgetreten: " & strDaten, 16, "Ausnahme"
      Resume MyExit

End Sub