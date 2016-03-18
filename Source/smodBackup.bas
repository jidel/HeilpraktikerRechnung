'======================================================================
' Modul Backup
'======================================================================

Option Compare Database
Option Explicit

Declare Function CopyFile& Lib "kernel32" Alias "CopyFileA" (ByVal _
lpExistingFilename As String, ByVal lbNewFileName As String, ByVal _
bFailIfExists As Long)

'======================================================================
'
' Public Sub CreateBackEndBackup()
'
' ---------------------------------------------------------------------
'
' Erzeugt ein Backup der Datenbank (ohne Formulare) an dem in den
' Einstellungen gespeicherten Ort.
'
'======================================================================

Public Function CreateBackEndBackup()

    Dim dbs As DAO.Database
    Dim strSQL As String
    Dim rs As DAO.Recordset
    Dim BackupDest As String
    
    'Backup-Location aus Einstellungen laden
    Set dbs = CurrentDb
    strSQL = "SELECT backuplocation FROM tblSettings WHERE id = 1 AND backuplocation IS NOT NULL AND backuplocation <> ''"
    Set rs = dbs.OpenRecordset(strSQL)
    
    If rs.EOF Then
        'kein Speicherort angegeben
        Err.Raise 512
        Exit Function
    Else
        rs.MoveFirst
        BackupDest = rs!BackupLocation
    End If
    
    BackUp getBackendLocation, BackupDest

End Function

'======================================================================
' Public Sub BackUp()
'======================================================================

Public Sub BackUp(DbFile As String, BackupDest As String)

    Dim BackupName As String
    Dim strMessage As String
    Dim dbs As DAO.Database
    Const FILEINUSE = 3704
    
    BackupName = BackupDest

    If Right(BackupDest, 1) <> "\" Then
        BackupName = BackupName & "\"
    End If
    
    BackupName = BackupName & Format(Now, "yyyy-mm-dd") & " " & Mid(DbFile, InStrRev(DbFile, "\") + 1) & ".bak"
    
    If MsgBox("Soll die Datenbank unter folgendem Pfad gesichert werden?" & vbCrLf & BackupName, vbYesNo + vbQuestion, "Backup erstellen?") = vbYes Then
    
        On Error Resume Next
        
        'Versuche Datenbankdatei zu öffnen
        Set dbs = OpenDatabase(DbFile, True)
        
        Select Case Err.Number
            Case 0
                'Keine Errors, also weiter
                On Error GoTo 0
                dbs.Close
                If Dir(BackupDest, vbDirectory) = vbNullString Then
                    'Backup-Verzeichnis nicht vorhanden
                    MsgBox "Das angegebene Backup-Verzeichnis '" & BackupDest & "' existiert nicht!" & vbCrLf & "Bitte erstellen Sie zuerst ein entsprechendes Verzeichnis.", vbCritical, "Backup fehlgeschlagen!"
                    Exit Sub
                End If
                
                'Datenbank kopieren
                If CopyFile(DbFile, BackupName, False) = 1 Then
                    MsgBox "Backup vollständig durchgeführt.", vbInformation, "Backup erfolgreich!"
                Else
                    MsgBox "Backup konnte nicht durchgeführt werden.", vbCritical, "Backup fehlgeschlagen!"
                End If
            
            Case FILEINUSE
                'Datei in Verwendung - User informieren
                strMessage = "Die Datei " & DbFile & _
                             " ist leider im Moment nicht Verfügbar." & vbCrLf & _
                             "Möglicherweise greift ein anderer " & _
                             "Nutzer darauf zu. "
                MsgBox strMessage, vbExclamation, "Backup nicht möglich"
            Case Else
                'unbekannter Fehler - User informieren
                MsgBox Err.Description, vbExclamation, "Error"
        End Select
    
    End If
    
End Sub


'======================================================================
' Public Function GetBackEndLocation()
'======================================================================

Public Function getBackendLocation()

    Dim strBackEnd As String
    Dim tdf As DAO.TableDef
    
    For Each tdf In CurrentDb.TableDefs
        If tdf.Connect <> "" Then
            strBackEnd = tdf.Connect
            strBackEnd = Mid(strBackEnd, InStr(strBackEnd, "DATABASE=") + 9)
        End If
    Next tdf
    
    getBackendLocation = strBackEnd

End Function