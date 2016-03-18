Option Compare Database
Option Explicit

Public Sub Callback_RibbonOnAction(RibbonControl As IRibbonControl)

    Dim ID As Variant
    
    'ID übernehmen und anpassen
    ID = RibbonControl.ID
    If InStr(ID, "___") > 0 Then ID = Mid(ID, InStr(ID, "___") + 3)
    If Left(ID, 3) = "btn" Then ID = Mid(ID, 4)
    
    On Error Resume Next
    
    'Je nach Aktion...
    Select Case Left(ID, 3)
        Case "frm" 'Entsprechendes Formular öffnen
            If InStr(ID, "Report") > 0 Then
                DoCmd.OpenForm ID, acFormDS
            Else
                DoCmd.OpenForm ID
            End If
        Case "rpt" 'Entsprechenden Bericht öffnen
            DoCmd.OpenReport ID, View:=acViewPreview
        Case "mcr" 'Entsprechendes Makro ausführen
            ID = Mid(ID, 4)
            DoCmd.RunMacro "mcrRibbon." & ID
        Case "run" 'Entsprechendes Makro ausführen
            ID = Replace(Mid(ID, 4), "_", ".")
            DoCmd.RunMacro ID
        Case Else
            'Weitere Aktionen ausführen...
            Select Case ID
                Case "FormView"
                    DoCmd.RunCommand acCmdFormView
                Case "DatasheetView"
                    DoCmd.RunCommand acCmdDatasheetView
                Case "PrintPreview"
                    DoCmd.RunCommand acCmdPrintPreview
                Case "Close"
                    SendKeys "^{F4}" 'STRG+F4
                Case "Exit", "Quit"
                    DoCmd.Quit
            End Select
    End Select
            
    If Err.Number <> 0 Then
        Rem MsgBox Err.Description
        Err.Clear
    End If
    On Error GoTo 0


End Sub