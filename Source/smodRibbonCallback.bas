Option Compare Database
Option Explicit

Public Sub Callback_RibbonOnAction(RibbonControl As IRibbonControl)

    Dim id As Variant
    
    'ID übernehmen und anpassen
    id = RibbonControl.id
    If InStr(id, "___") > 0 Then id = Mid(id, InStr(id, "___") + 3)
    If Left(id, 3) = "btn" Then id = Mid(id, 4)
    
    On Error Resume Next
    
    'Je nach Aktion...
    Select Case Left(id, 3)
        Case "frm" 'Entsprechendes Formular öffnen
            If InStr(id, "Report") > 0 Then
                DoCmd.OpenForm id, acFormDS
            Else
                DoCmd.OpenForm id
            End If
        Case "rpt" 'Entsprechenden Bericht öffnen
            DoCmd.OpenReport id, View:=acViewPreview
        Case "mcr" 'Entsprechendes Makro ausführen
            id = Mid(id, 4)
            DoCmd.RunMacro "mcrRibbon." & id
        Case "run" 'Entsprechendes Makro ausführen
            id = Replace(Mid(id, 4), "_", ".")
            DoCmd.RunMacro id
        Case Else
            'Weitere Aktionen ausführen...
            Select Case id
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