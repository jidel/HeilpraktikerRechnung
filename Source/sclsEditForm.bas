Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Form As Form

Public Sub GotoRecord(ByVal Filter As Variant)

    With Form.RecordsetClone
      .FindFirst Filter
      If Not .NoMatch Then Form.Bookmark = .Bookmark
    End With

End Sub