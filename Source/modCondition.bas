'========================================================================
' Public Function Modul Condition
'========================================================================

Option Compare Database
Option Explicit

'========================================================================
' Public Function Condition_Numeric_EqualOrLike
'========================================================================

Public Function Condition_Varchar_EqualOrLike(ByVal Field As String, ByVal Condition As String) As String

    '20140820: Suche ab jetzt nur mit Wildcards!
    
    'If InStr(Condition, "*") = 0 And InStr(Condition, "?") = 0 Then
    '    Condition_Varchar_EqualOrLike = Field & "='" & Condition & "'"
    'Else
        Condition_Varchar_EqualOrLike = Field & " LIKE '*" & Condition & "*'"
    'End If

End Function