Attribute VB_Name = "modGetFileName"
Public Function GetFileName(sFullPath As String) As String

    If InStr(1, sFullPath$, "/") > 0 Then
        GetFileName$ = Mid$(sFullPath$, InStrRev(sFullPath$, "/") + 1)
    ElseIf InStr(1, sFullPath$, "\") > 0 Then
        GetFileName$ = Mid$(sFullPath$, InStrRev(sFullPath$, "\") + 1)
    Else
        GetFileName$ = sFullPath$
    End If
End Function

