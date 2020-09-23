Attribute VB_Name = "modStripNull"
Public Function f_StripNullChr(sInput As String) As String
    Dim x As Integer
    x = InStr(1, sInput$, Chr$(0))
    If x > 0 Then
        f_StripNullChr = Left(sInput$, x - 1)
    End If
End Function

