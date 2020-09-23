Attribute VB_Name = "modFileSel"
Public Function f_ReturnFilePath(str_TreeView1Path As String) As String

    Dim int_Search(1) As Integer
    Dim str_RootPath As String
    
    int_Search%(0) = InStr(1, str_TreeView1Path$, "(", vbTextCompare)
    int_Search%(1) = InStr(1, str_TreeView1Path$, ")", vbTextCompare)
    
    If int_Search%(0) > 0 Then
        str_RootPath$ = Mid(str_TreeView1Path$, int_Search%(0) + 1, 2)
    End If
    
    If int_Search%(1) > 0 Then
        f_ReturnFilePath$ = str_RootPath$ & Mid(str_TreeView1Path$, int_Search%(1) + 1, Len(str_TreeView1Path$))
    End If
    
End Function

