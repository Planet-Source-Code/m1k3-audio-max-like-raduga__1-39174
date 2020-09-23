Attribute VB_Name = "modAddFolders"
Public Function FolderName(FPath As String) As String
    FolderName = StrReverse(FPath)
    FolderName = StrReverse(Mid(FolderName, 1, InStr(FolderName, "\") - 1))
End Function

Function FileFound(strFileName As String) As Boolean
    Dim lpFindFileData As WIN32_FIND_DATA
    Dim hFindFirst As Long
    hFindFirst = FindFirstFile(strFileName, lpFindFileData)

    If hFindFirst > 0 Then
        FindClose hFindFirst
        FileFound = True
    Else
        FileFound = False
    End If
End Function
Private Function FileList(ByVal PathName As String, Optional DirCount As Long, Optional FileCount As Long) As String
    
    Dim ShortName As String, LongName As String
    Dim NextDir As String
    Static FolderList As Collection
    Screen.MousePointer = vbHourglass

    If FolderList Is Nothing Then
        Set FolderList = New Collection
        FolderList.Add PathName
        DirCount = 0
        FileCount = 0
    End If


    Do
    NextDir = FolderList.Item(1)
    FolderList.Remove 1
    ShortName = Dir(NextDir & "\*.*", vbNormal Or _
    vbArchive Or _
    vbDirectory)

    Do While ShortName > ""

        If ShortName = "." Or ShortName = ".." Then
        Else
            LongName = NextDir & "\" & ShortName

            If (GetAttr(LongName) And vbDirectory) > 0 Then
                FolderList.Add LongName
                DirCount = DirCount + 1
            Else
                FileList = FileList & LongName & vbCrLf
                FileCount = FileCount + 1
            End If
        End If
        ShortName = Dir()
    Loop
Loop Until FolderList.Count = 0
Screen.MousePointer = vbNormal
End Function



