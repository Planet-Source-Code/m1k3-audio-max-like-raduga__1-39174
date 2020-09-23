Attribute VB_Name = "modGetFolders"
Option Explicit

Public Sub subFolderList(oFolderList As ListBox, oTreeView As TreeView, sDriveLetter As String, vParentID As Variant)
    
    Dim nNode As Node
    Dim lReturn As Long
    Dim lNextFile As Long
    Dim sPath As String
    Dim WFD As WIN32_FIND_DATA
    Dim sFolderName As String
    Dim x As Long
     Set oFolderList = frmMAXmain.List1
     Set oTreeView = frmMAXmain.TreeView1
        
    sPath$ = (sDriveLetter & "*.*") & Chr$(0)
    
    lReturn& = FindFirstFile(sPath$, WFD)
    frmMAXmain.MousePointer = 11
    
    Do
        If (WFD.dwFileAttributes And vbDirectory) Then
            sFolderName$ = modStripNull.f_StripNullChr(WFD.cFileName)
            If sFolderName$ <> "." And sFolderName$ <> ".." Then
                If WFD.dwFileAttributes <> 16 Then
                    oFolderList.AddItem sFolderName$ & "~A~"
                Else
                    oFolderList.AddItem sFolderName$ & "~~~"
                End If
            End If
        End If
        lNextFile& = FindNextFile(lReturn&, WFD)
    Loop Until lNextFile& = False
    
    frmMAXmain.MousePointer = 0
    
    lNextFile& = FindClose(lReturn&)

    For x = 0 To oFolderList.ListCount - 1

        If Right(oFolderList.List(x), 3) = "~A~" Then
            Set nNode = oTreeView.Nodes.Add(vParentID, tvwChild, , Left(oFolderList.List(x), Len(oFolderList.List(x)) - 3), "cldfolder", "opnfolder")
            nNode.ForeColor = RGB(120, 120, 120)
        Else
            Set nNode = oTreeView.Nodes.Add(vParentID, tvwChild, , Left(oFolderList.List(x), Len(oFolderList.List(x)) - 3), "cldfolder", "opnfolder")
        End If
    Next x

    oFolderList.Clear
    
    Set oFolderList = Nothing
    Set oTreeView = Nothing
    
End Sub





