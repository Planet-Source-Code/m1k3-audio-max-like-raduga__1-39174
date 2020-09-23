Attribute VB_Name = "modGetMP2"
Option Explicit
           
Public Sub subMP2List(oFileList As ListBox, oTreeView As TreeView, sFolderPath As String, vParentID As Variant)
    
    Dim nNode As Node
    Dim lReturn As Long
    Dim lNextFile As Long
    Dim sPath As String
     Set oFileList = frmMAXmain.List2
     Set oTreeView = frmMAXmain.TreeView1
    Dim WFD As WIN32_FIND_DATA
    Dim lst_Item As ListItem
    Dim lst_SubItem As ListSubItem
    Dim sFileName As String
    Dim xLoop As Long
    Dim str_File As String
    
    sPath$ = sFolderPath$ & "*.mp2"
    
    lReturn& = FindFirstFile(sPath$, WFD) & Chr$(0)
    frmMAXmain.MousePointer = 11
    
    With oFileList
                
        Do
            If Not (WFD.dwFileAttributes And vbDirectory) = vbDirectory Then
                sFileName$ = modStripNull.f_StripNullChr(WFD.cFileName)
                If sFileName > Trim("") Then
                        oFileList.AddItem sFileName$
                End If
            End If
            lNextFile& = FindNextFile(lReturn&, WFD)
        Loop Until lNextFile& <= Val(0)

        frmMAXmain.MousePointer = 0
    
        lNextFile& = FindClose(lReturn&)
        
       For xLoop = 0 To oFileList.ListCount
        If (oFileList.List(xLoop)) <> "" Then
        
       Set nNode = oTreeView.Nodes.Add(vParentID, tvwChild, , Left(oFileList.List(xLoop), Len(oFileList.List(xLoop)) - 0), "api")
       End If
       Next
       
       oFileList.Clear
       
    End With
 Set oTreeView = Nothing
 Set oFileList = Nothing
End Sub



