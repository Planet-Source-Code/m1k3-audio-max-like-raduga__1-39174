VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLoadDrives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub subLoadTreeView()

    Dim TreeList As TreeView
        Set TreeList = frmMAXmain.TreeView1
    Dim int_DrvNumber As Integer
    Dim str_DrvType As String
    Dim fso, d As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
    Dim x As Integer
    Dim RetStr(1) As String
    Dim nNode As Node
    Dim str_CompName As String
        str_CompName$ = modCompName.str_CompName
                        
    RetStr$(0) = f_DriveList

    With TreeList
        Set nNode = .Nodes.Add(, , str_CompName$, str_CompName$, "mycomputer", "mycomputer")
        For x = 1 To Len(RetStr$(0))
            RetStr$(1) = Mid(RetStr$(0), x, 1)
            Set d = fso.GetDrive(RetStr$(1))
            int_DrvNumber% = x
            Select Case d.drivetype
                
                'Unknown Drive.
                Case 0: str_DrvType = "Unknown"
                    If d.isready Then
                        Set nNode = .Nodes.Add(str_CompName$, tvwChild, "Parent" & int_DrvNumber%, d.volumename & " (" & d.driveletter & ":)", "drvunknown")
                    Else
                        Set nNode = .Nodes.Add(str_CompName$, tvwChild, "Parent" & int_DrvNumber%, " (" & d.driveletter & ":)", "drvunknown")
                    End If

                'Removable Drive.
                Case 1: str_DrvType = "Removable"
                    'If d.isready Then
                        'Set nNode = .Nodes.Add(str_CompName$, tvwChild, "Parent" & int_DrvNumber%, d.volumename & " (" & d.driveletter & ":", "drvremove")
                    'Else
                        'Set nNode = .Nodes.Add(str_CompName$, tvwChild, "Parent" & int_DrvNumber%, " (" & d.driveletter & ":)", "drvremove")
                    'End If
                    
                'Fixed Drive.
                Case 2: str_DrvType = "Fixed"
                    If d.isready Then
                        Set nNode = .Nodes.Add(str_CompName$, tvwChild, "Parent" & int_DrvNumber%, d.volumename & " (" & d.driveletter & ":)", "drvfixed")
                    Else
                        Set nNode = .Nodes.Add(str_CompName$, tvwChild, "Parent" & int_DrvNumber%, " (" & d.driveletter & ":)", "drvfixed")
                    End If

                'Network Drive.
                Case 3: str_DrvType = "Network"
                    If d.isready Then
                        Set nNode = .Nodes.Add(str_CompName$, tvwChild, "Parent" & int_DrvNumber%, d.volumename & " (" & d.driveletter & ":)", "drvremote")
                    Else
                        Set nNode = .Nodes.Add(str_CompName$, tvwChild, "Parent" & int_DrvNumber%, " (" & d.driveletter & ":)", "drvremote")
                    End If
                    
                'CD-Rom.
                Case 4: str_DrvType = "CD-ROM"
                    If d.isready Then
                        Set nNode = .Nodes.Add(str_CompName$, tvwChild, "Parent" & int_DrvNumber%, d.volumename & " (" & d.driveletter & ":)", "drvcd")
                    Else
                        Set nNode = .Nodes.Add(str_CompName$, tvwChild, "Parent" & int_DrvNumber%, " (" & d.driveletter & ":)", "drvcd")
                    End If
                    
                'Ram Disk.
                Case 5: str_DrvType = "Ram Disk"
                    If d.isready Then
                        Set nNode = .Nodes.Add(str_CompName$, tvwChild, "Parent" & int_DrvNumber%, d.volumename & " (" & d.driveletter & ":)", "drvram")
                    Else
                        Set nNode = .Nodes.Add(str_CompName$, tvwChild, "Parent" & int_DrvNumber%, " (" & d.driveletter & ":)", "drvram")
                    End If
            End Select
        Next x
    End With

Set TreeList = Nothing
Set nNode = Nothing

End Sub

Private Function f_DriveList()
  
    Dim fso, d, dc As Object
    Dim sDriveLetter As String
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set dc = fso.Drives
    
    For Each d In dc
        sDriveLetter$ = sDriveLetter$ & d.driveletter
    Next
    
    f_DriveList = sDriveLetter$

End Function


