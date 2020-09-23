Attribute VB_Name = "modGet_OS"
Option Explicit

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFOEX) As Long

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
    End Type
    Const VER_PLATFORM_WIN32s = 0
    Const VER_PLATFORM_WIN32_WINDOWS = 1
    Const VER_PLATFORM_WIN32_NT = 2
    Const VER_NT_WORKSTATION = 1
    Const VER_NT_SERVER = 3
    Const VER_SUITE_DATACENTER = 128
    Const VER_SUITE_ENTERPRISE = 2
    Const VER_SUITE_PERSONAL = 512

Public Function GetOSVer() As String
    Dim osv As OSVERSIONINFOEX
    osv.dwOSVersionInfoSize = Len(osv)

    If GetVersionEx(osv) = 1 Then

        Select Case osv.dwPlatformId
            Case Is = VER_PLATFORM_WIN32s
            GetOSVer = "Windows 3.x"
            Case Is = VER_PLATFORM_WIN32_WINDOWS

            Select Case osv.dwMinorVersion
                Case Is = 0

                If InStr(UCase(osv.szCSDVersion), "C") Then
                    GetOSVer = "Windows 95 OSR2"
                Else
                    GetOSVer = "Windows 95"
                End If
                Case Is = 10

                If InStr(UCase(osv.szCSDVersion), "A") Then
                    GetOSVer = "Windows 98 SE"
                Else
                    GetOSVer = "Windows 98"
                End If
                Case Is = 90
                GetOSVer = "Windows Me"
            End Select
        Case Is = VER_PLATFORM_WIN32_NT

        Select Case osv.dwMajorVersion
            Case Is = 3

            Select Case osv.dwMinorVersion
                Case Is = 0
                GetOSVer = "Windows NT 3"
                Case Is = 1
                GetOSVer = "Windows NT 3.1"
                Case Is = 5
                GetOSVer = "Windows NT 3.5"
                Case Is = 51
                GetOSVer = "Windows NT 3.51"
            End Select
        Case Is = 4
        GetOSVer = "Windows NT 4"
        Case Is = 5

        Select Case osv.dwMinorVersion
            Case Is = 0

            Select Case osv.wProductType
                Case Is = VER_NT_WORKSTATION
                GetOSVer = "Windows 2000 Professional"
                Case Is = VER_NT_SERVER

                Select Case osv.wSuiteMask
                    Case Is = VER_SUITE_DATACENTER
                    GetOSVer = "Windows 2000 DataCenter Server"
                    Case Is = VER_SUITE_ENTERPRISE
                    GetOSVer = "Windows 2000 Advanced Server"
                    Case Else
                    GetOSVer = "Windows 2000 Server"
                End Select
        End Select
    Case Is = 1

    Select Case osv.wProductType
        Case Is = VER_NT_WORKSTATION

        If osv.wSuiteMask = VER_SUITE_PERSONAL Then
            GetOSVer = "Windows XP Professional"
        Else
            GetOSVer = "Windows XP Home Edition"
        End If
        Case Else

        If osv.wSuiteMask = VER_SUITE_ENTERPRISE Then
            GetOSVer = "Windows .NET Enterprise Server"
        Else
            GetOSVer = "Windows .NET Server"
        End If
    End Select
End Select
End Select
End Select
End If
End Function



