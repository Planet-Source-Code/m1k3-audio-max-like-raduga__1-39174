Attribute VB_Name = "modMP3info"
Public MP3FileName As String

Public Type VBRinfo
  VBRrate As String
End Type

Public Type MP3Info
  BITRATE As String
  CHANNELS As String
  FREQ As String
  LAYER As String
  MPEG As String
End Type

Private MP3File As String

Public Sub getMP3Info(ByVal lpMP3File As String, ByRef lpMP3Info As MP3Info)
  Dim Buf As String * 4096
  Dim infoStr As String * 3
  Dim lpVBRinfo As VBRinfo
  Dim tmpByte As Byte
  Dim tmpNum As Byte
  Dim i As Integer
  Dim designator As Byte
  Dim baseFreq As Single
  Dim vbrBytes As Long
  
  Open lpMP3File For Binary As #1
    Get #1, 1, Buf
  Close #1
  
  For i = 1 To 4092
    If Asc(Mid(Buf, i, 1)) = &HFF Then
      tmpByte = Asc(Mid(Buf, i + 1, 1))
      If Between(tmpByte, &HF2, &HF7) Or Between(tmpByte, &HFA, &HFF) Then
        Exit For
      End If
    End If
  Next i
  If i = 4093 Then
    Exit Sub
  Else
    infoStr = Mid(Buf, i + 1, 3)
    tmpByte = Asc(Mid(infoStr, 1, 1))
    
    If Between(tmpByte, &HF2, &HF7) Then
      lpMP3Info.MPEG = "MPEG 2.0"
      designator = 1
    Else
      lpMP3Info.MPEG = "MPEG 1.0"
      designator = 2
    End If
    
    If Between(tmpByte, &HF2, &HF3) Or Between(tmpByte, &HFA, &HFB) Then
      lpMP3Info.LAYER = "Layer 3"
    Else
      If Between(tmpByte, &HF4, &HF5) Or Between(tmpByte, &HFC, &HFD) Then
        lpMP3Info.LAYER = "Layer 2"
      Else
        lpMP3Info.LAYER = "Layer 1"
      End If
    End If
    
    tmpByte = Asc(Mid(infoStr, 2, 1))
    
    If Between(tmpByte Mod 16, &H0, &H3) Then
      baseFreq = 22.05
    Else
      If Between(tmpByte Mod 16, &H4, &H7) Then
        baseFreq = 24
      Else
        baseFreq = 16
      End If
    End If
    lpMP3Info.FREQ = baseFreq * designator * 1000 & " Hz"
    
    tmpNum = tmpByte \ 16 Mod 16
    If designator = 1 Then
      If tmpNum < &H8 Then
        lpMP3Info.BITRATE = tmpNum * 8
      Else
        lpMP3Info.BITRATE = 64 + (tmpNum - 8) * 16
      End If
    Else
      If tmpNum <= &H5 Then
        lpMP3Info.BITRATE = (tmpNum + 3) * 8
      Else
        If tmpNum <= &H9 Then
          lpMP3Info.BITRATE = 64 + (tmpNum - 5) * 16
        Else
          If tmpNum <= &HD Then
            lpMP3Info.BITRATE = 128 + (tmpNum - 9) * 32
          Else
            lpMP3Info.BITRATE = 320
          End If
        End If
      End If
    End If
    If Mid(Buf, i + 36, 4) = "Xing" Then
      vbrBytes = Asc(Mid(Buf, i + 45, 1)) * &H10000
      vbrBytes = vbrBytes + (Asc(Mid(Buf, i + 46, 1)) * &H100&)
      vbrBytes = vbrBytes + Asc(Mid(Buf, i + 47, 1))
      GetVBRrate lpMP3File, vbrBytes, lpVBRinfo
      lpMP3Info.BITRATE = lpVBRinfo.VBRrate
    Else
      lpMP3Info.BITRATE = lpMP3Info.BITRATE & "Kbit"
    End If
    
    tmpByte = Asc(Mid(infoStr, 3, 1))
    tmpNum = tmpByte Mod 16
    
    tmpNum = (tmpByte \ 16) \ 4
    Select Case tmpNum
      Case 0
        lpMP3Info.CHANNELS = " Stereo"
      Case 1
        lpMP3Info.CHANNELS = " Joint Stereo"
      Case 2
        lpMP3Info.CHANNELS = " 2 Channel"
      Case 3
        lpMP3Info.CHANNELS = " Mono"
    End Select
  End If
End Sub

Private Sub GetVBRrate(ByVal lpMP3File As String, ByVal byteRead As Long, ByRef lpVBRinfo As VBRinfo)
  Dim i As Long
  Dim dDone As Boolean

  i = 0
  byteRead = byteRead - &H39
  Do
    If byteRead > 0 Then
      i = i + 1
      byteRead = byteRead - 38 - dBit(i)
    Else
      dDone = True
    End If
  Loop Until dDone
  lpVBRinfo.VBRrate = Trim(Str(Int(8 * FileLen(lpMP3File) / (1000 * i)))) & " Kbit (VBR)"
End Sub

Private Function dBit(ByVal Num As Long) As Byte
  If Num Mod 3 = 0 Then
    dBit = 1
  Else
    dBit = 0
  End If
End Function

Public Function Between(ByVal accNum As Byte, ByVal accDown As Byte, ByVal accUp As Byte) As Boolean
  If accNum >= accDown And accNum <= accUp Then
    Between = True
  Else
    Between = False
  End If
End Function

