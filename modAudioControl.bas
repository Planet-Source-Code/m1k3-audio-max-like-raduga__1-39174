Attribute VB_Name = "modAudioControl"
Public Sub AControl()

If MediaPosition1.CurrentPosition > MediaPosition2.CurrentPosition Then
   AudioXtoY
Else
If MediaPosition2.CurrentPosition > MediaPosition1.CurrentPosition Then
   Audio YtoX
   
End If
End If

End Sub



Private Function AudioXtoY()

Do
frmBaehr.Vol1.Value = frmBaehr.Vol1.Value - 1
frmBaehr.Vol2.Value = frmBaehr.Vol2.Value + 1
For tLoop = 1 To 39500
Next tLoop
Loop Until frmBaehr.Vol1.Value = -3950
MediaControl1.Stop
Set MediaControl1 = Nothing
Set Audio1 = Nothing
Set MediaPosition1 = Nothing

End Function

Private Function AudioYtoX()

Do
frmBaehr.Vol1.Value = frmBaehr.Vol1.Value + 1
frmBaehr.Vol2.Value = frmBaehr.Vol2.Value - 1
For tLoop = 1 To 39500
Next tLoop
Loop Until frmBaehr.Vol1.Value = -3950
MediaControl2.Stop
Set MediaControl2 = Nothing
Set Audio2 = Nothing
Set MediaPosition2 = Nothing

End Function

