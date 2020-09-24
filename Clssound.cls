VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clssound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Dim SStream(1 To 10) As MediaPlayer
Public SoundPath As String
Public sON As Boolean
Public mON As Boolean




Public Sub Initialise()



For i = 1 To 10
'Set SStream = New MediaPlayer
Load Frmpics.SStream(i)
Set SStream(i) = Frmpics.SStream(i)
'SStream(i) = Frmpic s.SStream(i)


Next



End Sub

Public Sub Terminate()

For i = 1 To 10

Unload SStream(i)

Next



End Sub


Public Sub LoadSound(ByVal filename As String, ByVal Stream As Byte, ByVal Repeat As Boolean, ByVal Stopprev As Boolean)

With SStream(Stream)

 .filename = SoundPath & filename
 If Stopprev = True Then
 .Stop
 End If
 
 If Repeat = True Then
 .PlayCount = 0
 End If
 
 .Play
End With


End Sub
Public Sub LoadOnly(ByVal filename As String, ByVal Stream As Byte)

With SStream(Stream)

 .filename = SoundPath & filename
 .Stop
End With


End Sub
Public Function GetFileName(ByVal Stream As Byte) As String

GetFileName = SStream(Stream).filename

End Function
Public Sub StopSound(ByVal Stream As Byte)

SStream(Stream).Stop

End Sub
Public Sub PlaySound(ByVal Stream As Byte, ByVal Rno As Byte, ByVal Stopprev As Boolean)
On Error Resume Next

If Stopprev = True Then
SStream(Stream).Stop
End If

SStream(Stream).PlayCount = Rno
SStream(Stream).Play
T = SStream(Stream).filename

End Sub
Public Sub Update()





End Sub

Public Sub SetVolume(ByVal Stream As Byte, ByVal Percent As Single)


If Percent >= 0 And Percent <= 100 Then
With SStream(Stream)


.Volume = (-1500) + (Percent * 1500) / 100

End With
End If



End Sub



