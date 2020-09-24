VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Speed As Single
Public x As Single
Public y As Single
Public TurnSpeed As Single
Public Dir As Integer
Public Accl As Single
Public Width As Single
Public Height As Single
Public Dist As Single
Public Health As Single
Public Windist As Single
Public Level As Integer
Public VDamage As Single
Public PoliceSpeed As Single






Public Sub Initialise()

Call Loadpics
x = Road.x + 100
y = Road.y + 300

Width = 31
Height = 50

PoliceSpeed = 100


End Sub
Public Sub Terminate()



End Sub

Public Sub Update()
Call CheckKeys
Call Move
Call Checkwin
Call Render
Call PlaySound

End Sub

Private Sub Move()



End Sub
Private Sub CheckKeys()


Dir = 0
If GetKeystate(vbKeyLeft) = True Then
Turnleft

End If



If GetKeystate(vbKeyRight) = True Then
Turnright
End If


If GetKeystate(vbKeyUp) = True Then
If Speed < 250 Then
Speed = Speed + Accl
End If

End If

If GetKeystate(vbKeyDown) = True Then
If Speed > 0 Then
Speed = Speed - Accl
Else
Speed = 0
End If
End If


Speed = Speed - Road.Getfic(Speed)

Dist = Dist + (Speed - PoliceSpeed) / 100
'dhoom1.Label2.Caption = Dist


End Sub
Private Sub Render()
BitBlt Frmpics.Buffer.hdc, x, y, Width, Height, Frmpics.Bike1.hdc, Dir * Width, 0, vbSrcPaint
BitBlt Frmpics.Buffer.hdc, x, y, Width, Height, Frmpics.bike1mask.hdc, Dir * Width, 0, vbSrcAnd

End Sub
Private Sub Loadpics()

Frmpics.Bike1.Picture = LoadPicture(Picpath & "\bike1.bmp")
Frmpics.bike1mask.Picture = LoadPicture(Picpath & "\bike1mask.bmp")

End Sub

Public Sub Turnright()
Dir = 1

x = x + (TurnSpeed * (Speed / 100))
If x > Road.x + Road.Width - Width Then
x = Road.x + Road.Width - Width
End If

End Sub
Public Sub Turnleft()
Dir = 2
x = x - (TurnSpeed * (Speed / 100))
If x < Road.x Then
x = Road.x
End If
End Sub


Public Sub clearlevel()


For i = 1 To 10
Vehicle(i).Terminate
Next
Health = 100



End Sub

Public Sub Checkwin()


If Dist >= Windist Then
Win Outrun

End If

If Dist <= 0 Then
Lose Bypolice

End If

If Health <= 0 Then
Lose Bycrash

End If



End Sub
Private Sub PlaySound()

If Sound.sON = False Then Exit Sub

Select Case Speed

   Case Is < 80
       
         If LCase(Sound.GetFileName(1)) <> LCase(Sound.SoundPath & "BK3IDL.wav") Then
         
         Sound.LoadSound "BK3IDL.wav", 1, True, False
        End If
        Sound.SetVolume 1, ((Speed * 100) / 160) + 50
   
   Case Is < 200
         If LCase(Sound.GetFileName(1)) <> LCase(Sound.SoundPath & "BK3M1.wav") Then
         
         Sound.LoadSound "BK3M1.wav", 1, True, False
         
        End If
         Sound.SetVolume 1, ((Speed * 100) / 170)
   

End Select

'Sound.SetVolume 1, ((Speed * 100) / 210) + 10



End Sub
Public Sub LoadLevel(ByVal Levelno As Byte)



'**************************************************************************
Dim Temp As String
 Call clearlevel ' ' call Clearlevel Procedure of Player object to Clear current level

Health = 100

Level = Levelno
                 Temp = LevelSys.EntryRead("message", "notfound", "level " & Trim(Str(Levelno)), LevelPath)
                If Temp = "notfound" Then GoTo LevelFinished
                Menu.ClearMenu
                Menu.AddMitem Temp, Message, 40, 20, 150, 5
                Menu.AddMitem "OK", Ok, 30, 200, , , 60
                Gamestate = paused
                
                Temp = LevelSys.EntryRead("dist", "1000", "level" & Str(Levelno), LevelPath)
                Dist = Val(Temp)

                Temp = LevelSys.EntryRead("windist", "2500", "level" & Str(Levelno), LevelPath)
                Windist = Val(Temp)
                
                Temp = LevelSys.EntryRead("turnspeed", "1.5", "level" & Str(Levelno), LevelPath)
                TurnSpeed = Val(Temp)
                
                Temp = LevelSys.EntryRead("accl", "1.6", "level" & Str(Levelno), LevelPath)
                Accl = Val(Temp)
                
                Temp = LevelSys.EntryRead("vdamage", "0.2", "level" & Str(Levelno), LevelPath)
                VDamage = Val(Temp)
                
                 Temp = LevelSys.EntryRead("friction", "0.01", "level" & Str(Levelno), LevelPath)
                Road.Fic_dcl = Val(Temp)
                
                 Temp = LevelSys.EntryRead("pspeed", "100", "level" & Str(Levelno), LevelPath)
                PoliceSpeed = Val(Temp)
                
                Speed = 40
                
                For i = 1 To 10
                
                Temp = LevelSys.EntryRead("ntype", "notfound", "l" & Trim(Str(Levelno)) & "vehicleadd" & Trim(Str(i)), LevelPath)
                 If Temp = "notfound" Then Exit For
                 Dim T1 As String, T2 As String
                 
                 T1 = Val(LevelSys.EntryRead("timeout", "notfound", "l" & Trim(Str(Levelno)) & "vehicleadd" & Trim(Str(i)), LevelPath))
                 T2 = Val(LevelSys.EntryRead("rand", "notfound", "l" & Trim(Str(Levelno)) & "vehicleadd" & Trim(Str(i)), LevelPath))
                 
                 Vehicle(i).Initialise Val(Temp), T1, T2, i
                 
                Next
                
                
    Exit Sub
    
LevelFinished:

  Menu.ClearMenu
  Menu.AddMitem "GOOD WORK,ALL LEVELS CLEARED! ", Message, 40, 20, 300
  Menu.AddMitem "MAIN MENU", Mainmenu, 30, 200
  Gamestate = paused

 clearlevel

                
Exit Sub

                
                
End Sub
Private Sub Class_Initialize()

'End
End Sub
