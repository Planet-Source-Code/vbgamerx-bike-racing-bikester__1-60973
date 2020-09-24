VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Style As Vehicletype
Public Alive As Boolean
Public x As Single
Public y As Single
Public Width As Single
Public Height As Single
Public Speed As Single
Public Timeout As Integer
Public RandomValue As Single
Public VNo As Byte





Enum Vehicletype

Truck1 = 1
car1 = 2

End Enum

Public Sub Initialise(ByVal sStyle As Vehicletype, ByVal tTimeout As Integer, ByVal Rand As Single, ByVal VeNo As Byte)
Alive = True
Style = sStyle

Sound.LoadOnly "MetalHt2.wav", 4
VNo = VeNo
RandomValue = Rand
Timeout = tTimeout
Select Case Style

    Case 1
       If Frmpics.Truck1.Picture = 0 Then
        Frmpics.Truck1.Picture = LoadPicture(Picpath & "\truck1.bmp")
        Frmpics.Truck1mask.Picture = LoadPicture(Picpath & "\truck1mask.bmp")
        End If
        Width = Frmpics.Truck1.Width / Screen.TwipsPerPixelX
        Height = Frmpics.Truck1.Height / Screen.TwipsPerPixelY
        Speed = 50
        
    Case 2
        Frmpics.car1.Picture = LoadPicture(Picpath & "\car1.bmp")
        Frmpics.car1mask.Picture = LoadPicture(Picpath & "\car1mask.bmp")
        Width = Frmpics.car1.Width / Screen.TwipsPerPixelX
        Height = Frmpics.car1.Height / Screen.TwipsPerPixelY
        Speed = 70
       
        
End Select

Call Getcood

End Sub
Public Sub Terminate()
Alive = False


End Sub
Public Sub Update()


Call Collide
Call Move

Call Render



End Sub
Private Sub Move()

y = y - ((Speed - Player.Speed) / 15)



End Sub
Private Sub Collide()

If Speed < Player.Speed Then
If y > Road.Height + Timeout Then
Call Getcood
End If
Else
If y + Height < -Timeout Then
Call Getcood
End If
End If

If Intersect(x, y, Width, Height, Player.x, Player.y, Player.Width, Player.Height) = True Then


If Sound.sON Then
Sound.SetVolume 4, 75
Sound.PlaySound 4, 1, False
End If



If Player.y > y And Player.Speed > Speed Then
Player.Speed = Speed
End If

If Player.y < y And Player.Speed < Speed Then
Player.Speed = Speed
End If

If Player.x > ((x + Width / 2) - (Player.Width / 2)) Then
Player.Dir = 1

Player.Turnright
Else
Player.Dir = 2
'Player.TurnSpeed = 0.5
Player.Turnleft

End If
Dim Te As Single
Te = (Player.Speed - Speed) / 10

If Te >= 0 Then
Player.Health = Player.Health - Te - Player.VDamage

Else
Player.Health = Player.Health + Te - Player.VDamage
End If






Else



End If



End Sub
Private Sub Render()


    Select Case Style
    
    Case 1
        If y < Road.Height Then
        BitBlt Frmpics.Buffer.hdc, x, y, Width, Height, Frmpics.Truck1.hdc, 0, 0, vbSrcPaint
         BitBlt Frmpics.Buffer.hdc, x, y, Width, Height, Frmpics.Truck1mask.hdc, 0, 0, vbSrcAnd

    End If
    
    Case 2
         If y < Road.Height Then
        BitBlt Frmpics.Buffer.hdc, x, y, Width, Height, Frmpics.car1.hdc, 0, 0, vbSrcPaint
        BitBlt Frmpics.Buffer.hdc, x, y, Width, Height, Frmpics.car1mask.hdc, 0, 0, vbSrcAnd

        End If
    
    
    
    End Select
    
End Sub
Public Sub Getcood()

Select Case Style

    Case 1
        Dim T As Byte

        T = Rnd * 1


        If T = 0 Then
        x = Road.x + 10
        Else
        x = Road.x + Road.Lane2_x + 20
        End If
        
    Case 2
    
       x = Road.x + Road.Lane1_x + 20
       
       
End Select



Randomize

If Speed < Player.Speed Then
y = 0 - (Rnd * (RandomValue + Height)) - Timeout - Height
Else

y = Road.Height + (Rnd * (RandomValue + Height)) + Timeout
End If


For i = 1 To 10

 If i <> VNo Then
If Vehicle(i).Alive = True Then


If Intersect(Vehicle(i).x, Vehicle(i).y, Vehicle(i).Width, Vehicle(i).Height, x, y, Width, Height) = True Then

Vehicle(VNo).Getcood

End If

End If

End If


Next i





End Sub
