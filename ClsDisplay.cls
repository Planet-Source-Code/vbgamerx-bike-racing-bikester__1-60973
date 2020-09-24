VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Meter_width As Single
Public Meter_height As Single
Public Meter_X As Single
Public Meter_Y As Single
Public Health_X As Single
Public Health_Y As Single
Public Health_Width As Single
Public Health_Height As Single
Public Dist_X As Single
Public Dist_Y As Single
Public Dist_Width As Single
Public Dist_Height As Single

Public Sub Initialise()


Frmpics.Smeter.Picture = LoadPicture(Picpath & "\speedometer.bmp")
Frmpics.Smetermask.Picture = LoadPicture(Picpath & "\speedometermask.bmp")
Meter_X = 500
Meter_Y = 400
Meter_width = Frmpics.Smeter.Width / Screen.TwipsPerPixelX
Meter_height = Frmpics.Smeter.Height / Screen.TwipsPerPixelY
Health_X = 20
Health_Y = 425
Health_Width = 100
Health_Height = 15
Dist_X = 20
Dist_Y = 460
Dist_Width = 130
Dist_Height = 5


Frmpics.Pcar.Picture = LoadPicture(Picpath & "\pcar.bmp")
Frmpics.Pcarmask.Picture = LoadPicture(Picpath & "\pcarmask.bmp")



End Sub
Public Sub Update()

Call Render

End Sub



Private Sub Render()

Frmpics.Buffer.FillColor = 0
Frmpics.Buffer.ForeColor = vbBlack
Rectangle Frmpics.Buffer.hdc, 0, Road.Height, Canvas_width, Canvas_height

Call RenderSmeter
Call Renderhealth
Call RenderDist

End Sub
Private Sub RenderSmeter()

dhoom1.Label2.Caption = Player.Speed
BitBlt Frmpics.Buffer.hdc, Meter_X, Meter_Y, Meter_width, Meter_height, Frmpics.Smeter.hdc, 0, 0, vbSrcAnd
BitBlt Frmpics.Buffer.hdc, Meter_X, Meter_Y, Meter_width, Meter_height, Frmpics.Smetermask.hdc, 0, 0, vbSrcPaint

Frmpics.Buffer.ForeColor = vbBlack
T = (Cos(angtorad((Player.Speed / 100) * 90)))
Frmpics.Buffer.DrawWidth = 2
GeDrawLine Frmpics.Buffer.hdc, Meter_X + Meter_width / 2 - (Sin(angtorad((Player.Speed / 200) * 180))) * ((Meter_width / 2) - 15), Meter_Y + Meter_height / 2 + Cos(angtorad(Player.Speed / 200 * 180)) * ((Meter_height / 2) - 15), (Meter_X + Meter_width / 2), (Meter_Y + Meter_height / 2)

'GeDrawLine Frmpics.Buffer.hdc, 0, 0, 300, 300


End Sub
Private Sub Renderhealth()

If Player.Health > 0 Then
Frmpics.Buffer.FillColor = vbRed
Frmpics.Buffer.ForeColor = vbWhite
Frmpics.Buffer.DrawWidth = 1
Frmpics.Buffer.FillStyle = 0

Rectangle Frmpics.Buffer.hdc, Health_X, Health_Y, Health_X + 20, Health_Y + 5
Rectangle Frmpics.Buffer.hdc, Health_X + 8, Health_Y - 10, Health_X + 12, Health_Y + 15

Frmpics.Buffer.FillStyle = 1
Rectangle Frmpics.Buffer.hdc, Health_X + 30, Health_Y, Health_X + Health_Width + 30, Health_Y + Health_Height



Frmpics.Buffer.FillColor = RGB((255 - (Player.Health * 255 / 100)), Player.Health * 255 / 100, 0)

Frmpics.Buffer.FillStyle = 0
Rectangle Frmpics.Buffer.hdc, Health_X + 30, Health_Y, Health_X + (Health_Width) * (Player.Health / 100) + 30, Health_Y + Health_Height

Frmpics.Buffer.CurrentX = (Health_X + 35) * Screen.TwipsPerPixelX
Frmpics.Buffer.CurrentY = Health_Y * Screen.TwipsPerPixelY

Frmpics.Buffer.ForeColor = vbBlue
Frmpics.Buffer.Print Int(Player.Health) & " %"



Else

End If

End Sub
Private Sub RenderDist()
Dim Percent As Single


Percent = (Player.Dist / Player.Windist) * 100

If Percent > 0 And Percent <= 100 Then
Frmpics.Buffer.FillColor = RGB((255 - (Percent * 255 / 100)), Percent * 255 / 100, 0)

Frmpics.Buffer.FillStyle = 0
Rectangle Frmpics.Buffer.hdc, Dist_X, Dist_Y + 20, Dist_X + Dist_Width + 20, Dist_Y + Dist_Height + 20


BitBlt Frmpics.Buffer.hdc, Dist_X + Dist_Width - 30 - Dist_Width * Percent / 100, Dist_Y, 40, 40, Frmpics.Pcar.hdc, 0, 0, vbSrcAnd
BitBlt Frmpics.Buffer.hdc, Dist_X + Dist_Width - 30 - Dist_Width * Percent / 100, Dist_Y, 40, 40, Frmpics.Pcarmask.hdc, 0, 0, vbSrcPaint

BitBlt Frmpics.Buffer.hdc, Dist_X + Dist_Width, Dist_Y, 40, 40, Frmpics.Pcar.hdc, 40, 0, vbSrcAnd
BitBlt Frmpics.Buffer.hdc, Dist_X + Dist_Width, Dist_Y, 40, 40, Frmpics.Pcarmask.hdc, 40, 0, vbSrcPaint


Frmpics.Buffer.CurrentX = (Dist_X) * Screen.TwipsPerPixelX
Frmpics.Buffer.CurrentY = (Dist_Y + 30) * Screen.TwipsPerPixelY

Frmpics.Buffer.ForeColor = vbWhite
Frmpics.Buffer.Print Int(Player.Dist) & "Meters"


Else




End If


End Sub
