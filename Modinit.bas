Attribute VB_Name = "Modinit"

Public Sub Init()
Dim DX As New DirectX7
Dim dd As DirectDraw7

Set dd = DX.DirectDrawCreate("")

Set dd = Nothing
Set DX = Nothing


Call Setpaths
Call FilesLoad
Call Loadsettings




dhoom1.Canvas.Height = Canvas_height * Screen.TwipsPerPixelX
dhoom1.Canvas.Width = Canvas_width * Screen.TwipsPerPixelY

GdiFlush



Sound.Initialise
Sound.sON = True
Sound.mON = True

Road.Initialise
Player.Initialise
Display.Initialise

'Initial settings

Menu.Initialise GoneIn

Menu.Dowork Mainmenu
'Menu.AddMitem "GET READY TO FLY ON ROAD ", Message, 40, 20, 200
'Menu.AddMitem "OK", Ok, 30, 200


If Sound.mON = True Then
Sound.LoadSound "h.mid", 3, True, True
End If


'Sound.PlayWave "BK3M3", SND_ASYNC Or SND_LOOP

'Sound.PlayWave "MetalHt1", SND_ASYNC Or SND_LOOP Or SND_NOSTOP



End Sub
Private Sub FilesLoad()
Dim fso

Trs.temptrsname = App.path & "\pictures.trs"

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.folderexists(Picpath) = False Then
 sucess = fso.createfolder(Picpath)
End If
Set fso = Nothing



Trs.savefile "bike1.bmp", Picpath & "\bike1.bmp"
Trs.savefile "bike1mask.bmp", Picpath & "\bike1mask.bmp"
Trs.savefile "car1.bmp", Picpath & "\car1.bmp"
Trs.savefile "car1mask.bmp", Picpath & "\car1mask.bmp"
Trs.savefile "Pcar.bmp", Picpath & "\Pcar.bmp"
Trs.savefile "Pcarmask.bmp", Picpath & "\Pcarmask.bmp"
Trs.savefile "speedometer.bmp", Picpath & "\speedometer.bmp"
Trs.savefile "speedometermask.bmp", Picpath & "\speedometermask.bmp"
Trs.savefile "Title.bmp", Picpath & "\Title.bmp"
Trs.savefile "truck1.bmp", Picpath & "\truck1.bmp"
Trs.savefile "truck1mask.bmp", Picpath & "\truck1mask.bmp"
Trs.savefile "speedometermask.bmp", Picpath & "\speedometermask.bmp"








End Sub
Public Sub RemoveFiles()
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")




fso.deletefolder (Picpath)


Set fso = Nothing


End Sub

Private Sub Setpaths()

Picpath = App.path & "\pictures"
LevelPath = App.path & "\Levels.blv"
Sound.SoundPath = App.path & "\Sounds\"

End Sub
Private Sub Loadsettings()


MFps = 50
FpsLimiter = 10



End Sub
