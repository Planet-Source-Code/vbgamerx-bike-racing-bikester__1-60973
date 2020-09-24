VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clsmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public x As Single
Public y As Single
Public Width As Single
Public Height As Single
Public Menustate As eMenustate
Public Mspeed As Single
Public SMno As Integer
Dim MItem(1 To 10) As MenuItem
'Dim mScrollDelay As Byte
Enum eMenustate

GoneIn = 1
GoingIn = 2
ComeOut = 3
CominOut = 4

End Enum
Enum eTask

Message = 1
Ok = 2
NewGame = 3
Quit = 4
Mainmenu = 5
Help1 = 6
Help2 = 7

End Enum


Private Type MenuItem
Caption As String
Task As eTask
Height As Single
Alive As Boolean
x As Single
y As Single
Mbuf As String
Mcount As Integer
Mc1 As Byte
MSDelay As Integer
mScrollDelay As Byte
Pt As Integer
End Type






Public Sub Initialise(ByVal State As eMenustate)
Menustate = State
Width = 450
Height = 200

x = 100
y = -Height

Mspeed = 10
SMno = 1
For i = 1 To 10
MItem(i).mScrollDelay = 20

Next


End Sub
Public Sub Terminate()



End Sub
Public Sub Update()

Call Checkevents
Call Move


If Menustate = ComeOut Then
Call CheckMenu
End If

If Not Menustate = GoneIn Then
Call Render

End If

End Sub
Private Sub Move()

If Menustate = CominOut Then


y = y + Mspeed
If y >= 0 Then
y = 0
Menustate = ComeOut
Sound.StopSound 1
Call IncSMno

End If
End If


If Menustate = GoingIn Then


y = y - Mspeed
If y <= (-Height) Then
y = -Height
Menustate = GoneIn

End If
End If




End Sub
Private Sub Checkevents()

If (Gamestate = paused Or Gamestate = Stopped) And Menustate = GoneIn Then
Menustate = CominOut
Frmpics.Mbuf.Width = Canvas_width * Screen.TwipsPerPixelX
Frmpics.Mbuf.Height = Canvas_height * Screen.TwipsPerPixelY
BitBlt Frmpics.Mbuf.hdc, 0, 0, Canvas_width, Canvas_height, dhoom1.Canvas.hdc, 0, 0, vbSrcCopy

End If

If GetKeystate(vbKeyEscape) = True And (Menustate = ComeOut) And Gamestate = paused Then
Menustate = GoingIn
Gamestate = running
Frmpics.Mbuf.Picture = LoadPicture()
Call ClearMenu
End If




End Sub
Private Sub Render()

Frmpics.Buffer.FillColor = vbBlack


 Frmpics.Mbuf.FillStyle = 0
 Frmpics.Buffer.ForeColor = vbWhite
If Not Menustate = GoingIn Then




 BitBlt Frmpics.Buffer.hdc, 0, 0, Canvas_width, Canvas_height, Frmpics.Mbuf.hdc, 0, 0, vbSrcCopy
Rectangle Frmpics.Buffer.hdc, x, y, x + Width, y + Height

Else
Rectangle Frmpics.Buffer.hdc, x, y, x + Width, y + Height


End If

If Menustate = ComeOut Then
Call RenderItems
End If




End Sub
Public Sub AddMitem(ByVal cCaption As String, ByVal cTask As eTask, ByVal cHeight As Single, ByVal cX As Single, Optional ByVal Delay As Integer, Optional ByVal MScDelay As Byte, Optional ByVal cY As Single)


For i = 1 To 10

If MItem(i).Alive = False Then

With MItem(i)
.Pt = 0
.Mcount = 0
.Caption = cCaption
.Task = cTask
.Height = cHeight
.Alive = True
.x = cX + x
.y = cY
.Mbuf = .Caption
If Delay <> 0 Then
.MSDelay = Delay
If MScDelay > 0 Then
.mScrollDelay = MScDelay
End If

End If

End With

Exit For
End If


Next



End Sub
Public Sub ClearMenu()

For i = 1 To 10

MItem(i).Alive = False

Next
SMno = 0


End Sub
Private Sub RenderItems()


For i = 1 To 10
If MItem(i).Alive = True Then

If SMno = i Then
Frmpics.Buffer.ForeColor = vbRed
Else
Frmpics.Buffer.ForeColor = vbWhite
End If



Frmpics.Buffer.CurrentX = MItem(i).x * Screen.TwipsPerPixelX
Frmpics.Buffer.CurrentY = (MItem(i).y + 30) * Screen.TwipsPerPixelY
Frmpics.Buffer.Print MItem(i).Caption

Dim Ti As Byte
Ti = i + 1
If Ti <= 10 Then

If MItem(Ti).y = 0 Then
MItem(Ti).y = MItem(i).y + MItem(i).Height
End If
End If


End If
Next


End Sub
Private Sub CheckMenu()




If GetKeyPress(vbKeyDown) = True Then
Call IncSMno
End If
If GetKeyPress(vbKeyUp) = True Then
Call DecSMno
End If

If GetKeystate(vbKeyReturn) = True Then

If SMno > 0 And SMno <= 10 Then
Dowork MItem(SMno).Task
End If

End If

For i = 1 To 10
If MItem(i).Task = Message Then
With MItem(i)

.Caption = Left(.Caption, 63)

If .Mcount < .MSDelay Then
.Mcount = .Mcount + 1
Else
.Mc1 = .Mc1 + 1
If .Mc1 > .mScrollDelay Then
.Mc1 = 0
.Pt = .Pt + 1
.Caption = Mid(.Mbuf, .Pt + 1, 63)
If Len(.Caption) = 0 Then
.Caption = Left(.Mbuf, 63)
.Mcount = 0
.Pt = 0
End If

End If


End If





End With
End If


Next































End Sub
Public Sub IncSMno()
Dim C As Byte

st:
SMno = SMno + 1
If SMno > 10 Then SMno = 1

If MItem(SMno).Alive = False Or MItem(SMno).Task = Message Then
C = C + 1
If C <= 10 Then GoTo st



End If



End Sub
Public Sub DecSMno()
Dim C As Byte

st:
SMno = SMno - 1
If SMno < 1 Then SMno = 10

If MItem(SMno).Alive = False Or MItem(SMno).Task = Message Then
C = C + 1
If C <= 10 Then GoTo st



End If



End Sub
Public Sub Dowork(ByVal tTask As eTask)


If tTask = Ok Then
Menustate = GoingIn
Gamestate = running
Frmpics.Mbuf.Picture = LoadPicture()
Sound.PlaySound 1, 0, False
Call ClearMenu
End If

If tTask = NewGame Then
Menustate = GoingIn
Gamestate = running
Frmpics.Mbuf.Picture = LoadPicture()
Call ClearMenu
Player.clearlevel
Player.LoadLevel 1
End If


If tTask = Quit Then
Unload dhoom1
End If

If tTask = Mainmenu Then
Menu.ClearMenu
Menu.AddMitem "NEWGAME", NewGame, 30, 30
Menu.AddMitem "HELP ABOUT THE GAME", Help1, 30, 30
Menu.AddMitem "HELP ABOUT THE CODE", Help2, 30, 30
Menu.AddMitem "QUIT GAME", Quit, 30, 30

Menu.Menustate = GoingIn
Gamestate = paused


End If

If tTask = Help1 Then

If Dir(App.path & "\help\HELP_normal.html") <> "" Then
Call ShellExecute(dhoom1.hwnd, "Open", App.path & "\help\HELP_normal.html", "", App.path, 1)
End If


End If

If tTask = Help2 Then

If Dir(App.path & "\help\HELP_code.html") <> "" Then
Call ShellExecute(dhoom1.hwnd, "Open", App.path & "\help\HELP_code.html", "", App.path, 1)
End If


End If
End Sub
