VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsRoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



Public x As Single
Public y As Single
Public Width As Single
Public Height As Single
Public Lane1_x As Single
Public Lane2_x As Single
Public Fic_dcl As Single










Public Sub Initialise()
x = 200
y = 0
Width = 250
Height = 400
Lane1_x = 75
Lane2_x = 160
Fic_dcl = 0.01
End Sub
Public Sub Terminate()


End Sub
Public Sub Update()

Call Commonwork
Call Render


End Sub
Private Sub Commonwork()





End Sub
Public Function Getfic(ByVal Speed As Single) As Single

Getfic = Speed * Fic_dcl



End Function
Private Sub Render()
Static Counter As Single



Frmpics.Buffer.FillColor = RGB(50, 50, 50)
Frmpics.Buffer.FillStyle = 0
Rectangle Frmpics.Buffer.hdc, x, y, x + Width, y + Height


If Counter > 35 Then
Counter = 0
Else
Counter = Counter + Player.Speed / 10

End If




For i = y To (y + Height - 50) Step 35

Frmpics.Buffer.FillColor = vbWhite
Frmpics.Buffer.DrawWidth = 1


Rectangle Frmpics.Buffer.hdc, x + Lane1_x, i + Counter, x + Lane1_x + 5, i + 21 + Counter
Rectangle Frmpics.Buffer.hdc, x + Lane2_x, i + Counter, x + Lane2_x + 5, i + 21 + Counter


Next i






End Sub
