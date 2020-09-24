VERSION 5.00
Begin VB.Form dhoom1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3480
      Top             =   0
   End
   Begin VB.PictureBox Canvas 
      AutoSize        =   -1  'True
      Height          =   7260
      Left            =   0
      ScaleHeight     =   7200
      ScaleWidth      =   9600
      TabIndex        =   0
      Top             =   240
      Width           =   9660
   End
   Begin VB.Image Imgexit 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9360
      Top             =   0
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   9600
      X2              =   9360
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   9360
      X2              =   9600
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "dhoom1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
DoEvents







Gamestate = running
Call Init
Call Gameloop
End Sub

Private Sub Form_Load()

Me.Canvas.Picture = LoadPicture(App.path & "\loadingtext.jpg")
Me.Show


End Sub

Private Sub Form_Unload(Cancel As Integer)
'Sound.Terminate

Call RemoveFiles
End
End Sub


Private Sub Imgexit_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()

Fps = tFps

tFps = 0
Label1.Caption = "Fps-" & Fps

FpsLimiter = FpsLimiter + (Fps - MFps) / 10
If FpsLimiter < 0 Then FpsLimiter = 0

End Sub
