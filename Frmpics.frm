VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Frmpics 
   Caption         =   "Form1"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   12345
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Mbuf 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   11
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox Pcarmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Pcar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox Smetermask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   4080
      Width           =   615
   End
   Begin VB.PictureBox Smeter 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox car1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   2880
      Width           =   495
   End
   Begin VB.PictureBox car1mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   2280
      Width           =   495
   End
   Begin VB.PictureBox Truck1mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   4320
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
   Begin VB.PictureBox Truck1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   4320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Bike1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   4320
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   2
      Top             =   720
      Width           =   180
   End
   Begin VB.PictureBox bike1mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   4320
      ScaleHeight     =   240
      ScaleWidth      =   120
      TabIndex        =   1
      Top             =   360
      Width           =   180
   End
   Begin VB.PictureBox Buffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   8295
      Left            =   360
      ScaleHeight     =   8265
      ScaleWidth      =   10785
      TabIndex        =   0
      Top             =   240
      Width           =   10815
   End
   Begin MediaPlayerCtl.MediaPlayer SStream 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   3240
      Width           =   615
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   -1  'True
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Frmpics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
