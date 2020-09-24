VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ledit 
   BackColor       =   &H0080C0FF&
   Caption         =   "EDITOR"
   ClientHeight    =   7035
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10425
   Icon            =   "ledit.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "ledit.frx":030A
   ScaleHeight     =   7035
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txtmspeed 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Height          =   405
      Left            =   8400
      TabIndex        =   33
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Cmdcmspeed 
      BackColor       =   &H00FF8080&
      Caption         =   "Calculate Max Speed"
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Cmdopen 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "OPEN"
      Height          =   195
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox TxtPspeed 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   30
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Cmdremovelevel 
      BackColor       =   &H000080FF&
      Caption         =   "REMOVE LEVEL"
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Cmdnewlevel 
      BackColor       =   &H000080FF&
      Caption         =   "NEW LEVEL"
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton CmdVupdate 
      BackColor       =   &H00008000&
      Caption         =   "UPDATE VALUES"
      Height          =   375
      Left            =   8400
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton CmdLupdate 
      BackColor       =   &H000080FF&
      Caption         =   "UPDATE VALUES"
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton CmdvRemove 
      BackColor       =   &H00008000&
      Caption         =   "REMOVE Vehicle"
      Height          =   375
      Left            =   8400
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton CmdVadd 
      BackColor       =   &H00008000&
      Caption         =   "ADD Vehicle"
      Height          =   375
      Left            =   8400
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.ComboBox CmbVtype 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "ledit.frx":0614
      Left            =   6600
      List            =   "ledit.frx":061E
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Txtrand 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   6600
      TabIndex        =   18
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox Txttimeout 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6600
      TabIndex        =   17
      Top             =   5880
      Width           =   1335
   End
   Begin VB.ListBox Lstvehicles 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      IntegralHeight  =   0   'False
      Left            =   2160
      TabIndex        =   15
      Top             =   5520
      Width           =   2775
   End
   Begin VB.TextBox Txtvdamage 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Txtfriction 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   13
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Txtaccl 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Txtturnspeed 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Txtwindist 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Txtdist 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox TxtMsg 
      BackColor       =   &H000080FF&
      Height          =   735
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   480
      Width           =   5775
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   840
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox Lstlevels 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Police Speed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4440
      TabIndex        =   29
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Ranmdom value"
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   5040
      TabIndex        =   21
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Time out"
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   5040
      TabIndex        =   20
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TYPE"
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   5040
      TabIndex        =   19
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicles Present"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Damage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Road Friction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Aceleration Force"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Turning Speed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Winning Distance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Initial distance between Police and Player"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MESSAGE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Menu mnuile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuGethelp 
         Caption         =   "Get help"
      End
   End
End
Attribute VB_Name = "ledit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Fini As CIniFile
Attribute Fini.VB_VarHelpID = -1
Dim Fname As String
Dim Lno As Byte
Dim Vno As Byte

Private Sub Cmdcmspeed_Click()
Dim t1 As Single, t2 As Single
Dim t3 As Single
Dim tSpeed As Single


t1 = Val(Txtaccl.Text)
t2 = Val(Txtfriction.Text)




Do
tSpeed = tSpeed + t1
t3 = tSpeed * t2
tSpeed = tSpeed - t3
DoEvents
Loop Until (t1 - t3) < 0.0001

Txtmspeed.Text = tSpeed



End Sub

Private Sub CmdLupdate_Click()
On Error GoTo hell

Dim temp As String

If Fname = "" Then Exit Sub
If Lno = 0 Then Exit Sub

                temp = TxtMsg.Text
                 Fini.EntryWrite "message", temp, "level " & Trim(Str(Lno)), Fname
                
                temp = Txtdist.Text
                Fini.EntryWrite "dist", temp, "level " & Trim(Str(Lno)), Fname
                
                temp = Txtwindist.Text
                Fini.EntryWrite "windist", temp, "level " & Trim(Str(Lno)), FileName
                 
                temp = Txtturnspeed.Text
                Fini.EntryWrite "turnspeed", temp, "level " & Trim(Str(Lno)), FileName
                
                 ValTemp = Txtaccl.Text
                Fini.EntryWrite "accl", temp, "level " & Trim(Str(Lno)), FileName
                
                temp = Txtvdamage.Text
                Fini.EntryWrite "vdamage", temp, "level " & Trim(Str(Lno)), Fname
                
                temp = Txtfriction.Text
                 Fini.EntryWrite "friction", temp, "level " & Trim(Str(Lno)), Fname
                 
                 temp = TxtPspeed.Text
                 Fini.EntryWrite "pspeed", temp, "level " & Trim(Str(Lno)), Fname
Exit Sub

hell:

MsgBox Err.Description & "--Progrem will EXIT"
End

                
End Sub

Private Sub Cmdnewlevel_Click()
Dim Ty As Byte
On Error GoTo hell

Ty = Lstlevels.ListCount + 1

If Fname = "" Then Exit Sub

               
                 Fini.EntryWrite "message", "HELLO", "level " & Trim(Str(Ty)), Fname
                
               
                Fini.EntryWrite "dist", "1000", "level " & Trim(Str(Ty)), Fname
                
                
                Fini.EntryWrite "windist", "2000", "level " & Trim(Str(Ty)), FileName
                 
                
                Fini.EntryWrite "turnspeed", "1.5", "level " & Trim(Str(Ty)), FileName
                
                 
                Fini.EntryWrite "accl", "1.5", "level " & Trim(Str(Ty)), FileName
                
                
                Fini.EntryWrite "vdamage", "2", "level " & Trim(Str(Ty)), Fname
                
                
                 Fini.EntryWrite "friction", "0.01", "level " & Trim(Str(Ty)), Fname
                 
                 
                 Fini.EntryWrite "pspeed", "100", "level " & Trim(Str(Ty)), Fname
                 
                 
Call Openlevel(Fname)
Call LoadLdata(Ty, Fname)
Exit Sub

hell:

MsgBox Err.Description & "--Progrem will EXIT"
End


End Sub

Private Sub Cmdopen_Click()
mnuOpen_Click
End Sub

Private Sub Cmdremovelevel_Click()
On Error GoTo hell

If Fname = "" Then Exit Sub



Fini.SectionDelete "level " & Trim(Str(Lstlevels.ListCount)), Fname

Call Openlevel(Fname)
LoadLdata Lstlevels.ListCount, Fname

Exit Sub

hell:

MsgBox Err.Description & "--Progrem will EXIT"
End



End Sub

Private Sub CmdVadd_Click()
Dim Tx As Byte
On Error GoTo hell

If Fname = "" Then Exit Sub
If Lno = 0 Then Exit Sub

Tx = Lstvehicles.ListCount + 1

If Tx <= 10 Then

Fini.EntryWrite "ntype", "1", "l" & Trim(Str(Lno)) & "vehicleadd" & Trim(Str(Tx)), Fname

Fini.EntryWrite "timeout", "100", "l" & Trim(Str(Lno)) & "vehicleadd" & Trim(Str(Tx)), Fname
 
 Fini.EntryWrite "rand", "200", "l" & Trim(Str(Lno)) & "vehicleadd" & Trim(Str(Tx)), Fname

End If

Vno = 0
Call LoadVehicles(Lno, Fname)
Lstvehicles.ListIndex = Lstvehicles.ListCount - 1

Exit Sub

hell:

MsgBox Err.Description & "--Progrem will EXIT"
End




End Sub

Private Sub CmdvRemove_Click()
Dim X1 As Byte
On Error GoTo hell


If Fname = "" Then Exit Sub
If Lno = 0 Then Exit Sub

X1 = Lstvehicles.ListCount

If X1 = 0 Then Exit Sub

Fini.SectionDelete "l" & Trim(Str(Lno)) & "vehicleadd" & Trim(Str(X1)), Fname

Vno = 0
Call LoadVehicles(Lno, Fname)
Lstvehicles.ListIndex = Lstvehicles.ListCount - 1
Exit Sub

hell:

MsgBox Err.Description & "--Progrem will EXIT"
End



End Sub

Private Sub CmdVupdate_Click()

On Error GoTo hell

Dim temp As String
If Fname = "" Then Exit Sub
If Lno = 0 Then Exit Sub
If Vno = 0 Then Exit Sub


Select Case CmbVtype.Text
   Case "car1"
       temp = 2
       Fini.EntryWrite "ntype", temp, "l" & Trim(Str(Lno)) & "vehicleadd" & Trim(Str(Vno)), Fname
   
   Case "truck1"
        temp = 1
       Fini.EntryWrite "ntype", temp, "l" & Trim(Str(Lno)) & "vehicleadd" & Trim(Str(Vno)), Fname
End Select

temp = Txttimeout.Text
Fini.EntryWrite "timeout", temp, "l" & Trim(Str(Lno)) & "vehicleadd" & Trim(Str(Vno)), Fname

temp = Txtrand.Text
Fini.EntryWrite "rand", temp, "l" & Trim(Str(Lno)) & "vehicleadd" & Trim(Str(Vno)), Fname

Exit Sub

hell:

MsgBox Err.Description & "--Progrem will EXIT"
End





End Sub



Private Sub Fini_EnumIniSection(ByVal SectionName As String, ByVal FileName As String, Cancel As Boolean)
End
End Sub

Private Sub Form_Click()
'Print Lstvehicles.ListCount
End Sub

Private Sub Form_Load()

Set Fini = New CIniFile


End Sub

Private Sub Lstlevels_Click()



Lno = Right(Lstlevels.List(Lstlevels.ListIndex), 1)
Vno = 0
Call LoadLdata(Lno, Fname)
Call LoadVehicles(Lno, Fname)






End Sub

Private Sub Lstvehicles_Click()


Vno = Right(Lstvehicles.List(Lstvehicles.ListIndex), 1)

Call LoadVdata(Vno, Lno, Fname)



End Sub

Private Sub mnuabout_Click()
frmAbout.Show 1

End Sub

Private Sub mnuOpen_Click()

On Error GoTo hell

Lstlevels.Clear
Lstvehicles.Clear



CD1.InitDir = makepath(App.path)
CD1.Filter = "BikesterLevels|*.blv"
CD1.CancelError = True
CD1.ShowOpen
Openlevel CD1.FileName
Fname = CD1.FileName



Exit Sub
hell:
Fname = ""


End Sub
Private Sub Openlevel(ByVal FileName As String)
Dim T As String

On Error GoTo hell

Lstlevels.Clear
For i = 1 To 20

 
 
 T = Fini.EntryRead("dist", "notfound", "Level " & Trim(Str(i)), FileName)
 
 If T <> "notfound" Then
 Lstlevels.AddItem "level" & Trim(Str(i))
 Else
 Exit For
 End If
'Fini.Filename = Filename
'Fini.EnumSections Filename
Next
Exit Sub

hell:

MsgBox Err.Description & "--Progrem will EXIT"
End


End Sub
Private Sub LoadLdata(ByVal LevelNo As Byte, ByVal FileName As String)
Dim temp As String

On Error GoTo hell

If Fname = "" Then Exit Sub


                temp = Fini.EntryRead("message", "N/A", "level " & Trim(Str(LevelNo)), FileName)
                TxtMsg.Text = temp
                
                temp = Fini.EntryRead("dist", "N/A", "level " & Trim(Str(LevelNo)), FileName)
                Txtdist.Text = Val(temp)

                temp = Fini.EntryRead("windist", "N/A", "level " & Trim(Str(LevelNo)), FileName)
                Txtwindist.Text = Val(temp)
                
                temp = Fini.EntryRead("turnspeed", "N/A", "level " & Trim(Str(LevelNo)), FileName)
                Txtturnspeed.Text = Val(temp)
                
                temp = Fini.EntryRead("accl", "N/A", "level " & Trim(Str(LevelNo)), FileName)
                Txtaccl.Text = Val(temp)
                
                temp = Fini.EntryRead("vdamage", "N/A", "level " & Trim(Str(LevelNo)), FileName)
                Txtvdamage.Text = Val(temp)
                
                 temp = Fini.EntryRead("friction", "N/A", "level " & Trim(Str(LevelNo)), FileName)
                Txtfriction.Text = Val(temp)

                 temp = Fini.EntryRead("pspeed", "N/A", "level " & Trim(Str(LevelNo)), FileName)
                TxtPspeed.Text = Val(temp)
Lno = LevelNo
Exit Sub

hell:

MsgBox Err.Description & "--Progrem will EXIT"
End


End Sub
Private Sub LoadVehicles(ByVal LevelNo As Byte, ByVal FileName As String)
Dim temp As String

On Error GoTo hell

Lstvehicles.Clear

For i = 1 To 10
temp = Fini.EntryRead("ntype", "notfound", "l" & Trim(Str(LevelNo)) & "vehicleadd" & Trim(Str(i)), FileName)
If temp = "notfound" Then
Exit For
Else


Lstvehicles.AddItem "l" & Trim(Str(LevelNo)) & "vehicleadd" & Trim(Str(i))

End If
Next
Exit Sub

hell:

MsgBox Err.Description & "--Progrem will EXIT"
End

    

End Sub
Private Sub LoadVdata(ByVal VehicleNo As Byte, ByVal LevelNo As Byte, ByVal FileName As String)

On Error GoTo hell

temp = Fini.EntryRead("ntype", "notfound", "l" & Trim(Str(LevelNo)) & "vehicleadd" & Trim(Str(VehicleNo)), FileName)
If temp = "notfound" Then
Exit Sub
Else
Select Case Val(temp)

   Case 1
      CmbVtype.Text = "truck1"
      
  Case 2
    CmbVtype.Text = "car1"
    
End Select

temp = Fini.EntryRead("timeout", "notfound", "l" & Trim(Str(LevelNo)) & "vehicleadd" & Trim(Str(VehicleNo)), FileName)
Txttimeout.Text = temp
temp = Fini.EntryRead("rand", "notfound", "l" & Trim(Str(LevelNo)) & "vehicleadd" & Trim(Str(VehicleNo)), FileName)
Txtrand.Text = temp

End If

Exit Sub

hell:

MsgBox Err.Description & "--Progrem will EXIT"
End


End Sub
