VERSION 5.00
Begin VB.UserControl DISPLAY 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF0000&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FillColor       =   &H000000C0&
   FillStyle       =   0  'Solid
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "display.ctx":0000
   Begin VB.Timer Tmrmsg 
      Left            =   360
      Top             =   2640
   End
   Begin VB.ListBox Lstmsgs 
      Height          =   840
      ItemData        =   "display.ctx":0312
      Left            =   1440
      List            =   "display.ctx":0314
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Lblmsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TANMAY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "DISPLAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim message() As String
Dim step As Integer
Dim curmessageno, nextmessageno As Integer
Dim noofmsgs As Integer
Dim run As Boolean
Dim char As Byte

'Default Property Values:
Const m_def_Interval = 200
Const m_def_cursorchar = "#"
Const m_def_tag = "MADE BY TANMAY"
'Property Variables:
Dim m_Interval As Integer
Dim m_cursorchar As String
Dim m_tag As String
'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."




Private Sub Tmrmsg_Timer()
If run = True Then
DisplayMsg
End If

End Sub

Public Sub start()

noofmsgs = Lstmsgs.ListCount
Lblmsg.Caption = ""

ReDim message(noofmsgs)
If noofmsgs > 0 Then
run = True
Tmrmsg.Interval = m_Interval
curmessageno = 0

nextmessageno = 1
For i = 0 To noofmsgs
   message(i) = Lstmsgs.List(i)


   
   
Next i



End If



End Sub
Private Sub DisplayMsg()
On Error Resume Next
Static msg As String
Static lefton As String, righton As String
lefton = step - 1
If Len(message(curmessageno)) >= step Then
righton = Len(message(curmessageno)) - step
Else: righton = 0
End If
If righton > 0 Then
If lefton > 0 Then
msg = Left(message(nextmessageno), lefton)
If step > Len(message(nextmessageno)) - 1 Then
For i = 0 To step - (Len(message(nextmessageno))) - 1
msg = msg + " "
Next
End If
msg = msg + m_cursorchar + Right(message(curmessageno), righton)
Lblmsg.Caption = msg
Else: Lblmsg.Caption = message(curmessageno)
End If
Else
If step < (Len(message(nextmessageno)) + 1) Then
Lblmsg.Caption = Left(message(nextmessageno), lefton) + m_cursorchar
Else: Lblmsg.Caption = Left(message(nextmessageno), lefton)
End If
End If
step = step + 1
If lefton > Len(message(nextmessageno)) And righton = 0 Then
curmessageno = curmessageno + 1
nextmessageno = nextmessageno + 1
If curmessageno > noofmsgs Then curmessageno = 0
If nextmessageno > noofmsgs Then nextmessageno = 0
step = 1
End If
End Sub

Private Sub UserControl_Resize()
Lblmsg.Height = UserControl.Height
Lblmsg.Width = UserControl.Width

End Sub
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Lblmsg,Lblmsg,-1,BackColor
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = Lblmsg.BackColor
'End Property
'
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    Lblmsg.BackColor() = New_BackColor
'    PropertyChanged "BackColor"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Lblmsg,Lblmsg,-1,BackStyle
'Public Property Get BackStyle() As Integer
'    BackStyle = Lblmsg.BackStyle
'End Property
'
'Public Property Let BackStyle(ByVal New_BackStyle As Integer)
'    Lblmsg.BackStyle() = New_BackStyle
'    PropertyChanged "BackStyle"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Lblmsg,Lblmsg,-1,BorderStyle
'Public Property Get BorderStyle() As Integer
'    BorderStyle = Lblmsg.BorderStyle
'End Property
'
'Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
'    Lblmsg.BorderStyle() = New_BorderStyle
'    PropertyChanged "BorderStyle"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Lblmsg,Lblmsg,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Lblmsg.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Lblmsg.Font = New_Font
    PropertyChanged "Font"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Lblmsg,Lblmsg,-1,ForeColor
'Public Property Get ForeColor() As OLE_COLOR
'    ForeColor = Lblmsg.ForeColor
'End Property
'
'Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
'    Lblmsg.ForeColor() = New_ForeColor
'    PropertyChanged "ForeColor"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Tmrmsg,Tmrmsg,-1,Interval
'Public Property Get Interval() As Long
'    Interval = Tmrmsg.Interval
'End Property
'
'Public Property Let Interval(ByVal New_Interval As Long)
'    Tmrmsg.Interval() = New_Interval
'    PropertyChanged "Interval"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Lstmsgs,Lstmsgs,-1,List
Public Property Get List(ByVal Index As Integer) As String
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
    List = Lstmsgs.List(Index)
End Property

Public Property Let List(ByVal Index As Integer, ByVal New_List As String)
    Lstmsgs.List(Index) = New_List
    PropertyChanged "List"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Lblmsg,Lblmsg,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = Lblmsg.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set Lblmsg.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14

Public Sub Halt()
run = False
Tmrmsg.Interval = 0
Lblmsg.Caption = ""
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,#
Public Property Get cursorchar() As String
Attribute cursorchar.VB_Description = "this is the char that acts as the corsor. ""#"" llooks kool! ---TANMAY"
    cursorchar = m_cursorchar
End Property

Public Property Let cursorchar(ByVal New_cursorchar As String)
    m_cursorchar = New_cursorchar
    PropertyChanged "cursorchar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,1,MADE BY TANMAY
Public Property Get tag() As String
Attribute tag.VB_Description = "ANYTHIN STRING"
    tag = m_tag
End Property

Public Property Let tag(ByVal New_tag As String)
    If Ambient.UserMode = False Then Err.Raise 387
    m_tag = New_tag
    PropertyChanged "tag"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_cursorchar = m_def_cursorchar
    m_tag = m_def_tag
    m_Interval = m_def_Interval
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer

'    Lblmsg.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
'    Lblmsg.BackStyle = PropBag.ReadProperty("BackStyle", 1)
'    Lblmsg.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Set Lblmsg.Font = PropBag.ReadProperty("Font", Ambient.Font)
'    Lblmsg.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
'    Tmrmsg.Interval = PropBag.ReadProperty("Interval", 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Lstmsgs.List(Index) = PropBag.ReadProperty("List" & Index, "")
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_cursorchar = PropBag.ReadProperty("cursorchar", m_def_cursorchar)
    m_tag = PropBag.ReadProperty("tag", m_def_tag)
'    Lblmsg.BackStyle = PropBag.ReadProperty("back", 0)
'    Lblmsg.BackColor = PropBag.ReadProperty("ForeColor", &H8000000F)
'    Lblmsg.ForeColor = PropBag.ReadProperty("backv", &H80000012)
    Lblmsg.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
'    Lblmsg.BackStyle = PropBag.ReadProperty("backv", 0)
    m_Interval = PropBag.ReadProperty("Interval", m_def_Interval)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer

'    Call PropBag.WriteProperty("BackColor", Lblmsg.BackColor, &H8000000F)
'    Call PropBag.WriteProperty("BackStyle", Lblmsg.BackStyle, 1)
'    Call PropBag.WriteProperty("BorderStyle", Lblmsg.BorderStyle, 0)
    Call PropBag.WriteProperty("Font", Lblmsg.Font, Ambient.Font)
'    Call PropBag.WriteProperty("ForeColor", Lblmsg.ForeColor, &H80000012)
'    Call PropBag.WriteProperty("Interval", Tmrmsg.Interval, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("List" & Index, Lstmsgs.List(Index), "")
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("cursorchar", m_cursorchar, m_def_cursorchar)
    Call PropBag.WriteProperty("tag", m_tag, m_def_tag)
'    Call PropBag.WriteProperty("back", Lblmsg.BackStyle, 0)
'    Call PropBag.WriteProperty("ForeColor", Lblmsg.BackColor, &H8000000F)
'    Call PropBag.WriteProperty("backv", Lblmsg.ForeColor, &H80000012)
    Call PropBag.WriteProperty("ForeColor", Lblmsg.ForeColor, &H80000008)
'    Call PropBag.WriteProperty("backv", Lblmsg.BackStyle, 0)
    Call PropBag.WriteProperty("Interval", m_Interval, m_def_Interval)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Lstmsgs,Lstmsgs,-1,AddItem
Public Sub AddItem(ByVal Item As String)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
    Lstmsgs.AddItem Item
    Call start
    
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Lstmsgs,Lstmsgs,-1,RemoveItem
Public Sub RemoveItem(ByVal Index As Integer)
Attribute RemoveItem.VB_Description = "Removes an item from a ListBox or ComboBox control or a row from a Grid control."
    Lstmsgs.RemoveItem Index
    Call start
    
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,BackColor
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = UserControl.BackColor
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,BackStyle
'Public Property Get BackStyle() As Integer
'    BackStyle = UserControl.BackStyle
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Lblmsg,Lblmsg,-1,BackStyle
'Public Property Get back() As Integer
'    back = Lblmsg.BackStyle
'End Property
'
'Public Property Let back(ByVal New_back As Integer)
'    Lblmsg.BackStyle() = New_back
'    PropertyChanged "back"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Lblmsg,Lblmsg,-1,BackColor
'Public Property Get ForeColor() As OLE_COLOR
'    ForeColor = Lblmsg.BackColor
'End Property
'
'Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
'    Lblmsg.BackColor() = New_ForeColor
'    PropertyChanged "ForeColor"
'End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=Lblmsg,Lblmsg,-1,BackColor
''Public Property Get BackColor() As OLE_COLOR
''    BackColor = Lblmsg.BackColor
''End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=Lblmsg,Lblmsg,-1,BackStyle
''Public Property Get BackStyle() As Integer
''    BackStyle = Lblmsg.BackStyle
''End Property
''
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,BackColor
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = UserControl.BackColor
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property
Public Property Let BackStyle(ByVal new_back As Integer)
 
If new_back >= 0 And new_back <= 1 Then

 UserControl.BackStyle = new_back
   PropertyChanged "backstyle"
   End If
   
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Lblmsg,Lblmsg,-1,ForeColor
'Public Property Get backv() As OLE_COLOR
'    backv = Lblmsg.ForeColor
'End Property
'
'Public Property Let backv(ByVal New_backv As OLE_COLOR)
'    Lblmsg.ForeColor() = New_backv
'    PropertyChanged "backv"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Lblmsg,Lblmsg,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Lblmsg.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Lblmsg.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Lblmsg,Lblmsg,-1,BackStyle
'Public Property Get backv() As Integer
'    backv = Lblmsg.BackStyle
'End Property
'
'Public Property Let backv(ByVal New_backv As Integer)
'If New_backv < 2 And New_backv >= 0 Then
'    Lblmsg.BackStyle() = New_backv
'    PropertyChanged "backv"
'
'End If
'
'End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,200
Public Property Get Interval() As Integer
Attribute Interval.VB_Description = "Returns/sets the number of milliseconds between calls to a Timer control's Timer event."
    Interval = m_Interval
End Property

Public Property Let Interval(ByVal New_Interval As Integer)
    m_Interval = New_Interval
    PropertyChanged "Interval"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Lstmsgs,Lstmsgs,-1,Clear
Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of a control or the system Clipboard."
    Lstmsgs.Clear
    Call Halt
    
End Sub

