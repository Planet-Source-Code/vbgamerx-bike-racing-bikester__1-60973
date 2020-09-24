Attribute VB_Name = "Modio"
'----- LINES -----
'(used in "GeDrawLine" function)
    Private Declare Function MoveToEx Lib "gdi32" _
    (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
    lpPoint As Any) As Long
        Private Declare Function LineTo Lib "gdi32" _
        (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) _
        As Long
   
Public Const PI = 3.14285
















Public Function GetKeystate(ByVal Keyno As Integer) As Boolean

If GetAsyncKeyState(Keyno) < 0 Then
GetKeystate = True
Else
GetKeystate = False
End If



End Function
Public Sub CheckKeys()

If GetKeyPress(vbKeyEscape) = True Then


               Select Case Gamestate

               Case running
               
               Menu.AddMitem "RESUME", Ok, 30, 30
               Menu.AddMitem "MAIN MENU", Mainmenu, 30, 30
               Menu.AddMitem "RESTART", NewGame, 30, 30
               Menu.AddMitem "EXIT GAME", Quit, 30, 30
               
               
               Gamestate = paused
               dhoom1.Label2.Caption = "PAUSED"

             Case paused
            
              'Gamestate = running
              dhoom1.Label2.Caption = ""
            End Select

End If

If GetKeyPress(vbKeyQ) Then
Player.LoadLevel Player.Level + 1
End If
If GetKeyPress(vbKeyW) Then
Player.LoadLevel Player.Level - 1
End If


If GetKeyPress(vbKeyReturn) = True Then
    
    
    If Gamestate = paused Then Gamestate = Stopped

End If








End Sub
Public Function GetKeyPress(ByVal Keyno As Integer) As Boolean

If GetAsyncKeyState(Keyno) = -32767 Then
GetKeyPress = True
Else
GetKeyPress = False
End If



End Function
Public Sub msgbox1(ByVal msg As String)
msgg = msg
Frmmsgbox.Show (1)
'Frmmsgbox.Label1.Caption = msg




End Sub
Public Sub GeDrawLine(SurfDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
'draw streight line on a Device Context
 MoveToEx SurfDC, X1, Y1, ByVal 0&
 LineTo SurfDC, X2, Y2
End Sub
Public Function angtorad(ByVal ang As Single) As Single

angtorad = (PI / 180) * ang




End Function

