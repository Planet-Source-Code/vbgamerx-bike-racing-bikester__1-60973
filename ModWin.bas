Attribute VB_Name = "ModWin"

'  This module Contains subs that care called when the player either loses or completes a level



'*********************************************************************************
'**********************************************************************************

'This sub is called when the player wins the level
' Takes argument on how the player has won. Currently there is only one way- by increasing the
'dist bet. player and the police.Still i kept the option open for further development.

Public Sub Win(ByVal Num As wintype)   ' Num defines how player has won



' Add a message to the menu object. Some text and the level no is added
' On more Information on menu See the clsMenu Class
Menu.AddMitem "GOOD WORK,  LEVEL " & Str(Player.Level) & " CLEARED! TRY THE NEXT ONE ", Message, 40, 20, 300


' Add a Ok button to the menu
Menu.AddMitem "OK", Ok, 30, 200
Gamestate = paused ' Pause the game



  Player.LoadLevel Player.Level + 1     ' Load the next Level


End Sub
'***********************************************************************************
'**********************************************************************************
' This sub is called when the player Loses. Takes arguments on How the player Lost
' Player can lose When police catches him or health becomes zero
Public Sub Lose(ByVal Num As Losetype)

      '****
       If Num = Bypolice Then              ' If Police caught the Player
               
            ' Add a message and a Ok button to the menu
            Menu.AddMitem "Sorry, The police has caught you! You will land In JAIL! Try Again", Message, 40, 20, 300
            Menu.AddMitem "EXIT TO MAINMENU", Mainmenu, 30, 150, , , 60


            Gamestate = paused 'pause game
            Player.clearlevel       ' Clear current level




     End If
    '****




         '****
         If Num = Bycrash Then       ' If player's health becomes 0 then
            '
            'Add message and a OK button to the menu
            Menu.AddMitem "Dont you know how to drive ! Try Again", Message, 40, 20, 200
            Menu.AddMitem "EXIT TO MAINMENU", Mainmenu, 30, 150, , , 60
            
           Gamestate = paused          ' Pause game
           Player.clearlevel                 ' Clear current level

        End If
        '*****




End Sub
'************************************************************************************
'***********************************************************************************

