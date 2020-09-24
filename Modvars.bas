Attribute VB_Name = "Modvars"
' AUTHOR - TANMAY DEHURY

'This module contains the declairations of the public variables used in the game,
' Instances of the classses, Public Constants and public Enums
'
'
'
'
'
'
'
'************************************************************************


Public Const Canvas_width = 805   '- Defines the width of the main canvas(size of game area)
Public Const Canvas_height = 595  '- defines the height of the main canvas(size of game area)




'*************************************************************************

Public T1 As Long         '- variable for limiting the frane rate.
Public T2 As Long         '- variable for limiting the frane rate.
' the gameloop proceeds only when the difference between t1 and t2 becomes greater than a certain value

Public FpsLimiter   As Single         '- this value limits the fps with the help of T1 and T2
Public Gamestate As Gstate  '- this  defines the state of the game. can be Paused- 2, stopped - 3, running -1
Public Picpath As String        '- stores the path where pictures are present
Public tFps As Integer               '- this is the fps buffer.This increases with every frame and after 1 sec
'its value becomes the fps and tfps is reset to 0
Public Fps As Integer    ' this stores the Frammes per second value. Updated every sec from tfps
Public LevelPath As String   ' This stores the Path of level data("Levels.tlv") file
Public MFps As Single         'this stores the Fps that is ideal
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



Public Player As New ClsPlayer  '- This is a instance of the clsPlayer class
Public Sound As New Clssound  '- This is a instance of the clsSound class
Public Road As New ClsRoad     '- This is a instance of the clsRoad class
Public Vehicle(1 To 10) As New ClsVehicle   '- 10 instances of the clsVehicle class
Public Display As New ClsDisplay   '- This is a instance of the clsDisplay class
Public Menu As New Clsmenu     '- This is a instance of the clsMenu class
Public LevelSys As New CIniFile        '- Instance of Cinifile class(Not written by me)
Public Trs As New ClsTrs
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


 Public Enum Gstate  '- this defines the gamestate
 ' Doing enum makes it easier to remember the game states
 
 running = 1  'Game is going on as usual
 paused = 2 ' Game is Paused
 Stopped = 3  ' Game is stopped
 End Enum
 
 
Public Enum Losetype   'Defines the ways in which the game can be lost

Bypolice = 1    ' When the police catches the player
Bycrash = 2     ' When player health becomes zero

End Enum

Public Enum wintype    'Defines the ways in which the game can be won

Outrun = 1       ' distance between player and the police exceeds a certain value


End Enum

