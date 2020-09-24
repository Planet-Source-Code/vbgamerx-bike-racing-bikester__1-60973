VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Most code, except for INI and some Midi Functions
' ©2005 Michael Bentley (ikillkenny@comcast.net)
' Feel free to use code from this project, but it would nice
' if you gave me a little acknowledgement in the credits
' E-mail me any questions, and if you like the code please vote
' for me at www.planetsourcecode.com


Public ENABLE_DEMO As Boolean 'Whether the demo has been enabled
Public ENABLE_MIDI As Boolean 'Whether or not to play music files
Public ENABLE_WAV As Boolean 'Whehter or not to play sounds
Public Soundpath As String

Public Buf As Long
'Dim DX As New DirectX7
'Dim Ds As DirectSound







Public curMidi As String 'Current midi being played


Public anType As Long





Private curSound As Long 'Index of last sound to be played (from 1-10)




'Wave Functions
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Midi functions
Private Declare Function mciSendString Lib "WINMM.DLL" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndcallback As Long) As Long
Private Declare Function mciGetErrorString Lib "WINMM.DLL" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function mciSetRepeat Lib "WINMM.DLL" Alias "mciSetRepeatA" ()
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

'Ini Functions
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Enum SND_FLAGS
    SND_ASYNC = &H1 '//lets you play a new wav sound, interrupting another
    SND_LOOP = &H8 '//loops the wav sound
    SND_NODEFAULT = &H2 '//if wav file not there, then make sure NOTHING plays
    SND_SYNC = &H0 '//no control to program til wav is done playing
    SND_NOSTOP = &H10 '//if a wav file is already playing then it wont interrupt
End Enum







'Get values back from an INI file
Function GetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    GetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function
'Write values to an INI file
Function WriteIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    WriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function

Sub PlySound(Strsound As String)
'Plays a wave file
If ENABLE_WAV = False Then Exit Sub 'Sounds not enabled, so just exit
curSound = 1
'Get the short (DOS friendly) location of the wave file, since mci needs it
Strsound = Soundpath & Strsound & ".wav"
If Strsound = "" Then Exit Sub 'No path found, exit

t = mciSendString("close wav" & CStr(curSound), 0&, 0, 0)  'Gets rid of the previous file existing on this alias
t = mciSendString("open " & Strsound & " Alias wav" & CStr(curSound), 0&, 0, 0) 'Open the new wav file
t = mciSendString("play wav" & CStr(curSound), 0&, 0, 0) 'Play the wave file

'curSound = curSound + 1 'This is the number of the sound to play
'If curSound > 10 Then 'Loop back to the 0th sound
  '  curSound = 0
'End If

End Sub
Sub PlayMidi(strMidi As String)
'Plays background music files
If ENABLE_MIDI = False Then Exit Sub 'Music not enabled, exit
strMidi = Soundpath & strMidi 'Get the short location of the midi file
If strMidi = "" Then Exit Sub 'No midi exists, exit
Call mciSendString("close all", 0&, 0, 0) 'Stop all previous music
Call mciSendString("open " & strMidi$ & " Alias midi", 0&, 0, 0) 'Open the new file
Call mciSendString("play midi", 0&, 0, 0) 'Play the new file

End Sub
Sub StopMidi(strMidi As String)
On Error Resume Next
'Below no longer used
'strMidi = GetShortPath(App.Path) & "\" & strMidi
'If strMidi = "" Then Exit Sub
'Close and stop the current midi being played
Call mciSendString("stop midi", 0&, 0, 0)
Call mciSendString("close all", 0&, 0, 0)
End Sub
Public Function GetShortPath(strFileName As String) As String
    Dim lngRes As Long, strPath As String
    'Create a buffer
    strPath = String$(165, 0)
    'retrieve the short pathname
    lngRes = GetShortPathName(strFileName, strPath, 164)
    'remove all unnecessary chr$(0)'s
    GetShortPath = Left$(strPath, lngRes)
End Function
Public Sub PauseMidi(strMidi As String)
On Error Resume Next
If ENABLE_MIDI = False Then Exit Sub 'Music not enabled, exit
Dim intReturn As Long
strMidi = Soundpath & strMidi 'Get the short path of the midi
If strMidi = "" Then Exit Sub 'No midi exists
intReturn = mciSendString("Pause midi", 0&, 0, 0) 'Send the pause command
If intReturn <> 0 Then 'Error pausing
    Debug.Print "pause error"
End If
End Sub
Public Sub ResumeMidi(strMidi As String)
'Resumes a midi from the position when it was paused
On Error Resume Next
If ENABLE_MIDI = False Then Exit Sub 'Music not enabled, exit
Dim dwReturn As Long
Dim pos As String * 128
strMidi = Soundpath & strMidi 'Get short path of midi
If strMidi = "" Then Exit Sub 'No midi exists

dwReturn = mciSendString("status midi position", pos, 128, 0&) 'Get the midi's position
Call mciSendString("play midi from " & pos, 0&, 0&, 0&) 'Play the midi from the position you got

End Sub
Public Sub RepeatMidi()
'Repeats a midi file
On Error Resume Next
If ENABLE_MIDI = False Then Exit Sub 'Music not enabled, exit
Dim dwReturn As Long
Dim Total As String * 128

dwReturn = mciSendString("set midi time format frames", Total, 128, 0&) 'Set the midi to the proper format
dwReturn = mciSendString("status midi length", Total, 128, 0&) 'Get the total length of the midi

Dim pos As String * 128

dwReturn = mciSendString("status midi position", pos, 128, 0&) 'Get the current position of the midi

If pos = Total Then 'If the position is that the end, repeat
    dwReturn = mciSendString("seek midi to 0", 0&, 0&, 0&) 'Seek the midi to 0 position (the start)
    mciSendString "Play midi", 0&, 0&, 0& 'Play the midi again
End If

End Sub



'Public Sub Initialise()

'Set Ds = DX.DirectSoundCreate("")


'End Sub
Public Function PlayWave(Strsound As String, Flags As SND_FLAGS)
Strsound = Soundpath & Strsound & ".wav"
sndPlaySound Strsound, Flags
End Function
'//stops a wave sound
Public Function StopWave()
sndPlaySound "", SND_NODEFAULT Or SND_ASYNC
End Function

Public Sub Terminate()
StopWave
End Sub
