VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTrs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public currentres As String
Public currentfile As String
Public temptrsname As String
Dim Buf As String


Public Function newtrs(ByVal filename As String) As Boolean

Open filename For Binary Access Write As #1

Put #1, , "<tanmayresource file>"






Close #1


End Function

Public Function addfile(ByVal filename As String, ByVal nametoappear As String) As Boolean
Dim tbuf As String
Dim Fbuf As String
Dim st As Single
Dim sp As Single
Dim Temp

If nametoappear <> "" Then
Open filename For Binary Access Read As #3

 tbuf = String(LOF(3), Chr(0))
Get #3, , tbuf

Close #3


Open temptrsname For Binary As #2
If getattr(nametoappear, st, sp) = False Then

Put #2, LOF(2) + 1, "<file<" & nametoappear & ">>"





x = Seek(2)
Put #2, Seek(2), tbuf

Put #2, LOF(2) + 1, "<\file<" & nametoappear & ">>"



Else

 Fbuf = String(LOF(2), Chr(0))
Get #2, 1, Fbuf

Temp1 = Left(Fbuf, st - 1)
temp2 = tbuf + Right(Fbuf, Len(Fbuf) - sp)

Fbuf = Left(Fbuf, st - 1) + "<file<" & nametoappear & ">>" + tbuf + Right(Fbuf, Len(Fbuf) - sp + 1)

Put #2, 1, Fbuf




End If


Close #2
End If

End Function
Public Function makepath(ByVal path As String) As String

If Right(path, 1) <> "\" Then
path = path + "\"
End If



makepath = path

End Function
Public Function checkpath(ByVal path As String)

Dim x As Integer

x = InStr(1, path, "\\")
Do Until x = 0

path = Left(path, x - 1) + Right(path, (Len(path) - x))
x = InStr(1, path, "\\")
Loop
 checkpath = path
End Function

Private Function getattr(ByVal name As String, Optional ByRef st As Single, Optional ByRef sp As Single) As Boolean
On Error GoTo handler
Dim fnum As Integer
Dim tbuf As String
Dim Temp


fnum = FreeFile
Open temptrsname For Binary Access Read As #fnum

tbuf = String(LOF(fnum), Chr(0))
Get #fnum, , tbuf


Temp = "<file<" & name & ">>"
st = InStr(1, tbuf, Temp)

Temp = "<\file<" & name & ">>"
sp = InStr(1, tbuf, Temp)




Close #fnum

If st = 0 Then
getattr = False
Else
getattr = True
End If

handler:



End Function
Public Function savefile(ByVal filename As String, ByVal destfname As String) As Boolean
Dim tbuf As String
Dim Fbuf As String
Dim st As Single
Dim sp As Single
Dim Temp As String

Open temptrsname For Binary Access Read As #2

 tbuf = String(LOF(2), Chr(0))
Get #2, , tbuf


Temp = "<file<" & filename & ">>"
st = InStr(1, tbuf, Temp)

Temp = "<\file<" & filename & ">>"
sp = InStr(1, tbuf, Temp)

If sp <> 0 Then

Temp = "<file<" & filename & ">>"
Temp = sp + Len(Temp) - st
Fbuf = Mid(tbuf, st + Len("<file<" & filename & ">>"), Temp)

Dim fnum As Integer
fnum = FreeFile

Open destfname For Output As #fnum

Print #fnum, Fbuf

Close #fnum

End If



Close #2


tbuf = ""
Fbuf = ""
Temp = ""
sp = 0








End Function
Public Sub createfilelist(ByRef listbox As listbox)
Dim tbuf As String
Open temptrsname For Binary Access Read As #4

tbuf = String(LOF(4), Chr(0))
Get #4, , tbuf

Dim x As Single
Dim y As Single
x = 1
listbox.Clear

Do While InStr(x, tbuf, "<file<") <> 0

x = InStr(x, tbuf, "<file<") + Len("<file<")
y = InStr(x, tbuf, ">>")

listbox.AddItem Mid(tbuf, x, y - x)

Loop

Close #4


End Sub
Public Function removefile(ByVal filename As String) As Boolean

Dim Fbuf As String
Dim st As Single
Dim sp As Single
Dim Temp

If filename <> "" Then



Open temptrsname For Binary As #2
If getattr(filename, st, sp) = True Then




 Fbuf = String(LOF(2), Chr(0))
Get #2, 1, Fbuf

Temp1 = Left(Fbuf, st - 1)
temp2 = Right(Fbuf, Len(Fbuf) - sp - Len("<\file<" & filename & ">>"))

Fbuf = Left(Fbuf, st - 3) + Right(Fbuf, Len(Fbuf) - sp - Len("<\file<" & filename & ">>") + 1)
'"<file<" & nametoappear & ">>"


Else


End If


Close #2


Open temptrsname For Output As #3
Print #3, Fbuf
Close #3

End If

End Function
Public Function Savetrs(ByVal filename As String) As Boolean
Dim Fbuf As String

Open temptrsname For Binary Access Read As #2
 Fbuf = String(LOF(2), Chr(0))
Get #2, , Fbuf
Close #2


Open filename For Output As #3
Print #3, Fbuf
Close #3





tbuf = ""











End Function

