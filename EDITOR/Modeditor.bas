Attribute VB_Name = "Modeditor"
Public Function makepath(ByVal path As String) As String

If Right(path, 1) <> "\" Then
path = path + "\"
End If



makepath = path




End Function
