Attribute VB_Name = "Modapis"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function Rectangle Lib "gdi32" _
            (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, _
            ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GdiFlush Lib "gdi32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long





Public Function Intersect(ByVal X1 As Single, ByVal Y1 As Single, ByVal w1 As Single, ByVal h1 As Single, _
                                  ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal w2 As Single, Optional ByVal h2 As Single) As Boolean
                                  



If X2 >= (X1 - w2) And X2 <= (X1 + w1) Then

If Y2 >= (Y1 - h2) And Y2 <= (Y1 + h1) Then

Intersect = True

End If
End If


End Function

