Attribute VB_Name = "Kernel"
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Public Sub TimeOut(HowLong)
Dim TheBeginning
Dim NoFreeze As Integer

TheBeginning = Timer
Do
    If Timer - TheBeginning >= HowLong Then Exit Sub
    NoFreeze% = DoEvents()
Loop
End Sub


