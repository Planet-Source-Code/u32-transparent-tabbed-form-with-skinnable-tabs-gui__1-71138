Attribute VB_Name = "mdlMoveForm"
Option Explicit

' Made by: Found at MSDN

' Move Me around the screen.
Public Declare Function ReleaseCapture Lib "user32.dll" () As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function MoveMe(frm As Form, Optional Ctl As Control)
    
    ReleaseCapture
    SendMessage frm.hWnd, &HA1, 2, 0&
    
End Function
