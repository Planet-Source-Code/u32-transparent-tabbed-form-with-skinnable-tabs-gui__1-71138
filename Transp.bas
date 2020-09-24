Attribute VB_Name = "mdlTransp"
Option Explicit

' Made by: Found at MSDN

' Transparency bas.
' Makes a specified rgb value transparent
' Works only on XP or newer.

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Const LWA_COLORKEY = &H1
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000

Public Function Transparency(frm As Form, R As Integer, G As Integer, B As Integer)
    
    Dim RetVal, Color As Long
    
    Color = RGB(R, G, B)
    RetVal = RetVal Or WS_EX_LAYERED
    SetWindowLong frm.hWnd, GWL_EXSTYLE, RetVal
    SetLayeredWindowAttributes frm.hWnd, Color, 0, LWA_COLORKEY
    
End Function

