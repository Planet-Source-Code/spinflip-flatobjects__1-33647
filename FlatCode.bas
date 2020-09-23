Attribute VB_Name = "FlatCode"
Option Explicit
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'I left these in here in case any of you wanted
'to modify the code a little, there are more
'than just these though...

Public Const WS_BORDER = &H800000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_DLGFRAME = &H400000
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_THICKFRAME = &H40000
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const WM_CLOSE = &H10
Public Const WS_CAPTION = &HC00000
Public Const WS_CHILD = &H40000000


Public Sub Flatten(myhwnd As Long)
Dim mystyle As Long
mystyle = GetWindowLong(myhwnd, -20)
mystyle = mystyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
SetWindowLong myhwnd, -20, mystyle
mystyle = GetWindowLong(myhwnd, -16)
mystyle = mystyle And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION Or WS_BORDER Or WS_SIZEBOX Or WS_THICKFRAME)
SetWindowLong myhwnd, -16, mystyle
SetWindowPos myhwnd, 0, 0, 0, 0, 0, &H20 Or &H10 Or &H4 Or &H2 Or &H1
End Sub
