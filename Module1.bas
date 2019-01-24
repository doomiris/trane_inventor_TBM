Attribute VB_Name = "Module1"
'Option Explicit
'
'Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'
'Public Const HWND_TOPMOST = -1
'Public Const SWP_NOMOVE = &H2
'Public Const SWP_NOSIZE = &H1
'Public Const SWP_FRAMECHANGED = &H20
'Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
'Public Const HWND_BOTTOM = 1
'Public Const HWND_BROADCAST = &HFFFF&
'Public Const HWND_DESKTOP = 0
'Public Const HWND_NOTOPMOST = -2
'Public Const HWND_TOP = 0
'Public Const SWP_HIDEWINDOW = &H80
'Public Const SWP_ACTIVATE = &H10
'Public Const SWP_NOCOPYBITS = &H100
'Public Const SWP_NOOWNERZORDER = &H200
'Public Const SWP_NOREDRAW = &H8
'Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
'Public Const SWP_NOZORDER = &H4
'Public Const SWP_SHOWWINDOW = &H40
'Public Const Flags = SWP_DRAWFRAME Or SWP_NOMOVE Or SWP_NOSIZE
