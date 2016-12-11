Attribute VB_Name = "Module2"
Option Explicit
Const Hwndx = -1
Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long


Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long _
, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, _
ByVal nHeight As Long, ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long



