Attribute VB_Name = "Module1"
Public Type RECT

    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' DrawText constants
Public Const DT_CENTER = &H1
Public Const DT_VCENTER = &H4
Public Const DT_SINGLELINE = &H20

' SetBkMode constants
Public Const OPAQUE = 2
Public Const TRANSPARENT = 1

Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Declare Function CreateEllipticRgn Lib "gdi32" _
    (ByVal X1 As Long, ByVal Y1 As Long, _
     ByVal X2 As Long, ByVal Y2 As Long) As Long

Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hbrush As Long) As Long


Declare Function CreateSolidBrush Lib "gdi32" _
    (ByVal crColor As Long) As Long

Declare Function SelectObject Lib "gdi32" _
    (ByVal hdc As Long, ByVal hObject As Long) As Long

Declare Function DrawText Lib "user32" Alias "DrawTextA" _
    (ByVal hdc As Long, ByVal lpStr As String, _
     ByVal nCount As Long, lpRect As RECT, _
     ByVal wFormat As Long) As Long

Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" _
    (ByVal lHeight As Long, ByVal lWidth As Long, _
     ByVal lEscapement As Long, ByVal lOrientation As Long, _
     ByVal lWeight As Long, ByVal lItalic As Long, _
     ByVal lUnderline As Long, ByVal lStrikeOut As Long, _
     ByVal lCharSet As Long, ByVal lOutPrecision As Long, _
     ByVal lClipPrecision As Long, ByVal lQuality As Long, _
     ByVal lPitch As Long, ByVal FaceName As String) As Long

Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long

Declare Function SetTextColor Lib "gdi32" _
    (ByVal hdc As Long, ByVal crColor As Long) As Long

Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Declare Function SetBkMode Lib "gdi32" _
    (ByVal hdc As Long, ByVal nBkMode As Long) As Long



