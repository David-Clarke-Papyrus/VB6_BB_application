Attribute VB_Name = "Module1"
Type POINTAPI
    x As Long
    Y As Long
End Type

Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Type WINDOWPLACEMENT
    Length As Long
    FLAGS As Long
    ShowCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type
 
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

'Public Function PointsToMe(hWnd As Long, XCoord As Long, YCoord As Long) As Boolean
'Dim WinPos As WINDOWPLACEMENT
'Dim Point As POINTAPI, lResult As Long
'
'lResult = GetCursorPos(Point)
'
'lResult = GetWindowPlacement(hWnd, WinPos)
'
''If Point.X >= WinPos.rcNormalPosition.Left And _
''        Point.X <= WinPos.rcNormalPosition.Right And _
''        Point.Y >= WinPos.rcNormalPosition.Top And _
''        Point.Y <= WinPos.rcNormalPosition.Bottom Then
'    PointsToMe = True
'  ScreenToClient hWnd, Point
'    XCoord = Point.x * Screen.TwipsPerPixelX
'    YCoord = Point.Y * Screen.TwipsPerPixelY
'
''End If
'End Function

'
'
'Type POINTAPI
'    x As Long
'    Y As Long
'End Type
'
'Type RECT
'    left As Long
'    top As Long
'    Right As Long
'    Bottom As Long
'End Type
'
'Type WINDOWPLACEMENT
'    Length As Long
'    FLAGS As Long
'    ShowCmd As Long
'    ptMinPosition As POINTAPI
'    ptMaxPosition As POINTAPI
'    rcNormalPosition As RECT
'End Type
'
''Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
''Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
''Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
''
''Public Function PointsToMe(hwnd As Long, XCoord As Long, YCoord As Long) As Boolean
''Dim WinPos As WINDOWPLACEMENT
''Dim Point As POINTAPI, lResult As Long
''lResult = GetCursorPos(Point)
''
''lResult = GetWindowPlacement(hwnd, WinPos)
''
'''If Point.X >= WinPos.rcNormalPosition.Left And _
'''        Point.X <= WinPos.rcNormalPosition.Right And _
'''        Point.Y >= WinPos.rcNormalPosition.Top And _
'''        Point.Y <= WinPos.rcNormalPosition.Bottom Then
''    PointsToMe = True
''  ScreenToClient hwnd, Point
''    XCoord = Point.x * Screen.TwipsPerPixelX
''    YCoord = Point.Y * Screen.TwipsPerPixelY
''
'''End If
''End Function
'
'
'
'>>>>>>> .merge-right.r375
