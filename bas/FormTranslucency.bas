Attribute VB_Name = "FormTranslucency"
Option Explicit
Dim mbWindows98orHigher As Boolean
'
' Constants used by AnimateWindow
'
'Duration of animation in milliseconds
Const AW_DURATION_DEFAULT = 200
'Animate from left to right
Const AW_HOR_POSITIVE = &H1
'Animate from right to left
Const AW_HOR_NEGATIVE = &H2
'Animate from top to bottom
Const AW_VER_POSITIVE = &H4
'Animate from bottom to top
Const AW_VER_NEGATIVE = &H8
'Collapse window inward when used with
'  AW_HIDE or outward otherwise
Const AW_CENTER = &H10
'Hides the window
Const AW_HIDE = &H10000
'Activates the window
Const AW_ACTIVATE = &H20000
'Slide animation. Cannot use with AW_CENTER
Const AW_SLIDE = &H40000
'Fade window. Only works with top level windows
Const AW_BLEND = &H80000

Const VER_PLATFORM_WIN32s = 0
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Declare Function AnimateWindow Lib "user32" ( _
    ByVal hWnd As Long, ByVal dwTime As Long, _
    ByVal dwFlags As Long) As Boolean

Private Declare Function GetVersionEx Lib "kernel32" Alias _
    "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'hWnd    - handle to window to layer.
'crKey   - specifies the color key
'bAlpha  - value for the blend function
'dwFlags - action
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
    ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, _
    ByVal dwFlags As Long) As Long
    
Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long
    
Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1&
Private Const LWA_ALPHA = &H2&

Private Declare Function GetParent Lib "user32" _
    (ByVal hWnd As Long) As Long
    
Private Declare Function IsWindowVisible Lib "user32" _
    (ByVal hWnd As Long) As Long



Public Function fGetOSVersion()
Dim os As OSVERSIONINFO
'
' Returns True if Win98 or Win2000
'
fGetOSVersion = False
With os
    .dwOSVersionInfoSize = Len(os)
    Call GetVersionEx(os)

    ' Windows 2000
    If .dwMajorVersion > 4 Then fGetOSVersion = True

    If .dwMajorVersion = 4 And _
       .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS And _
       .dwMinorVersion > 0 Then
        fGetOSVersion = True
    End If
End With
End Function






Public Function fSetTranslucency(ByVal hWnd As Long, ByVal alpha As Byte) As Boolean
Dim lStyle As Long

'
' Layering only works with Win2K or above.
'
If fIsWin2000 Then
    '
    ' Only a top level window can be translucent.
    '
    hWnd = fGetTopLevel(hWnd)
    '
    ' Make the window translucent by setting its
    ' extended style.
    '
    lStyle = GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    If SetWindowLong(hWnd, GWL_EXSTYLE, lStyle) Then
        fSetTranslucency = CBool(SetLayeredWindowAttributes(hWnd, 0, CLng(alpha), LWA_ALPHA))
    End If
End If
End Function
Public Function fClearTranslucency(ByVal hWnd As Long) As Boolean
Dim lStyle As Long

'
' Layering only works with Win2K or above.
'
If fIsWin2000 Then
    '
    ' Only a top level window can be translucent.
    '
    hWnd = fGetTopLevel(hWnd)
    '
    ' Clear translucency - make the window opaque.
    '
    Call SetLayeredWindowAttributes(hWnd, 0, 255&, LWA_ALPHA)
    '
    ' Clear the extended style bit.
    '
    lStyle = GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_LAYERED
    fClearTranslucency = CBool(SetWindowLong(hWnd, GWL_EXSTYLE, lStyle))
End If
End Function


Private Function fIsWin2000() As Boolean
Dim os As OSVERSIONINFO
'
' Returns True if Win98 or Win2000
'
fIsWin2000 = False
With os
    .dwOSVersionInfoSize = Len(os)
    Call GetVersionEx(os)

    ' Windows 2000
    If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
        fIsWin2000 = (.dwMajorVersion > 4)
    End If
End With

End Function

Public Function fGetTopLevel(ByVal hChild As Long) As Long
Dim hWnd As Long

hWnd = hChild
Do While IsWindowVisible(GetParent(hWnd))
    hWnd = GetParent(hChild)
    hChild = hWnd
Loop
fGetTopLevel = hWnd
End Function

'Private Sub lblClear_Click()
'
'Call fClearTranslucency(Me.hWnd)
'End Sub
'
'Private Sub lblMake_Click()
''
'' Try values between 0 (completely invisible)
'' to 255 (fully opaque).
''
'Call fSetTranslucency(Me.hWnd, 160)
'End Sub




