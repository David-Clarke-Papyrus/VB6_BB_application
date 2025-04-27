Attribute VB_Name = "WINDOWS_API"
Option Explicit

Type POINTAPI
    x As Long
    Y As Long
End Type

Type RECT
    Left As Long
    TOP As Long
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

Private Declare Function Beep Lib "kernel32" _
      (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long




Public Const FOL_EDI_PURCHASORDERS_SEND = "EDI\POs\OUT"
Public Const WM_CLOSE = &H10
Public Const WM_DESTROY = &H2
Const MAX_FILENAME_LEN = 260
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public defWinProc As Long

Public Const GWL_WNDPROC As Long = -4
Private Const CBN_DROPDOWN As Long = 7
Private Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const MK_CONTROL As Long = &H8
Public Const MK_SHIFT As Long = &H4
Private Const WM_KEYDOWN As Long = &H100
Public Const VK_LSHIFT = &HA0 ' Left SHIFT
Private Const VK_F4 As Long = &H73
Public Const KEYEVENTF_KEYUP As Long = &H2
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" _
    (ByVal hInstance As Long, lpIconName As Any) As Long

Public Const GWL_HWNDPARENT = (-8)

Public Const WM_GETICON = &H7F
Public Const WM_SETICON = &H80

Public Const ICON_SMALL = 0
Public Const ICON_BIG = 1
Declare Function SetWindowPos Lib "user32" _
      (ByVal hWnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal Y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long
      
Public Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
   (ByVal hWnd As Long, _
   ByVal nIndex As Long) As Long
   
Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hWnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
   
Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As _
  Long, lParam As String) As Long
Public Const LB_SETTABSTOPS = &H192

Public Const LB_SETSEL = &H185&

Const WM_USER = &H400
Const gXLENGTH = 24
'Public Const Nullstring As String = &O0
Public Const gMAXPOLRECs As Long = 2000
Public Const gMAXPERSRECs As Long = 100
Dim cnt

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 3
Public Const MAX_PATH = 260
Public Const ERRAPP_NOTRUNNING As Long = 429

Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function CallWindowProc Lib "user32" _
   Alias "CallWindowProcA" _
  (ByVal lpPrevWndFunc As Long, _
   ByVal hWnd As Long, ByVal msg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long



Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Declare Function GetVersionEx Lib "kernel32" _
Alias "GetVersionExA" (lpVersionInformation As _
OSVERSIONINFOEX) As Long


Public Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags  As Long
   lpfnCallback  As Long
   lParam As Long
   iImage As Long
End Type

Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
      
      
Private Declare Function GetClassName Lib "user32" _
    Alias "GetClassNameA" (ByVal hWnd&, _
    ByVal lpClassName$, ByVal nMaxCount&) As Long
      
      
      
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long


Private Const EM_CHARFROMPOS& = &HD7
'Private Type POINTAPI
'    x As Long
'    Y As Long
'End Type
Private Declare Function SendMessageLong Lib "user32" Alias _
    "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long




Public Const CL_LIGHTGREEN = 12648384
Public Const CL_PURPLE = 16384150
Public Const CL_DARKPURPLE = &H900090
Public Const CL_ORANGE = 25800
Public Const CL_LIGHTRED = 9869055
Public Const CL_DARKGREEN = &H8000&
Public Const CL_DARKBLUE = 13107200

'Windows API Constants
   Global Const CB_SHOWDROPDOWN = &H14F


Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
   hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal _
   dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal _
   dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject _
   As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal _
   dwMilliseconds As Long)
Declare Function CoCreateGuid_Alt Lib "OLE32.DLL" Alias "CoCreateGuid" (pGuid As Any) As Long
Declare Function StringFromGUID2_Alt Lib "OLE32.DLL" Alias "StringFromGUID2" (pGuid As Any, ByVal Address As Long, ByVal Max As Long) As Long


Private Declare Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const INFINITE = &HFFFF
Private Const SYNCHRONIZE = &H200000
Private Const WAIT_TIMEOUT = &H102


Private Const NETWORK_ALIVE_LAN = &H1  'net card connection
Private Const NETWORK_ALIVE_WAN = &H2  'RAS connection
Private Const NETWORK_ALIVE_AOL = &H4  'AOL
       
Private Declare Function IsNetworkAlive Lib "Sensapi.DLL" _
  (lpdwFlags As Long) As Long

Private Const HWND_BROADCAST As Long = &HFFFF&
Private Const WM_WININICHANGE As Long = &H1A

Private Declare Function GetProfileString Lib "kernel32" _
   Alias "GetProfileStringA" _
  (ByVal lpAppName As String, _
   ByVal lpkeyname As String, _
   ByVal lpDefault As String, _
   ByVal lpreturnedstring As String, _
   ByVal nSize As Long) As Long

Private Declare Function WriteProfileString Lib "kernel32" _
   Alias "WriteProfileStringA" _
  (ByVal lpszSection As String, _
   ByVal lpszKeyName As String, _
   ByVal lpszString As String) As Long
   
Private Declare Function SendNotifyMessage Lib "user32" _
   Alias "SendNotifyMessageA" _
  (ByVal hWnd As Long, _
   ByVal msg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

'********************************************
'*    (c) 1999-2000 Sergey Merzlikin        *
'********************************************

Private Const STATUS_TIMEOUT = &H102&
Private Const QS_KEY = &H1&
Private Const QS_MOUSEMOVE = &H2&
Private Const QS_MOUSEBUTTON = &H4&
Private Const QS_POSTMESSAGE = &H8&
Private Const QS_TIMER = &H10&
Private Const QS_PAINT = &H20&
Private Const QS_SENDMESSAGE = &H40&
Private Const QS_HOTKEY = &H80&
Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT _
        Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON _
        Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Private Declare Function MsgWaitForMultipleObjects Lib "user32" _
        (ByVal nCount As Long, pHandles As Long, _
        ByVal fWaitAll As Long, ByVal dwMilliseconds _
        As Long, ByVal dwWakeMask As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const STATUS_PENDING = &H103&
Private Const PROCESS_QUERY_INFORMATION = &H400

Private Const NERR_SUCCESS = 0&
Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2

Private Type TIME_OF_DAY_INFO
  tod_elapsedt As Long
  tod_msecs As Long
  tod_hours As Long
  tod_mins As Long
  tod_secs As Long
  tod_hunds As Long
  tod_timezone As Long
  tod_tinterval As Long
  tod_day As Long
  tod_month As Long
  tod_year As Long
  tod_weekday As Long
End Type

Private Type SYSTEMTIME
   wYear         As Integer
   wMonth        As Integer
   wDayOfWeek    As Integer
   wDay          As Integer
   wHour         As Integer
   wMinute       As Integer
   wSecond       As Integer
   wMilliseconds As Integer
End Type

Private Declare Function NetRemoteTOD Lib "netapi32" _
  (UncServerName As Byte, _
   BufferPtr As Long) As Long

Private Declare Function SetSystemTime Lib "kernel32" _
  (lpSystemTime As SYSTEMTIME) As Long

Private Declare Function NetLocalGroupEnum Lib "netapi32" _
  (servername As Byte, _
   ByVal Level As Long, _
   buff As Long, _
   ByVal buffsize As Long, _
   entriesread As Long, _
   totalentries As Long, _
   resumehandle As Long) As Long
   
Private Declare Function NetApiBufferFree Lib "netapi32" _
  (ByVal lpBuffer As Long) As Long
   
Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (pTo As Any, uFrom As Any, _
   ByVal lSize As Long)



Public Declare Function DrawText _
 Lib "user32.dll" Alias "DrawTextA" ( _
 ByVal hdc As Long, _
 ByVal lpStr As String, _
 ByVal nCount As Long, _
 ByRef lpRect As RECT, _
 ByVal wFormat As Long) As Long
Public Const DT_CALCRECT As Long = &H400
'Public Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type

Public Declare Function BitBlt Lib "gdi32" _
    (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, _
     ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, _
     ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'Public Declare Function DrawText Lib "user32" Alias "DrawTextA" _
'    (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, _
'     lpRect As RECT, ByVal wFormat As Long) As Long
Public Const DT_EDITCONTROL = &H2000&

Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_GETDROPPEDWIDTH = &H15F


Public Function GetFolder(pCaption As String) As String
'Opens a Treeview control that displays the directories in a computer
Dim lngpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

    szTitle = pCaption   '"Please select the database connection as it has either not been set or has been moved."
    With tBrowseInfo
       .hWndOwner = 0
       .lpszTitle = lstrcat(szTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lngpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lngpIDList) Then
       sBuffer = Space(MAX_PATH)
       SHGetPathFromIDList lngpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        ' comment buy Urs:
        ' The path name will be saved in z_DatabasePersist.dbConnect only if DB is opened
        ' successfuly, else it will be saved as empty string to force the select path box
        ' to be opened again....
'       SaveSetting "Dispatcher", "Settings", "Databasefolder", sBuffer
        
       GetFolder = sBuffer
   End If
End Function
Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) _
   As Long

   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, _
         0, FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, _
         0, 0, FLAGS)
      SetTopMostWindow = False
   End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you can not publish
'               or reproduce this code on any web site,
'               on any online service, or distribute on
'               any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Public Sub Unhook(hWnd As Long)
    
   If defWinProc <> 0 Then
   
      Call SetWindowLong(hWnd, _
                         GWL_WNDPROC, _
                         defWinProc)
      defWinProc = 0
   End If
    
End Sub


Public Sub Hook(hWnd As Long)

   'Don't hook twice or you will
   'be unable to unhook it.
    If defWinProc = 0 Then
    
      defWinProc = SetWindowLong(hWnd, _
                                 GWL_WNDPROC, _
                                 AddressOf WindowProc)
      
    End If
    
End Sub


Public Function WindowProc(ByVal hWnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
  
'  'only if the window is the combo box...
' '  If hwnd = Form1.Combo1.hwnd Then
'
'      Select Case uMsg
'
'         Case CBN_DROPDOWN  'the list box of a combo
'                            'box is about to be made visible.
'
'           'return 1 to indicate we ate the message
'            WindowProc = 1
'
'         Case WM_KEYDOWN   'prevent the F4 key from showing
'                           'the combo's list
'
'            If wParam = VK_F4 Then
'
'              'set up the parameters as though a
'              'mouse click occurred on the combo,
'              'and call this routine again
'               Call WindowProc(hwnd, WM_LBUTTONDOWN, 1, 2000)
'
'            Else
'
'              'there's nothing to do keyboard-wise
'              'with the combo, so return 1 to
'              'indicate we ate the message
'               WindowProc = 1
'
'            End If
'
'         Case WM_LBUTTONDOWN  'process mouse clicks
'
'           'if the list is hidden, position and show it
'            If Form1.List1.Visible = False Then
'
'               With Form1
'                  .List1.Left = .Combo1.Left
'                  .List1.Width = .Combo1.Width
'                  .List1.Top = .Combo1.Top + .Combo1.Height + 1
'                  .List1.Visible = True
'                  .List1.SetFocus
'               End With
'
'            Else
'
'              'the list must be visible, so hide it
'               Form1.List1.Visible = False
'            End If
'
'           'return 1 to indicate we processed the message
'            WindowProc = 1
'
'         Case Else
'
'           'call the default window ErrHandler
'            WindowProc = CallWindowProc(defWinProc, _
'                                        hwnd, _
'                                        uMsg, _
'                                        wParam, _
'                                        lParam)
'
'      End Select
'
' '  End If  'If hwnd = Form1.Combo1.hwnd
   
End Function
'--end block--'
 

'Public Function CurrentOS() As String
'Dim sys As New s
'   Select Case sys.OSPlatform
'      Case 0
'         CurrentOS = "Unknown"
'      Case 1
'        CurrentOS = "Win95"
'   '      MsgEnd = "Windows 95, ver. " & CStr(sysDetectOS.OSVersion) & "(" & CStr(sysDetectOS.OSBuild) & ")"
'      Case 2
'        CurrentOS = "WinNT"
'     '    MsgEnd = "Windows NT, ver. " & CStr(sysDetectOS.OSVersion) & "(" & CStr(sysDetectOS.OSBuild) & ")"
'   End Select
'End Function



Public Sub ShellandWaitold(PathName, Optional WindowStyle As _
   VbAppWinStyle = vbMinimizedFocus, Optional bDoEvents As _
   Boolean = False)
 '   On Error GoTo ErrHandler

    Dim dwProcessId As Long
    Dim hProcess As Long
    
    dwProcessId = Shell(PathName, WindowStyle)
    
    If dwProcessId = 0 Then
        Exit Sub
    End If
    
    hProcess = OpenProcess(SYNCHRONIZE, False, dwProcessId)
    
    If hProcess = 0 Then
        Exit Sub
    End If
    
    If bDoEvents Then
        Do While WaitForSingleObject(hProcess, 100) = WAIT_TIMEOUT
            DoEvents
        Loop
    Else
        WaitForSingleObject hProcess, INFINITE
    End If
    
    CloseHandle hProcess
    
    Exit Sub
End Sub
Public Function ShellandWait(ExeFullPath As String, _
Optional TimeOutValue As Long = 0) As Boolean
    
    Dim lInst As Long
    Dim lStart As Long
    Dim lTimeToQuit As Long
    Dim sExeName As String
    Dim lProcessId As Long
    Dim lExitCode As Long
    Dim bPastMidnight As Boolean
    
    On Error GoTo ErrorHandler

    lStart = CLng(Timer)
    sExeName = ExeFullPath

    'Deal with timeout being reset at Midnight
    If TimeOutValue > 0 Then
        If lStart + TimeOutValue < 86400 Then
            lTimeToQuit = lStart + TimeOutValue
        Else
            lTimeToQuit = (lStart - 86400) + TimeOutValue
            bPastMidnight = True
        End If
    End If

    lInst = Shell(sExeName, vbHide)
    
lProcessId = OpenProcess(PROCESS_QUERY_INFORMATION, False, lInst)

    Do
        Call GetExitCodeProcess(lProcessId, lExitCode)
        DoEvents
        If TimeOutValue And Timer > lTimeToQuit Then
            If bPastMidnight Then
                 If Timer < lStart Then Exit Do
            Else
                 Exit Do
            End If
        End If
    Loop While lExitCode = STATUS_PENDING
    
    ShellandWait = True
   
ErrorHandler:
ShellandWait = False
Exit Function
End Function



Function IsNetConnectionAlive() As Boolean

   Dim tmp As Long
   IsNetConnectionAlive = IsNetworkAlive(tmp) = 1
   
End Function

Public Function OSVersion() As String
    
    Dim udtOSVersion As OSVERSIONINFOEX
    Dim lMajorVersion  As Long
    Dim lMinorVersion As Long
    Dim lPlatformID As Long
    Dim sAns As String
    
    
    udtOSVersion.dwOSVersionInfoSize = Len(udtOSVersion)
    GetVersionEx udtOSVersion
    lMajorVersion = udtOSVersion.dwMajorVersion
    lMinorVersion = udtOSVersion.dwMinorVersion
    lPlatformID = udtOSVersion.dwPlatformId
    
    Select Case lMajorVersion
        Case 5
        
            ' Added the following to give suppport for Windows XP!
            If lMinorVersion = 0 Then
                sAns = "Windows 2000"
            ElseIf lMinorVersion = 1 Then
                sAns = "Windows XP"
            End If
        Case 4
            If lPlatformID = VER_PLATFORM_WIN32_NT Then
                sAns = "Windows NT 4.0"
            Else
                sAns = IIf(lMinorVersion = 0, _
                "Windows 95", "Windows 98")
            End If
        Case 3
            If lPlatformID = VER_PLATFORM_WIN32_NT Then
                sAns = "Windows NT 3.x"
 
              'below should only happen if person has Win32s
                'installed
            Else
                sAns = "Windows 3.x"
            End If
            
        Case Else
            sAns = "Unknown Windows Version"
    End Select
                    
    OSVersion = sAns
    
End Function

Function GetPDFExecutable(pfile As String) As String
    'KPD-Team 1999
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
   Dim i As Integer, s2 As String
'   Const sFile = "C:\PBKS\Printing\IN_Tmp.PDF"

   'Check if the file exists
'   If Dir(sFile) = "" Or sFile = "" Then
'        MsgBox "File not found!", vbCritical
'        Exit Function
'   End If
   'Create a buffer
   s2 = String(MAX_FILENAME_LEN, 32)
   'Retrieve the name and handle of the executable, associated with this file
   i = FindExecutable(pfile, vbNullString, s2)
   If i > 32 Then
      GetPDFExecutable = Left$(s2, InStr(s2, Chr$(0)) - 1)
   Else
      GetPDFExecutable = ""
      LogSaveToFile "No executable associated with file:" & pfile & "."
   End If
End Function
Sub SetDefaultPrinter(pPrinter As String)
    On Error GoTo errHandler
Dim p As Printer
Dim tmp As String
Dim wshNetwork, strPrinterPath As String

'    If mDebugmodeOn Then
'        EnumeratePrinter
'        LogSaveToFile "SetDefaultPrinter;1: pPrinter=" & pPrinter
'    End If

    tmp = ""
    For Each p In Printers
        If InStr(1, UCase(p.DeviceName), UCase(pPrinter)) > 0 Then
            tmp = p.DeviceName
            Exit For
        End If
    Next
'    If mDebugmodeOn Then
'        LogSaveToFile "SetDefaultPrinter;2: tmp=" & tmp
'    End If
      
    If tmp > "" Then
        Set wshNetwork = CreateObject("WScript.Network")
'        If mDebugmodeOn Then
'            MsgBox "pos 3"
'        End If
        wshNetwork.SetDefaultPrinter tmp
'        If mDebugmodeOn Then
'            MsgBox "pos 4: after setting new default printer"
'        End If
    Else
    
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WINDOWS_API.SetDefaultPrinter(pPrinter)", pPrinter
End Sub
Private Sub EnumeratePrinter()
Dim oPrinters As Object
Dim wshNetwork
  Dim i As Integer
Dim s As String

    Set wshNetwork = CreateObject("WScript.Network")
    Set oPrinters = wshNetwork.EnumPrinterConnections
    s = ""
    MsgBox "oPrinters.Count: " & oPrinters.Count
    For i = 0 To oPrinters.Count - 1 Step 2
        MsgBox "Port " & oPrinters.Item(i) & " = " & oPrinters.Item(i + 1)
       s = s & IIf(s > "", vbCrLf, "") & "Port " & oPrinters.Item(i) & " = " & oPrinters.Item(i + 1)
    Next
 '   LogSaveToFile "EnumeratePrinter: set= " & s
    
End Sub
' The MsgWaitObj function replaces Sleep,
' WaitForSingleObject, WaitForMultipleObjects functions.
' Unlike these functions, it
' doesn't block thread messages processing.
' Using instead Sleep:
'     MsgWaitObj dwMilliseconds
' Using instead WaitForSingleObject:
'     retval = MsgWaitObj(dwMilliseconds, hObj, 1&)
' Using instead WaitForMultipleObjects:
'     retval = MsgWaitObj(dwMilliseconds, hObj(0&), n),
'     where n - wait objects quantity,
'     hObj() - their handles array.

Public Function MsgWaitObj(Interval As Long, _
            Optional hObj As Long = 0&, _
            Optional nObj As Long = 0&) As Long
Dim T As Long, t1 As Long
If Interval <> INFINITE Then
    T = GetTickCount()
    On Error Resume Next
    T = T + Interval
    ' Overflow prevention
    If Err <> 0& Then
        If T > 0& Then
            T = ((T + &H80000000) _
            + Interval) + &H80000000
        Else
            T = ((T - &H80000000) _
            + Interval) - &H80000000
        End If
    End If
    On Error GoTo 0
    ' T contains now absolute time of the end of interval
Else
    t1 = INFINITE
End If
Do
    If Interval <> INFINITE Then
        t1 = GetTickCount()
        On Error Resume Next
     t1 = T - t1
        ' Overflow prevention
        If Err <> 0& Then
            If T > 0& Then
                t1 = ((T + &H80000000) _
                - (t1 - &H80000000))
            Else
                t1 = ((T - &H80000000) _
                - (t1 + &H80000000))
            End If
        End If
        On Error GoTo 0
        ' T1 contains now the remaining interval part
        If IIf((t1 Xor Interval) > 0&, _
            t1 > Interval, t1 < 0&) Then
            ' Interval expired
            ' during DoEvents
            MsgWaitObj = STATUS_TIMEOUT
            Exit Function
        End If
    End If
    ' Wait for event, interval expiration
    ' or message appearance in thread queue
    MsgWaitObj = MsgWaitForMultipleObjects(nObj, _
            hObj, 0&, t1, QS_ALLINPUT)
    ' Let's message be processed
    DoEvents
    If MsgWaitObj <> nObj Then Exit Function
    ' It was message - continue to wait
Loop
End Function



' Open the default browser on a given URL
' Returns True if successful, False otherwise

Public Function OpenBrowser(ByVal URL As String, Optional bIsFile As Boolean) As Boolean
    Dim Res As Long
    
    ' it is mandatory that the URL is prefixed with http:// or https://
    If InStr(1, URL, "http", vbTextCompare) <> 1 And Not bIsFile Then
        URL = "http://" & URL
    End If
    
    Res = ShellExecute(0&, "open", URL, vbNullString, vbNullString, vbNormalFocus)
    OpenBrowser = (Res > 32)
End Function


Function CreateGUID() As String
    Dim Res As String, resLen As Long, guid(15) As Byte
    Res = Space$(128)
    CoCreateGuid_Alt guid(0)
    resLen = StringFromGUID2_Alt(guid(0), ByVal StrPtr(Res), 128)
    CreateGUID = Left$(Res, resLen - 1)
End Function




' Return the character position under the mouse.
Public Function TextBoxCursorPos(ByVal txt As TextBox, _
    ByVal x As Single, ByVal Y As Single) As Long
    ' Convert the position to pixels.
    x = x \ Screen.TwipsPerPixelX
    Y = Y \ Screen.TwipsPerPixelY

    ' Get the character number
    TextBoxCursorPos = SendMessageLong(txt.hWnd, _
        EM_CHARFROMPOS, 0&, CLng(x + Y * &H10000)) And _
        &HFFFF&
End Function

Public Sub SetFileProperty(pFilePath As String, pProperty As String, val As String)
Dim m_oDocumentProps As DSOFile.OleDocumentProperties
Dim oSummProps As DSOFile.SummaryProperties
    Set m_oDocumentProps = New DSOFile.OleDocumentProperties
    m_oDocumentProps.open pFilePath, False, dsoOptionDefault

   On Error Resume Next
   Set oSummProps = m_oDocumentProps.SummaryProperties
    If pProperty = "Author" Then oSummProps.Author = val
    If pProperty = "Title" Then oSummProps.Title = val
    If pProperty = "Category" Then oSummProps.Category = val
    m_oDocumentProps.Save
    m_oDocumentProps.Close
   Set oSummProps = Nothing
   Set m_oDocumentProps = Nothing
End Sub


Private Function SynchronizeTOD(ByVal sRemoteServer As String) As Date
  
   Dim newdate  As Date
   Dim sys_sync As SYSTEMTIME
   Dim server_date As TIME_OF_DAY_INFO
   Dim local_date As TIME_OF_DAY_INFO
  
  'Obtain a TIME_OF_DAY_INFO structure from the
  'remote machine with which to synchronize to.
   server_date = GetRemoteTOD(sRemoteServer)
   
  'case returned values into a SYSTEMTIME structure
  'and pass to the SetSystemTime api
   With sys_sync
      .wHour = server_date.tod_hours
      .wMinute = server_date.tod_mins
      .wSecond = server_date.tod_secs
      .wDay = server_date.tod_day
      .wMonth = server_date.tod_month
      .wYear = server_date.tod_year
   End With
   
   If SetSystemTime(sys_sync) <> 0 Then
   
    'sync was successful, so return Now
     SynchronizeTOD = Now
   
   End If
   
   
'''  '--- for demo only ---
'''  'The first shows calculating the
'''  'date using the tod_elapsedt member.
'''  'tod_elapsedt is a value that contains
'''  'the number of seconds since
'''  '00:00:00, January 1, 1970, GMT.
'''  'Since tod_elapsedt is based on GMT (UTC),
'''  'the next date applies the tod_timezone
'''  'offset to adjust the date to the local time.
'''   newdate = DateAdd("s", server_date.tod_elapsedt, #1/1/1970#)
'''   Text2.Text = newdate
'''   newdate = DateAdd("n", -server_date.tod_timezone, newdate)
'''   Text3.Text = newdate
  '-----------------------
 
End Function
Private Function GetRemoteTOD(ByVal sServer As String) As TIME_OF_DAY_INFO

   Dim bServer()  As Byte
   Dim tod        As TIME_OF_DAY_INFO
   Dim bufptr     As Long

  'A null passed as sServer retrieves
  'the date for the local machine. If
  'sServer is null, no slashes are added.
   If sServer <> vbNullChar Then
    
     'If a server name was specified,
     'assure it has leading double slashes
      If Left$(sServer, 2) <> "\\" Then
         bServer = "\\" & sServer & vbNullChar
      Else
         bServer = sServer & vbNullChar
      End If
      
   Else
   
     'null or empty string was passed
      bServer = sServer & vbNullChar
   
   End If
   
   
  'get the time of day (TOD) from the specified server
   If NetRemoteTOD(bServer(0), bufptr) = NERR_SUCCESS Then

     'copy the buffer into a
     'TIME_OF_DAY_INFO structure
      CopyMemory tod, ByVal bufptr, LenB(tod)

   End If
   
   Call NetApiBufferFree(bufptr)
   
  'return the TIME_OF_DAY_INFO structure
   GetRemoteTOD = tod

End Function


Public Function AutoSizeDropDownWidth(Combo As Object) As Boolean
'**************************************************************
'PURPOSE: Automatically size the combo box drop down width
'         based on the width of the longest item in the combo box

'PARAMETERS: Combo - ComboBox to size

'RETURNS: True if successful, false otherwise

'ASSUMPTIONS: 1. Form's Scale Mode is vbTwips, which is why
'                conversion from twips to pixels are made.
'                API functions require units in pixels
'
'             2. Combo Box's parent is a form or other
'                container that support the hDC property

'EXAMPLE: AutoSizeDropDownWidth Combo1
'****************************************************************
Dim lRet As Long
Dim bAns As Boolean
Dim lCurrentWidth As Single
Dim rectCboText As RECT
Dim lParentHDC As Long
Dim lListCount As Long
Dim lCtr As Long
Dim lTempWidth As Long
Dim lWidth As Long
Dim sSavedFont As String
Dim sngSavedSize As Single
Dim bSavedBold As Boolean
Dim bSavedItalic As Boolean
Dim bSavedUnderline As Boolean
Dim bFontSaved As Boolean

On Error GoTo ErrorHandler

If Not TypeOf Combo Is ComboBox Then Exit Function
lParentHDC = Combo.Parent.hdc
If lParentHDC = 0 Then Exit Function
lListCount = Combo.ListCount
If lListCount = 0 Then Exit Function


'Change font of parent to combo box's font
'Save first so it can be reverted when finished
'this is necessary for drawtext API Function
'which is used to determine longest string in combo box
With Combo.Parent

    sSavedFont = .FontName
    sngSavedSize = .FontSize
    bSavedBold = .FontBold
    bSavedItalic = .FontItalic
    bSavedUnderline = .FontUnderline
    
    .FontName = Combo.FontName
    .FontSize = Combo.FontSize
    .FontBold = Combo.FontBold
    .FontItalic = Combo.FontItalic
    .FontUnderline = Combo.FontItalic

End With

bFontSaved = True

'Get the width of the largest item
For lCtr = 0 To lListCount
   DrawText lParentHDC, Combo.List(lCtr), -1, rectCboText, _
        DT_CALCRECT
   'adjust the number added (20 in this case to
   'achieve desired right margin
   lTempWidth = rectCboText.Right - rectCboText.Left + 20

   If (lTempWidth > lWidth) Then
      lWidth = lTempWidth
   End If
Next
 
lCurrentWidth = SendMessageLong(Combo.hWnd, CB_GETDROPPEDWIDTH, _
    0, 0)

If lCurrentWidth > lWidth Then 'current drop-down width is
'                               sufficient

    AutoSizeDropDownWidth = True
    GoTo ErrorHandler
    Exit Function
End If
 
'don't allow drop-down width to
'exceed screen.width
 
   If lWidth > Screen.Width \ Screen.TwipsPerPixelX - 20 Then _
    lWidth = Screen.Width \ Screen.TwipsPerPixelX - 20

lRet = SendMessageLong(Combo.hWnd, CB_SETDROPPEDWIDTH, lWidth, 0)

AutoSizeDropDownWidth = lRet > 0
ErrorHandler:
On Error Resume Next
If bFontSaved Then
'restore parent's font settings
  With Combo.Parent
    .FontName = sSavedFont
    .FontSize = sngSavedSize
    .FontUnderline = bSavedUnderline
    .FontBold = bSavedBold
    .FontItalic = bSavedItalic
 End With
End If
End Function


Public Function fRunningInIde(hWnd As Long) As Boolean
    On Error GoTo errHandler
Dim sClassName As String
Dim nStrLen    As Long

    '
    ' See if we're running in the IDE.
    '
    sClassName = String$(260, vbNullChar)
    nStrLen = GetClassName(hWnd, sClassName, Len(sClassName))
    If nStrLen Then sClassName = Left$(sClassName, nStrLen)
    
    fRunningInIde = (sClassName = "ThunderMDIForm")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.fRunningInIde"
End Function

Public Function PointsToMe(hWnd As Long, XCoord As Long, YCoord As Long) As Boolean
Dim WinPos As WINDOWPLACEMENT
Dim Point As POINTAPI, lResult As Long

lResult = GetCursorPos(Point)

lResult = GetWindowPlacement(hWnd, WinPos)

'If Point.X >= WinPos.rcNormalPosition.Left And _
'        Point.X <= WinPos.rcNormalPosition.Right And _
'        Point.Y >= WinPos.rcNormalPosition.Top And _
'        Point.Y <= WinPos.rcNormalPosition.Bottom Then
    PointsToMe = True
  ScreenToClient hWnd, Point
    XCoord = Point.x * Screen.TwipsPerPixelX
    YCoord = Point.Y * Screen.TwipsPerPixelY
    
'End If
End Function

