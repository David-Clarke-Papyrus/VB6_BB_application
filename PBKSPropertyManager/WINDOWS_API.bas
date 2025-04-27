Attribute VB_Name = "WINDOWS_API"
Option Explicit
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
Private Const WM_KEYDOWN As Long = &H100
Private Const VK_F4 As Long = &H73
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" _
    (ByVal hInstance As Long, lpIconName As Any) As Long

Public Const GWL_HWNDPARENT = (-8)

Public Const WM_GETICON = &H7F
Public Const WM_SETICON = &H80

Public Const ICON_SMALL = 0
Public Const ICON_BIG = 1
Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal Y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long
      
Public Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
   (ByVal hwnd As Long, _
   ByVal nIndex As Long) As Long
   
Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hwnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
   
Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
   
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As _
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
   ByVal hwnd As Long, ByVal msg As Long, _
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

Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
      
      
      
      
      
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long


Private Const EM_CHARFROMPOS& = &HD7
Private Type POINTAPI
    x As Long
    Y As Long
End Type
Private Declare Function SendMessageLong Lib "user32" Alias _
    "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
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
  (ByVal hwnd As Long, _
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
Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
   As Long

   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
         0, FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
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



Public Sub Unhook(hwnd As Long)
    
   If defWinProc <> 0 Then
   
      Call SetWindowLong(hwnd, _
                         GWL_WNDPROC, _
                         defWinProc)
      defWinProc = 0
   End If
    
End Sub


Public Sub Hook(hwnd As Long)

   'Don't hook twice or you will
   'be unable to unhook it.
    If defWinProc = 0 Then
    
      defWinProc = SetWindowLong(hwnd, _
                                 GWL_WNDPROC, _
                                 AddressOf WindowProc)
      
    End If
    
End Sub


Public Function WindowProc(ByVal hwnd As Long, _
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
    
    On Error GoTo errorHandler

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
   
errorHandler:
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
'Private Sub EnumeratePrinter()
'Dim oPrinters As Object
'Dim wshNetwork
'  Dim i As Integer
'Dim s As String
'
'    Set wshNetwork = CreateObject("WScript.Network")
'    Set oPrinters = wshNetwork.EnumPrinterConnections
'    s = ""
'    MsgBox "oPrinters.Count: " & oPrinters.Count
'    For i = 0 To oPrinters.Count - 1 Step 2
'        MsgBox "Port " & oPrinters.Item(i) & " = " & oPrinters.Item(i + 1)
'       s = s & IIf(s > "", vbCrLf, "") & "Port " & oPrinters.Item(i) & " = " & oPrinters.Item(i + 1)
'    Next
'    LogSaveToFile "EnumeratePrinter: set= " & s
'
'End Sub
'' The MsgWaitObj function replaces Sleep,
'' WaitForSingleObject, WaitForMultipleObjects functions.
'' Unlike these functions, it
'' doesn't block thread messages processing.
'' Using instead Sleep:
''     MsgWaitObj dwMilliseconds
'' Using instead WaitForSingleObject:
''     retval = MsgWaitObj(dwMilliseconds, hObj, 1&)
'' Using instead WaitForMultipleObjects:
''     retval = MsgWaitObj(dwMilliseconds, hObj(0&), n),
''     where n - wait objects quantity,
''     hObj() - their handles array.

Public Function MsgWaitObj(Interval As Long, _
            Optional hObj As Long = 0&, _
            Optional nObj As Long = 0&) As Long
Dim T As Long, T1 As Long
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
    T1 = INFINITE
End If
Do
    If Interval <> INFINITE Then
        T1 = GetTickCount()
        On Error Resume Next
     T1 = T - T1
        ' Overflow prevention
        If Err <> 0& Then
            If T > 0& Then
                T1 = ((T + &H80000000) _
                - (T1 - &H80000000))
            Else
                T1 = ((T - &H80000000) _
                - (T1 + &H80000000))
            End If
        End If
        On Error GoTo 0
        ' T1 contains now the remaining interval part
        If IIf((T1 Xor Interval) > 0&, _
            T1 > Interval, T1 < 0&) Then
            ' Interval expired
            ' during DoEvents
            MsgWaitObj = STATUS_TIMEOUT
            Exit Function
        End If
    End If
    ' Wait for event, interval expiration
    ' or message appearance in thread queue
    MsgWaitObj = MsgWaitForMultipleObjects(nObj, _
            hObj, 0&, T1, QS_ALLINPUT)
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
    TextBoxCursorPos = SendMessageLong(txt.hwnd, _
        EM_CHARFROMPOS, 0&, CLng(x + Y * &H10000)) And _
        &HFFFF&
End Function

Public Sub SetFileProperty(pFilePath As String, pProperty As String, val As String)
Dim m_oDocumentProps As DSOFile.OleDocumentProperties
Dim oSummProps As DSOFile.SummaryProperties
    Set m_oDocumentProps = New DSOFile.OleDocumentProperties
    m_oDocumentProps.Open pFilePath, False, dsoOptionDefault

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



