Attribute VB_Name = "WINDOWS_API"
Option Explicit

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
   


Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As _
  Long, lParam As String) As Long
Public Const LB_SETTABSTOPS = &H192

Public Const LB_SETSEL = &H185&

Const WM_USER = &H400
Const gXLENGTH = 24
'Public Const Nullstring As String = &O0
Public Const gMAXPOLRECs As Long = 1000
Public Const gMAXPERSRECs As Long = 100
Dim cnt

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 3
Public Const MAX_PATH = 260
Public Const ERR_APP_NOTRUNNING As Long = 429

Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function CallWindowProc Lib "user32" _
   Alias "CallWindowProcA" _
  (ByVal lpPrevWndFunc As Long, _
   ByVal hwnd As Long, ByVal Msg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long


'Public Declare Function SetWindowLong Lib "user32" _
'   Alias "SetWindowLongA" _
'  (ByVal hwnd As Long, ByVal nIndex As Long, _
'   ByVal dwNewLong As Long) As Long

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
      

Public Const CL_LIGHTGREEN = 12648384
Public Const CL_PURPLE = 16384150
Public Const CL_DARKPURPLE = &H900090
Public Const CL_ORANGE = 25800
Public Const CL_LIGHTRED = 9869055
Public Const CL_DARKGREEN = &H8000&
Public Const CL_DARKBLUE = 13107200

'Windows API Constants
   Global Const CB_SHOWDROPDOWN = &H14F

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
'               Call WindowProc(hwnd, WM_LBUTTONDOWN, 1, 1000)
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
'           'call the default window handler
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
 

 




