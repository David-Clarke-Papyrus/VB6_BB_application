VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinInet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private variables
Private mlConnectionNumber      As Long
Private mbDisconnectOnTerminate As Boolean

' Used to list the DUN connections
Private Type RAS_ENTRIES
    dwSize           As Long
    szEntryname(256) As Byte
End Type

Private Declare Function RasEnumEntriesA Lib "rasapi32.dll" _
    (ByVal reserved As String, ByVal lpszPhonebook As String, _
    lprasentryname As Any, lpcb As Long, lpcEntries As Long) As Long

' For the fActiveConnection function
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long
    
Private Declare Function RegOpenKey Lib "advapi32.dll" _
    Alias "RegOpenKeyA" (ByVal hKey As Long, _
    ByVal sSubKey As String, hKey As Long) As Long
    
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
    Alias "RegQueryValueExA" (ByVal hKey As Long, _
    ByVal sKeyValue As String, ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, nSizeData As Long) As Long

' For Dial and Hangup functions
Private Declare Function InternetDial Lib "wininet.dll" _
    (ByVal hwnd As Long, ByVal sConnectoid As String, _
    ByVal dwFlags As Long, lpdwConnection As Long, _
    ByVal dwReserved As Long) As Long
'Returns ERROR_SUCCESS if successfull
'        ERROR_INVALID_PARAMETER - one or more parameters are incorrect
'        ERROR_NO_CONNECTION - There is a problem with the dial-up connection
'        ERROR_USER_DISCONNECTION - The user clicked either the work offline or cancel button on the dialog box
    
    
Private Declare Function InternetHangUp Lib "wininet.dll" _
    (ByVal dwConnection As Long, ByVal dwReserved As Long) As Long
' Returns   ERROR_SUCCESS if successfull or an error value otherwise.
Private Declare Function InternetAutoDial Lib "wininet.dll" _
    (ByVal dwFlags As Long, ByVal hwnd As Long)
Private Declare Function InternetAutoDialHangUp Lib "wininet.dll" _
    (ByVal dwReserved As Long) As Long

' InternetAutodial flags
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = &H1
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = &H2
Private Const INTERNET_AUTODIAL_FAILIFSECURITYCHECK = &H4

' InternetDial Flags - must not conflict with InternetAutodial
'                      flags as they are valid here also.
Private Const INTERNET_DIAL_FORCE_PROMPT = &H2000
Private Const INTERNET_DIAL_SHOW_OFFLINE = &H4000
Private Const INTERNET_DIAL_UNATTENDED = &H8000

' Windows error constants.
Private Const ERROR_SUCCESS As Long = 0&
Private Const ERROR_INVALID_PARAMETER = 87&

' RAS error constants
Private Const RASBASE As Long = 600& 'not sure about this couldn't find raserror.h anywhere on MSDN so
                                     'best-guessed the value based on return code of 631 for cancel button
Private Const ERROR_NO_CONNECTION = (RASBASE + 68&)
Private Const ERROR_USER_DISCONNECTION = (RASBASE + 31&)


Private Declare Function InternetGetConnectedState Lib "wininet" _
  (ByRef dwFlags As Long, _
   ByVal dwReserved As Long) As Long

'Local system uses a modem to connect to the Internet.
Private Const INTERNET_CONNECTION_MODEM As Long = &H1

'Local system uses a LAN to connect to the Internet.
Private Const INTERNET_CONNECTION_LAN As Long = &H2

'Local system uses a proxy server to connect to the Internet.
Private Const INTERNET_CONNECTION_PROXY As Long = &H4

'No longer used.
Private Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8

Private Const INTERNET_RAS_INSTALLED As Long = &H10
Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
Private Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
'  Events for this module
Public Event ConnectionMade()
Public Event ConnectionClosed()


'InternetGetConnectedState wrapper functions
Public Function IsNetConnectViaLAN() As Boolean
    On Error GoTo errHandler

   Dim dwFlags As Long
   
  'pass an empty variable into which the API will
  'return the flags associated with the connection
   Call InternetGetConnectedState(dwFlags, 0&)

  'return True if the flags indicate a LAN connection
   IsNetConnectViaLAN = dwFlags And INTERNET_CONNECTION_LAN
     
    Exit Function
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WinInet.IsNetConnectViaLAN"
End Function


Public Function IsNetConnectViaModem() As Boolean
    On Error GoTo errHandler

   Dim dwFlags As Long
   
  'pass an empty variable into which the API will
  'return the flags associated with the connection
   Call InternetGetConnectedState(dwFlags, 0&)

  'return True if the flags indicate a modem connection
   IsNetConnectViaModem = dwFlags And INTERNET_CONNECTION_MODEM
     
    Exit Function
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WinInet.IsNetConnectViaModem"
End Function


Public Function IsNetConnectViaProxy() As Boolean
    On Error GoTo errHandler

   Dim dwFlags As Long
   
  'pass an empty variable into which the API will
  'return the flags associated with the connection
   Call InternetGetConnectedState(dwFlags, 0&)

  'return True if the flags indicate a proxy connection
   IsNetConnectViaProxy = dwFlags And INTERNET_CONNECTION_PROXY
     
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WinInet.IsNetConnectViaProxy"
End Function


Public Function IsNetConnectOnline(Optional pMsg As Variant) As Boolean
    On Error GoTo errHandler
Dim dwFlags As Long
Dim MSG As String

  'no flags needed here - the API returns True
  'if there is a connection of any type
   If InternetGetConnectedState(dwFlags, 0&) Then

      If dwFlags And INTERNET_CONNECTION_CONFIGURED Then
         MSG = MSG & "You have a network connection configured." & vbCrLf
      End If

      If dwFlags And INTERNET_CONNECTION_LAN Then
         MSG = MSG & "The local system connects to the Internet via a LAN"
      End If
      
      If dwFlags And INTERNET_CONNECTION_PROXY Then
            MSG = MSG & ", and uses a proxy server. "
      Else: MSG = MSG & "."
      End If
      
      If dwFlags And INTERNET_CONNECTION_MODEM Then
         MSG = MSG & "The local system uses a modem to connect to the Internet. "
      End If
      
      If dwFlags And INTERNET_CONNECTION_MODEM_BUSY Then
         MSG = MSG & "The local system's modem is busy with a non-Internet connection. "
      End If
      
      If dwFlags And INTERNET_RAS_INSTALLED Then
         MSG = MSG & "Remote Access Services are installed on this system."
      End If
      If dwFlags And INTERNET_CONNECTION_OFFLINE Then
         MSG = MSG & "The connection is currently offline. "
        IsNetConnectOnline = False
      Else
         MSG = MSG & "The connection is online. "
        IsNetConnectOnline = True
      End If
    
    End If
    
    If Not IsMissing(pMsg) Then
        pMsg = MSG
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WinInet.IsNetConnectOnline"
End Function


Public Function IsNetRASInstalled() As Boolean
    On Error GoTo errHandler

   Dim dwFlags As Long
   
  'pass an empty variable into which the API will
  'return the flags associated with the connection
   Call InternetGetConnectedState(dwFlags, 0&)

  'return True if the flags include RAS installed
   IsNetRASInstalled = dwFlags And INTERNET_RAS_INSTALLED
     
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WinInet.IsNetRASInstalled"
End Function


Public Function GetNetConnectString() As String
    On Error GoTo errHandler

   Dim dwFlags As Long
   Dim MSG As String

  'build a string for display
   If InternetGetConnectedState(dwFlags, 0&) Then
     
      If dwFlags And INTERNET_CONNECTION_CONFIGURED Then
         MSG = MSG & "You have a network connection configured." & vbCrLf
      End If

      If dwFlags And INTERNET_CONNECTION_LAN Then
         MSG = MSG & "The local system connects to the Internet via a LAN"
      End If
      
      If dwFlags And INTERNET_CONNECTION_PROXY Then
            MSG = MSG & ", and uses a proxy server. "
      Else: MSG = MSG & "."
      End If
      
      If dwFlags And INTERNET_CONNECTION_MODEM Then
         MSG = MSG & "The local system uses a modem to connect to the Internet. "
      End If
      
      If dwFlags And INTERNET_CONNECTION_OFFLINE Then
         MSG = MSG & "The connection is currently offline. "
      End If
      
      If dwFlags And INTERNET_CONNECTION_MODEM_BUSY Then
         MSG = MSG & "The local system's modem is busy with a non-Internet connection. "
      End If
      
      If dwFlags And INTERNET_RAS_INSTALLED Then
         MSG = MSG & "Remote Access Services are installed on this system."
      End If
      
   Else
    
      MSG = "Not connected to the internet now."
      
   End If
   
   GetNetConnectString = MSG

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WinInet.GetNetConnectString"
End Function
'--end block--'






Private Sub Class_Initialize()
    On Error GoTo errHandler

    mlConnectionNumber = 0&
    mbDisconnectOnTerminate = False
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WinInet.Class_Initialize"
End Sub
Private Sub Class_Terminate()
    On Error GoTo errHandler

    If mbDisconnectOnTerminate And mlConnectionNumber <> 0 Then
        Call InternetHangUp(mlConnectionNumber, 0&)
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WinInet.Class_Terminate"
End Sub
Public Property Get Connected() As Boolean
    On Error GoTo errHandler

    Connected = fActiveConnection()
    
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WinInet.Connected"
End Property
Public Property Get DisconnectOnTerminate() As Boolean
    On Error GoTo errHandler

    DisconnectOnTerminate = mbDisconnectOnTerminate
    
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WinInet.DisconnectOnTerminate"
End Property
Public Property Let DisconnectOnTerminate(ByVal bValue As Boolean)

    mbDisconnectOnTerminate = bValue
    
End Property
Public Function Hangup() As Long
    On Error GoTo errHandler

    If mlConnectionNumber = 0 Then
        '
        ' No connection from this module.
        '
        Hangup = -1
    Else
        Hangup = InternetHangUp(mlConnectionNumber, 0&)
        mlConnectionNumber = 0&
        
        RaiseEvent ConnectionClosed
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WinInet.HangUp"
End Function
Public Sub ListDUNs(sDunList() As String)
    On Error GoTo errHandler
Dim lngSize    As Long
Dim lngEntries As Long
Dim strConName As String
Dim lngIndex   As Long
Dim RAS(255)   As RAS_ENTRIES
    
    Erase sDunList()
    
    RAS(0).dwSize = 264
    lngSize = 256 * RAS(0).dwSize
    
    Call RasEnumEntriesA(vbNullString, vbNullString, RAS(0), lngSize, lngEntries)
    
    lngEntries = lngEntries - 1
    
    If lngEntries >= 0 Then
        ReDim sDunList(lngEntries)
        
        For lngIndex = 0 To lngEntries
            strConName = StrConv(RAS(lngIndex).szEntryname(), vbUnicode)
            sDunList(lngIndex) = Left$(strConName, InStr(strConName, vbNullChar) - 1)
        Next
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WinInet.ListDUNs(sDunList)", sDunList
End Sub
Public Function StartDUN(hwnd As Long, strDUN As String, Optional pUnattended As Boolean) As Long
    On Error GoTo errHandler
Dim lngResult As Long
    
        If Not IsMissing(pUnattended) Then
            If pUnattended Then
                lngResult = InternetDial(hwnd, strDUN, INTERNET_DIAL_UNATTENDED, mlConnectionNumber, 0)
'MsgBox "Attempted connection when connected: result = " & lngResult (Note to me result = 0)
            Else
                lngResult = InternetDial(hwnd, strDUN, 4, mlConnectionNumber, 0)
            End If
        End If
        
        If lngResult = ERROR_SUCCESS Then
            RaiseEvent ConnectionMade
        Else
            mlConnectionNumber = 0
        End If
        
        StartDUN = lngResult
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WinInet.StartDUN(hwnd,strDUN,pUnattended)", Array(hwnd, strDUN, pUnattended)
    
End Function

Private Function fActiveConnection() As Boolean
    On Error GoTo errHandler
Dim lngKey      As Long
Dim lngData     As Long
Dim lngSizeData As Long
   
Const sSubKey = "System\CurrentControlSet\Services\RemoteAccess"
Const sKeyValue = "Remote Connection"
    '
    ' Check registry for an active connection.
    '
    fActiveConnection = False
    
    If RegOpenKey(HKEY_LOCAL_MACHINE, sSubKey, lngKey) = ERROR_SUCCESS Then
        lngData = 0&
        lngSizeData = Len(lngData)
        
        If RegQueryValueEx(lngKey, sKeyValue, 0&, 0&, lngData, lngSizeData) = ERROR_SUCCESS Then
            fActiveConnection = lngData <> 0
        End If
        
        Call RegCloseKey(lngKey)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "WinInet.fActiveConnection"
End Function

