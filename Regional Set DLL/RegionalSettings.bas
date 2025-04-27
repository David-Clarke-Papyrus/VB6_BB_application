Attribute VB_Name = "RegionalSettings"
Option Explicit



Public tempColl As Collection 'temporary collection object to load enumerated settings

Public Const RS_FAILED = 2777   'or for failing to change regional settings

Public Const LOCALE_SLANGUAGE As Long = &H2     'localized name of language
Public Const LOCALE_SSHORTDATE As Long = &H1F   'short date format string
Public Const LOCALE_SLONGDATE As Long = &H20    'long date format string
Public Const LOCALE_SDATE As Long = &H1D        'date separator


Public Const DATE_LONGDATE As Long = &H2
Public Const DATE_SHORTDATE As Long = &H1
Public Const LOCALE_IDATE As Long = &H21        'short date format ordering
Public Const LOCALE_ILDATE As Long = &H22       'long date format ordering

Public Const LOCALE_STIME As Long = &H1E        'time separator
Public Const LOCALE_STIMEFORMAT As Long = &H1003 'time format string


Public Const HWND_BROADCAST As Long = &HFFFF&
Public Const WM_SETTINGCHANGE As Long = &H1A
'***************************************************************************
'Stuff for Currency settings
Public Const LOCALE_SCURRENCY             As Long = &H14    'local monetary symbol
Public Const LOCALE_SINTLSYMBOL           As Long = &H15    'intl monetary symbol
Public Const LOCALE_SMONDECIMALSEP        As Long = &H16    'monetary decimal separator
Public Const LOCALE_SMONTHOUSANDSEP       As Long = &H17    'monetary thousand separator
Public Const LOCALE_SMONGROUPING          As Long = &H18    'monetary grouping
Public Const LOCALE_ICURRDIGITS           As Long = &H19    '# local monetary digits
Public Const LOCALE_IINTLCURRDIGITS       As Long = &H1A    '# intl monetary digits
Public Const LOCALE_ICURRENCY             As Long = &H1B    'positive currency mode
Public Const LOCALE_INEGCURR              As Long = &H1C    'negative currency mode
Public Const LOCALE_IPOSSIGNPOSN          As Long = &H52    'positive sign position
Public Const LOCALE_INEGSIGNPOSN          As Long = &H53    'negative sign position
Public Const LOCALE_IPOSSYMPRECEDES       As Long = &H54    'mon sym precedes pos amt
Public Const LOCALE_IPOSSEPBYSPACE        As Long = &H55    'mon sym sep by space from pos amt
Public Const LOCALE_INEGSYMPRECEDES       As Long = &H56    'mon sym precedes neg amt
Public Const LOCALE_INEGSEPBYSPACE        As Long = &H57    'mon sym sep by space from neg amt
Public Const LOCALE_SENGCURRNAME          As Long = &H1007  'english name of currency
Public Const LOCALE_SNATIVECURRNAME       As Long = &H1008  'native name of currency

'***************************************************************************
Public Declare Function PostMessage Lib "user32" _
   Alias "PostMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Public Declare Function EnumDateFormats Lib "kernel32" _
   Alias "EnumDateFormatsA" _
  (ByVal lpDateFmtEnumProc As Long, _
   ByVal Locale As Long, _
   ByVal dwFlags As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)

Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long


Public Declare Function GetLocaleInfo Lib "kernel32" _
   Alias "GetLocaleInfoA" _
  (ByVal Locale As Long, _
   ByVal LCType As Long, _
   ByVal lpLCData As String, _
   ByVal cchData As Long) As Long

Public Declare Function SetLocaleInfo Lib "kernel32" _
    Alias "SetLocaleInfoA" _
   (ByVal Locale As Long, _
    ByVal LCType As Long, _
    ByVal lpLCData As String) As Long
    
Public Declare Function GetThreadLocale Lib "kernel32" () As Long

Public Declare Function GetLastor Lib "kernel32" () As Long





Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, _
                                  ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim r As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
  'if successful..
   If r Then
    
     'pad the buffer with spaces
      sReturn = Space$(r)
       
     'and call again passing the buffer
      r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
     
     'if successful (r > 0)
      If r Then
      
        'r holds the size of the string
        'including the terminating null
         GetUserLocaleInfo = Left$(sReturn, r - 1)
      
      End If
   
   End If
    
End Function

Public Function EnumCalendarDateProc(lpDateFormatString As Long) As Long
    Dim sTemp As String
  'application-defined callback function for EnumDateFormats
  
  'populates combo assigned to global var thisCombo
  
   tempColl.Add CStr(StringFromPointer(lpDateFormatString))
  'return 1 to continue enumeration
   EnumCalendarDateProc = 1
   
End Function

Private Function StringFromPointer(lpString As Long) As String

   Dim pos As Long
   Dim buffer As String
   
  'pad a string to hold the data
   buffer = Space$(128)
   
  'copy the string pointed to by the return value
   CopyMemory ByVal buffer, lpString, ByVal Len(buffer)
   
  'remove the trailing null and trim
   pos = InStr(buffer, Chr$(0))
   
   If pos Then
      StringFromPointer = Left$(buffer, pos - 1)
   End If

End Function











