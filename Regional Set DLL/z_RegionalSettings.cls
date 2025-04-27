VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_RegionalSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Date Info
Private ShortDateColl As Collection
Private LongDateColl As Collection

Private sTimeMask As String
Private sTimeSepSymbol As String
Private sShortDate As String
Private sDateSep As String
Private sLongDate As String
Private sDateSepSymbol As String
Private sLanguage As String
   'Currency Info
Private sCurrSymbol As String
Private sCurrInternatSymbol As String
Private sCurrDecSep As String
Private sCurrThousandSep As String
Private sCurrGroup As String
Private sCurrDigits As String
Private sCurrInternatDigits As String
Private sCurrPosSymbol As String
Private sCurrNegSymbol As String
Private sCurrNegSign As String
Private sCurrPosSign As String
Private sCurrSymbPosPrec As String
Private sCurrSymbNegPrec As String
Private sCurrPosSep As String
Private sCurrNegSep As String
Private sCurrWin2KEngName As String
Private sCurrWin2KLocName As String


Private Sub Class_Initialize()

    Set ShortDateColl = New Collection
    Set LongDateColl = New Collection
    GetDateSettings
    GetCurrencySettings
    GetTimeSettings
End Sub
Private Sub Class_Terminate()
    Set ShortDateColl = Nothing
    Set LongDateColl = Nothing
End Sub


Public Property Let SetShortDate(newFormat As String)
    Dim LCID As Long
    Dim r As Long
    Dim msg As String
    LCID = GetUserDefaultLCID()
   
    If newFormat <> "" Then
     'newFormat = newFormat & Chr(0)
     'set the new short date format
      r = SetLocaleInfo(LCID, LOCALE_SSHORTDATE, newFormat)

     'send a system notification message that a change was made
     'If r Then
     '   r = PostMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0&, ByVal 0&)
     'Else
     '   msg = GetLastor
     '   GoTo err_Handler
     'End If
     
     'update the textbox and label
      sShortDate = GetUserLocaleInfo(LCID, LOCALE_SSHORTDATE)
   
     'create new temp Text list
      Set tempColl = New Collection
      'thisCombo.Clear
       
     'enumerate new short date formats
      r = EnumDateFormats(AddressOf EnumCalendarDateProc, LCID, DATE_SHORTDATE)
     If r = 0 Then
        msg = GetLastor
        GoTo ERR_Handler
     End If
     'assign temp Text list to ShortDate list
     Set ShortDateColl = tempColl
     Set tempColl = Nothing
        
   End If
   Exit Property
ERR_Handler:
   'MsgBox "Can't set new regional short date settings!" & vbLf & _
          "or: " & msg, vbOKOnly + vbCritical, "or"
    Err.Raise vbObjectError + RS_FAILED
End Property

Public Property Let SetLongDate(newFormat As String)
    Dim LCID As Long
    Dim r As Long
   LCID = GetUserDefaultLCID()
   
   If newFormat <> "" Then
   
     'set the new long date format
      r = SetLocaleInfo(LCID, LOCALE_SLONGDATE, newFormat)
      If r = 0 Then GoTo ERR_Handler
     'send a system notification message that a change was made
      'Call PostMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0&, ByVal 0&)
     'update the textbox and label
      sLongDate = GetUserLocaleInfo(LCID, LOCALE_SLONGDATE)
      
     'create new temp Text list
      Set tempColl = New Collection
      'thisCombo.Clear
       
     'enumerate new long date formats
      r = EnumDateFormats(AddressOf EnumCalendarDateProc, LCID, DATE_LONGDATE)
      If r = 0 Then GoTo ERR_Handler
     'assign temp Text list to ShortDate list
     Set LongDateColl = tempColl
    
   
   End If
My_Exit:
    Set tempColl = Nothing
   Exit Property
ERR_Handler:
   Err.Raise vbObjectError + RS_FAILED
   
End Property

Private Sub GetDateSettings()
    
   Dim LCID As Long
   
   LCID = GetUserDefaultLCID()
   
  'show localized name of language
   sLanguage = GetUserLocaleInfo(LCID, LOCALE_SLANGUAGE)
   
  '-------------------------
   Set tempColl = New Collection
  'enumerate available long date formats
   Call EnumDateFormats(AddressOf EnumCalendarDateProc, LCID, DATE_LONGDATE)
   
   Set LongDateColl = tempColl
   Set tempColl = Nothing
   
  'Show the user's Long date format string
   sLongDate = GetUserLocaleInfo(LCID, LOCALE_SLONGDATE)
   
  '-------------------------
  'enumerate available short date formats
   Set tempColl = New Collection
   Call EnumDateFormats(AddressOf EnumCalendarDateProc, LCID, DATE_SHORTDATE)
   Set ShortDateColl = tempColl
   Set tempColl = Nothing
   
  'Show the user's Short date format string
   sShortDate = GetUserLocaleInfo(LCID, LOCALE_SSHORTDATE)

  'Show the user's Short date separetor
   sDateSepSymbol = GetUserLocaleInfo(LCID, LOCALE_SDATE)
  
End Sub

Private Sub GetCurrencySettings()
    Dim LCID As Long
   
   LCID = GetUserDefaultLCID()

  'LOCALE_SCURRENCY
  'String used as the local monetary symbol. The
  'maximum number of characters allowed is six.
   sCurrSymbol = GetUserLocaleInfo(LCID, LOCALE_SCURRENCY)

  'LOCALE_SINTLSYMBOL
  'Three characters of the international monetary symbol
  'specified in ISO 4217 followed by the character
  'separating this string from the amount.
   sCurrInternatSymbol = GetUserLocaleInfo(LCID, LOCALE_SINTLSYMBOL)

  'LOCALE_SMONDECIMALSEP
  'Character(s) used as the monetary decimal separator.
  'The maximum characters allowed is four.
   sCurrDecSep = GetUserLocaleInfo(LCID, LOCALE_SMONDECIMALSEP)

  'LOCALE_SMONTHOUSANDSEP
  'Character(s) used as the monetary separator between groups
  'of digits to the left of the decimal. The maximum number
  'of characters allowed is four.
   sCurrThousandSep = GetUserLocaleInfo(LCID, LOCALE_SMONTHOUSANDSEP)
   
  'LOCALE_SMONGROUPING
  'Sizes for each group of monetary digits to the left of
  'the decimal. An explicit size is needed for each group,
  'and sizes are separated by semicolons. If the last value
  'is zero, the preceding value is repeated. For example,
  'to group thousands, specify 3;0. Indic languages group
  'the first thousand and then group by hundreds - for
  'example, 12,34,56,789, which is represented by 3;2;0.
  'The maximum characters allowed is four.
   sCurrGroup = GetUserLocaleInfo(LCID, LOCALE_SMONGROUPING)

  'LOCALE_ICURRDIGITS
  'Number of fractional digits for the local monetary format.
  'The maximum characters allowed is three.
   sCurrDigits = GetUserLocaleInfo(LCID, LOCALE_ICURRDIGITS)

  'LOCALE_IINTLCURRDIGITS
  'Number of fractional digits for the international monetary
  'format. The maximum characters allowed is three.
   sCurrInternatDigits = GetUserLocaleInfo(LCID, LOCALE_IINTLCURRDIGITS)

  'LOCALE_ICURRENCY
  'Position of the monetary symbol in the positive currency mode.
  'The maximum characters allowed is two. The mode can be
  'one of the following values:
   Select Case GetUserLocaleInfo(LCID, LOCALE_ICURRENCY)
      Case "0": sCurrPosSymbol = "0 - Prefix, no separation, ie $1.1"
      Case "1": sCurrPosSymbol = "1 - Suffix, no separation, ie 1.1$"
      Case "2": sCurrPosSymbol = "2 - Prefix, 1-character separation, ie $ 1.1"
      Case "3": sCurrPosSymbol = "3 - Suffix, 1-character separation, ie 1.1 $"
   End Select

  'LOCALE_INEGCURR
  'Negative currency mode. The maximum characters allowed is
  'three. The mode can be one of the following values:
   Select Case GetUserLocaleInfo(LCID, LOCALE_INEGCURR)
      Case "0": sCurrNegSymbol = "0 - Left parenthesis,monetary symbol,number,right parenthesis, ie ($1.1)"
      Case "1": sCurrNegSymbol = "1 - Negative sign, monetary symbol, number, ie -$1.1"
      Case "2": sCurrNegSymbol = "2 - Monetary symbol, negative sign, number, ie $-1.1"
      Case "3": sCurrNegSymbol = "3 - Monetary symbol, number, negative sign, ie $1.1-"
      Case "4": sCurrNegSymbol = "4 - Left parenthesis, number, monetary symbol, right parenthesis, ie (1.1$)"
      Case "5": sCurrNegSymbol = "5 - Negative sign, number, monetary symbol, ie -1.1$"
      Case "6": sCurrNegSymbol = "6 - Number, negative sign, monetary symbol, ie 1.1-$"
      Case "7": sCurrNegSymbol = "7 - Number, monetary symbol, negative sign, ie 1.1$-"
      Case "8": sCurrNegSymbol = "8 - Negative sign, number, space, monetary symbol (like #5, but w/ space before monetary symbol), ie -1.1 $"
      Case "9": sCurrNegSymbol = "9 - Negative sign, monetary symbol, space, number (like #1, but w/ space after monetary symbol), ie -$ 1.1"
      Case "10": sCurrNegSymbol = "10 - Number, space, monetary symbol, negative sign (like #7, but w/ space before monetary symbol), ie 1.1 $-"
      Case "11": sCurrNegSymbol = "11 - Monetary symbol, space, number, negative sign (like #3, but w/ space after monetary symbol), ie $ 1.1-"
      Case "12": sCurrNegSymbol = "12 - Monetary symbol, space, negative sign, number (like #2, but w/ space after monetary symbol), ie $ -1.1"
      Case "13": sCurrNegSymbol = "13 - Number, negative sign, space, monetary symbol (like #6, but w/ space before monetary symbol), ie 1.1- $"
      Case "14": sCurrNegSymbol = "14 - Left parenthesis, monetary symbol, space, number, right parenthesis (like #0, but w/ space after monetary symbol), ie ($ 1.1)"
      Case "15": sCurrNegSymbol = "15 - Left parenthesis, number, space, monetary symbol, right parenthesis (like #4, but w/ space before monetary symbol), ie (1.1 $)"
   End Select
   
   
  'LOCALE_INEGSIGNPOSN
  'Formatting index for the negative sign in currency
  'values. The maximum allowed is two. The index can
  'be one of the following values.
   Select Case GetUserLocaleInfo(LCID, LOCALE_INEGSIGNPOSN)
      Case "0": sCurrNegSign = "0 - Parentheses surround the amount and the monetary symbol."
      Case "1": sCurrNegSign = "1 - The sign precedes the number."
      Case "2": sCurrNegSign = "2 - The sign follows the number."
      Case "3": sCurrNegSign = "3 - The sign precedes the monetary symbol."
      Case "4": sCurrNegSign = "4 - The sign follows the monetary symbol"
   End Select
   
  'LOCALE_IPOSSIGNPOSN
  'Formatting index for positive values. This index
  'uses the same values as LOCALE_INEGSIGNPOSN, except
  'that it does not use the zero index. The maximum
  'number of characters allowed for this string is two.
   Select Case GetUserLocaleInfo(LCID, LOCALE_IPOSSIGNPOSN)
      Case "0": sCurrPosSign = "0 - Parentheses surround the amount and the monetary symbol."
      Case "1": sCurrPosSign = "1 - The sign precedes the number."
      Case "2": sCurrPosSign = "2 - The sign follows the number."
      Case "3": sCurrPosSign = "3 - The sign precedes the monetary symbol."
      Case "4": sCurrPosSign = "4 - The sign follows the monetary symbol"
   End Select
   
  'LOCALE_IPOSSYMPRECEDES
  'Position of monetary symbol in a positive monetary value.
  'This value is 1 if the monetary symbol precedes the positive
  'amount, zero if it follows it. The maximum allowed is two.
   Select Case GetUserLocaleInfo(LCID, LOCALE_IPOSSYMPRECEDES)
      Case "0": sCurrSymbPosPrec = "0 - Monetary symbol follows a positive monetary amount."
      Case "1": sCurrSymbPosPrec = "1 - Monetary symbol precedes a positive monetary amount."
   End Select
   
  'LOCALE_INEGSYMPRECEDES
  'Position of monetary symbol in a negative monetary value.
  'This value is 1 if the monetary symbol precedes the negative
  'amount, 0 if it follows it. The maximum allowed is two.
   Select Case GetUserLocaleInfo(LCID, LOCALE_INEGSYMPRECEDES)
      Case "0": sCurrSymbNegPrec = "0 - Monetary symbol follows a negative monetary amount."
      Case "1": sCurrSymbNegPrec = "1 - Monetary symbol precedes a negative monetary amount."
   End Select
   
  'LOCALE_IPOSSEPBYSPACE
  'Separation of monetary symbol in a positive monetary value.
  'This value is 1 if the monetary symbol is separated by a
  'space from a positive amount, 0 if not. The maximum
  'number allowed is two.
   Select Case GetUserLocaleInfo(LCID, LOCALE_IPOSSEPBYSPACE)
      Case "0": sCurrPosSep = "0 - Monetary symbol has no space preceeding a positive amount."
      Case "1": sCurrPosSep = "1 - Monetary symbol has space preceeding a positive amount."
   End Select
   
  'LOCALE_INEGSEPBYSPACE
  'Separation of the negative sign in a monetary value.
  'This value is 1 if the monetary symbol is separated
  'by a space from the negative amount, 0 if it is not.
  'The maximum allowed is two.
   Select Case GetUserLocaleInfo(LCID, LOCALE_INEGSEPBYSPACE)
      Case "0": sCurrNegSep = "0 - Monetary symbol has no space preceeding a negative amount."
      Case "1": sCurrNegSep = "1 - Monetary symbol has space preceeding a negative amount."
   End Select
   
  'LOCALE_SENGCURRNAME
  'Windows 2000: The full English name of the currency
  'associated with the locale.
   sCurrWin2KEngName = GetUserLocaleInfo(LCID, LOCALE_SENGCURRNAME)
   
  'LOCALE_SNATIVECURRNAME
  'Windows 2000: The native name of the currency
  'associated with the locale.
   sCurrWin2KLocName = GetUserLocaleInfo(LCID, LOCALE_SNATIVECURRNAME)
   
   If sCurrWin2KEngName = "" Then sCurrWin2KEngName = "Sorry, Windows 2000 only"
   If sCurrWin2KLocName = "" Then sCurrWin2KLocName = "Sorry, Windows 2000 only"

End Sub

Public Property Get ShortDates() As Collection
    Set ShortDates = ShortDateColl
End Property
Public Property Get LongDates() As Collection
    Set LongDates = LongDateColl
End Property

Public Property Get ShortDate() As String
    ShortDate = Trim(sShortDate)
End Property

Public Property Get DateSepSymbol() As String
    DateSepSymbol = Trim(sDateSepSymbol)
End Property

Public Property Get LongDate() As String
    LongDate = Trim(sLongDate)
End Property

Public Property Get CurrencySymbol() As String
    CurrencySymbol = sCurrSymbol
End Property
Public Property Let SetCurrencySymbol(newFormat As String)
    Dim LCID As Long
   
    LCID = GetUserDefaultLCID()
   
    If Len(newFormat) = 1 Then
   
     'set the new currency symbol
      Call SetLocaleInfo(LCID, LOCALE_SCURRENCY, newFormat)

     'send a system notification message that a change was made
      'Call PostMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0&, ByVal 0&)
            
     'update the textbox and label
      sCurrSymbol = GetUserLocaleInfo(LCID, LOCALE_SCURRENCY)
   
        
   End If
End Property

Public Property Let SetDateSepSymbol(newFormat As String)
    Dim LCID As Long
    Dim r As Long
   
    LCID = GetUserDefaultLCID()
   
    If Len(newFormat) = 1 Then
   
     'set the new currency symbol
      r = SetLocaleInfo(LCID, LOCALE_SDATE, newFormat)

     'send a system notification message that a change was made
      'Call PostMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0&, ByVal 0&)
            
     'update the textbox and label
      sDateSepSymbol = GetUserLocaleInfo(LCID, LOCALE_SDATE)
   
        
   End If
End Property

Public Property Let SetTimeSepSymbol(newFormat As String)
    Dim LCID As Long
    Dim r As Long
   
    LCID = GetUserDefaultLCID()
   
    If Len(newFormat) = 1 Then
   
     'set the new currency symbol
      r = SetLocaleInfo(LCID, LOCALE_STIME, newFormat)

     'send a system notification message that a change was made
      'Call PostMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0&, ByVal 0&)
            
     'update the textbox and label
      sTimeSepSymbol = GetUserLocaleInfo(LCID, LOCALE_STIME)
   
        
   End If
End Property

Public Property Let SetTimeMask(newFormat As String)
    Dim LCID As Long
    Dim r As Long
   
    LCID = GetUserDefaultLCID()
   
    If Len(newFormat) > 0 Then
   
     'set the new time mask
      r = SetLocaleInfo(LCID, LOCALE_STIMEFORMAT, newFormat)

     'send a system notification message that a change was made
      'Call PostMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0&, ByVal 0&)
            
     'update the textbox and label
      sTimeMask = GetUserLocaleInfo(LCID, LOCALE_STIMEFORMAT)
   
     If r = 0 Then Err.Raise vbObjectError + RS_FAILED
   End If
End Property

Private Sub GetTimeSettings()
    
   Dim LCID As Long
   
   LCID = GetUserDefaultLCID()
   
  
  'Show the user's time format string
   sTimeMask = GetUserLocaleInfo(LCID, LOCALE_STIMEFORMAT)
  

  'Show the user's time separetor
   sTimeSepSymbol = GetUserLocaleInfo(LCID, LOCALE_STIME)
  
End Sub



Public Property Get TimeMask() As String
    TimeMask = Trim(sTimeMask)
End Property
Public Property Get TimeSepSymbol() As String
    TimeSepSymbol = Trim(sTimeSepSymbol)
End Property

