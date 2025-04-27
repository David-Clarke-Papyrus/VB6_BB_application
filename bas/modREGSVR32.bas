Attribute VB_Name = "modREGSVR32"
Option Explicit


'=============================================================================================================
'
' modREGSVR32 Module
' ------------------
'
' Created By  : Kevin Wilson
'               http://www.TheVBZone.com   ( The VB Zone )
'               http://www.TheVBZone.net   ( The VB Zone .net )
'
' Last Update : December 20, 2001
'
' VB Versions : 5.0 / 6.0
'
' Requires    : REGSVR32.EXE
'
' Description : This module makes it easy for you to programmatically register and unregister ActiveX components.
'
' NOTE        : This does not work for some reason on ActiveX .EXE's.  It should by all accounts.
'
'=============================================================================================================
'
' LEGAL:
'
' You are free to use this code as long as you keep the above heading information intact and unchanged. Credit
' given where credit is due.  Also, it is not required, but it would be appreciated if you would mention
' somewhere in your compiled program that that your program makes use of code written and distributed by
' Kevin Wilson (www.TheVBZone.com).  Feel free to link to this code via your web site or articles.
'
' You may NOT take this code and pass it off as your own.  You may NOT distribute this code on your own server
' or web site.  You may NOT take code created by Kevin Wilson (www.TheVBZone.com) and use it to create products,
' utilities, or applications that directly compete with products, utilities, and applications created by Kevin
' Wilson, TheVBZone.com, or Wilson Media.  You may NOT take this code and sell it for profit without first
' obtaining the written consent of the author Kevin Wilson.
'
' These conditions are subject to change at the discretion of the owner Kevin Wilson at any time without
' warning or notice.  Copyright© by Kevin Wilson.  All rights reserved.
'
'=============================================================================================================


' Constants
Private Const MAX_PATH = 260
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H2000 ' Specifies that the function should search the system message-table resource(s) for the requested message. If this flag is specified with FORMAT_MESSAGE_FROM_HMODULE, the function searches the system message table if the message is not found in the module specified by lpSource. Cannot be used with FORMAT_MESSAGE_FROM_STRING.  If this flag is specified, an application can pass the result of the GetLastError function to retrieve the message text for a system-defined error.

' Win32 API Declarations
Private Declare Sub SetLastError Lib "KERNEL32" (ByVal dwErrCode As Long)
Private Declare Function CallWindowProc Lib "USER32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FormatMessage Lib "KERNEL32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long
Private Declare Function FreeLibrary Lib "KERNEL32" (ByVal hLibrary As Long) As Long 'BOOL
Private Declare Function GetLastError Lib "KERNEL32" () As Long
Private Declare Function GetProcAddress Lib "KERNEL32" (ByVal hLibrary As Long, ByVal strFunctionName As String) As Long
Private Declare Function LoadLibrary Lib "KERNEL32" Alias "LoadLibraryA" (ByVal strFileName As String) As Long


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


Public Function RegisterCom(ByVal strFileName As String, _
                            Optional ByVal blnSilent As Boolean = True, _
                            Optional ByRef Return_ErrNum As Long, _
                            Optional ByRef Return_ErrDesc As String) As Boolean
    On Error GoTo errHandler
  
  ' Clear returns
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' Validate parameters
  strFileName = Trim(strFileName)
  If Dir(strFileName, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then
    Return_ErrNum = -1
    Return_ErrDesc = "The file specified '" & strFileName & "' could not be found"
    Exit Function
  End If
  
  ' Run REGSVR32.EXE to register the COM object
  If blnSilent = True Then
    Shell "REGSVR32.EXE " & strFileName, vbNormalFocus
  Else
    Shell "REGSVR32.EXE /S " & strFileName, vbHide
  End If
  
  RegisterCom = True
  Exit Function
  
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "modREGSVR32.RegisterCom(strFileName,blnSilent,Return_ErrNum,Return_ErrDesc)", _
         Array(strFileName, blnSilent, Return_ErrNum, Return_ErrDesc)
End Function

Public Function UnregisterCom(ByVal strFileName As String, _
                              Optional ByVal blnSilent As Boolean = True, _
                              Optional ByRef Return_ErrNum As Long, _
                              Optional ByRef Return_ErrDesc As String) As Boolean
    On Error GoTo errHandler
  
  ' Clear returns
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' Validate parameters
  strFileName = Trim(strFileName)
  If Dir(strFileName, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then
    Return_ErrNum = -1
    Return_ErrDesc = "The file specified '" & strFileName & "' could not be found"
    Exit Function
  End If
  
  ' Run REGSVR32.EXE to unregister the COM object
  If blnSilent = True Then
    Shell "REGSVR32.EXE /U " & strFileName, vbNormalFocus
  Else
    Shell "REGSVR32.EXE /U /S " & strFileName, vbHide
  End If
  
  UnregisterCom = True
  Exit Function
  
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "modREGSVR32.UnregisterCom(strFileName,blnSilent,Return_ErrNum,Return_ErrDesc)", _
         Array(strFileName, blnSilent, Return_ErrNum, Return_ErrDesc)
End Function

Public Function RegisterComEx(ByVal strFileName As String, _
                              Optional ByRef Return_ErrNum As Long, _
                              Optional ByRef Return_ErrDesc As String) As Boolean
    On Error GoTo errHandler
  
  Dim hLibrary  As Long
  Dim hFunction As Long
  
  ' Clear returns
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' Validate parameters
  strFileName = Trim(strFileName)
  If Dir(strFileName, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then
    Return_ErrNum = -1
    Return_ErrDesc = "The file specified '" & strFileName & "' could not be found"
    Exit Function
  End If
  If Right(strFileName, 1) <> Chr(0) Then strFileName = strFileName & Chr(0)
  
  ' Load the COM object using the LoadLibrary function
  hLibrary = LoadLibrary(strFileName)
  If hLibrary = 0 Then
    GetLastErr_Msg Err.LastDllError, , Return_ErrNum, Return_ErrDesc, False
    Exit Function
  End If
  
  ' Get the handle to the function to call
  hFunction = GetProcAddress(hLibrary, "DllRegisterServer" & Chr(0))
  If hFunction = 0 Then
    GetLastErr_Msg Err.LastDllError, , Return_ErrNum, Return_ErrDesc, False
    GoTo CleanUp
  End If
  
  ' Call the function
  If CallWindowProc(hFunction, 0, 0, 0, 0) = 0 Then
    RegisterComEx = True
  Else
    GetLastErr_Msg Err.LastDllError, , Return_ErrNum, Return_ErrDesc, False
  End If
  
CleanUp:
  
  If hLibrary <> 0 Then FreeLibrary hLibrary
  
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "modREGSVR32.RegisterComEx(strFileName,Return_ErrNum,Return_ErrDesc)", Array(strFileName, _
         Return_ErrNum, Return_ErrDesc)
End Function

Public Function UnregisterComEx(ByVal strFileName As String, _
                                Optional ByRef Return_ErrNum As Long, _
                                Optional ByRef Return_ErrDesc As String) As Boolean
    On Error GoTo errHandler
  
  Dim hLibrary  As Long
  Dim hFunction As Long
  
  ' Clear returns
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' Validate parameters
  strFileName = Trim(strFileName)
  If Dir(strFileName, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then
    Return_ErrNum = -1
    Return_ErrDesc = "The file specified '" & strFileName & "' could not be found"
    Exit Function
  End If
  If Right(strFileName, 1) <> Chr(0) Then strFileName = strFileName & Chr(0)
  
  ' Load the COM object using the LoadLibrary function
  hLibrary = LoadLibrary(strFileName)
  If hLibrary = 0 Then
    GetLastErr_Msg Err.LastDllError, , Return_ErrNum, Return_ErrDesc, False
    Exit Function
  End If
  
  ' Get the handle to the function to call
  hFunction = GetProcAddress(hLibrary, "DllUnregisterServer" & Chr(0))
  If hFunction = 0 Then
    GetLastErr_Msg Err.LastDllError, , Return_ErrNum, Return_ErrDesc, False
    GoTo CleanUp
  End If
  
  ' Call the function
  If CallWindowProc(hFunction, 0, 0, 0, 0) = 0 Then
    UnregisterComEx = True
  Else
    GetLastErr_Msg Err.LastDllError, , Return_ErrNum, Return_ErrDesc, False
  End If
  
CleanUp:
  
  If hLibrary <> 0 Then FreeLibrary hLibrary
  
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "modREGSVR32.UnregisterComEx(strFileName,Return_ErrNum,Return_ErrDesc)", _
         Array(strFileName, Return_ErrNum, Return_ErrDesc)
End Function


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


Private Function GetLastErr_Msg(Optional ByVal ErrorNumber As Long, _
                               Optional ByVal LastAPICalled As String = "last", _
                               Optional ByRef Return_ErrNum As Long, _
                               Optional ByRef Return_ErrDesc As String, _
                               Optional ByVal ShowErrors As Boolean = False) As Boolean
    On Error GoTo errHandler
  
  ' Clear the return values first
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' If no error message is specified then check for one
  If ErrorNumber = 0 Then
    ErrorNumber = GetLastError
    If ErrorNumber = 0 Then
      GetLastErr_Msg = False
      Exit Function
    End If
  End If
  
  ' Allocate a buffer for the error description
  Return_ErrDesc = String(MAX_PATH, 0)
  
  ' Get the error description
  FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, ErrorNumber, 0, Return_ErrDesc, MAX_PATH, 0
  Return_ErrNum = ErrorNumber
  Return_ErrDesc = Left(Return_ErrDesc, InStr(Return_ErrDesc, Chr(0)) - 1)
  If Right(Return_ErrDesc, Len(vbCrLf)) = vbCrLf Then
    Return_ErrDesc = Left(Return_ErrDesc, Len(Return_ErrDesc) - Len(vbCrLf))
  End If
  
  ' Display the error message
  If ShowErrors = True Then
    MsgBox "An error occured while calling the " & LastAPICalled & " Windows API function." & Chr(13) & "Below is the error information:" & Chr(13) & Chr(13) & "Error Number = " & CStr(ErrorNumber) & Chr(13) & "Error Description = " & Return_ErrDesc, vbOKOnly + vbExclamation, "  Windows API Error"
  End If
  GetLastErr_Msg = True
  
  ' Set the last error to 0 (no error) so next time through it doesn't report the same error twice
  SetLastError 0
  
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "modREGSVR32.GetLastErr_Msg(ErrorNumber,LastAPICalled,Return_ErrNum,Return_ErrDesc," & _
        "ShowErrors)", Array(ErrorNumber, LastAPICalled, Return_ErrNum, Return_ErrDesc, ShowErrors)
End Function

