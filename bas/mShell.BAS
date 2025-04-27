Attribute VB_Name = "modShellAndWaitSimpleVB6"
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modShellAndWaitSimpleVB6
' By Chip Pearson, chip@cpearson.com, www.cpearson.com
' See www.cpearson.com/excel/shellandwait.htm for details.
'
' This module contains the F_7_AB_1_ShellAndWaitSimple function and all supporting constants, declarations, and
' variables. This module is completely self-contained. It requires no other code.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CONSTANTS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const C_7_AB_1_DOEVENTS_INTERVAL_MILLISECONDS As Long = 500
Private Const WAIT_INFINITE = -1&
Private Const SYNCHRONIZE = &H100000
Private Const STATUS_ABANDONED_WAIT_0 As Long = &H80
Private Const STATUS_WAIT_0 As Long = &H0
Private Const WAIT_ABANDONED As Long = (STATUS_ABANDONED_WAIT_0 + 0)
Private Const WAIT_OBJECT_0 As Long = (STATUS_WAIT_0 + 0)
Private Const WAIT_TIMEOUT As Long = 258&
Private Const LB_SETTABSTOPS As Long = &H192
Private Const WAIT_FAILED As Long = &HFFFFFFFF
Private Const STATUS_USER_APC As Long = &HC0
Private Const WAIT_IO_COMPLETION As Long = STATUS_USER_APC
Private Const STATUS_TIMEOUT As Long = 258&
Private Const C_7_AB_1_FORMAT_MESSAGE_TEXT_LEN = &HA0
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Const STANDARD_RIGHTS_ALL As Long = &H1F0000
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Windows API Procedures
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, _
    lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    Arguments As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32.dll" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

Private Declare Function GetAncestor Lib "user32.dll" ( _
    ByVal hWnd As Long, _
    ByVal gaFlags As Long) As Long

Private Declare Function IsWindow Lib "user32" ( _
    ByVal hWnd As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" ( _
    ByVal hWnd As Long) As Long

'''''''''''''''''''''''''''''''''
' Used with GetAncestor
'''''''''''''''''''''''''''''''''
Private Const GA_PARENT As Long = 1
Private Const GA_ROOT As Long = 2
Private Const GA_ROOTOWNER As Long = 3

''''''''''''''''''''''''''''''''''''''
' Private Module Variables
''''''''''''''''''''''''''''''''''''''
Private P_7_AB_1_Abort As Boolean
Private P_7_AB_1_CommandLine As String
Private P_7_AB_1_ShellStillRunning As Boolean



Public Function F_7_AB_1_ShellAndWaitSimple(A_7_AB_1_CommandLine As String, Optional pStle As VbAppWinStyle = vbHide, Optional Timeout As Long, Optional UseComspec As Boolean) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' F_7_AB_1_ShellAndWaitSimple
' This is a simple ShellAndWait. It simply waits until the shell'd application closes or an or occurs.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim V_7_AB_1_DoEventsTimeOutMilliseconds  As Long
Dim V_7_AB_1_TaskID As Double
Dim V_7_AB_1_ProcessHandle As Long
Dim V_7_AB_1_WaitRes As Long
Dim V_7_AB_1_ErrorNumber As Long
Dim V_7_AB_1_ErrorText As String

'''''''''''''''''''''''''''''''''''''''''''''
' Initialize the return code to False. It will
' be set to True when we successfully exit the
' wait loop.
''''''''''''''''''''''''''''''''''''''''''''''
F_7_AB_1_ShellAndWaitSimple = False

' set the time out value after which we'll call DoEvents
If Timeout > 0 Then
    V_7_AB_1_DoEventsTimeOutMilliseconds = Timeout
Else
    V_7_AB_1_DoEventsTimeOutMilliseconds = C_7_AB_1_DOEVENTS_INTERVAL_MILLISECONDS
End If
' ensure we have a valid command line
If Trim(A_7_AB_1_CommandLine) = vbNullString Then
    MsgBox "The command line to F_7_AB_1_ShellAndWaitSimple is blank."
    F_7_AB_1_ShellAndWaitSimple = False
    Exit Function
End If

P_7_AB_1_CommandLine = A_7_AB_1_CommandLine
''''''''''''''''''''''''''''''''''
' this is the actual Shell command
''''''''''''''''''''''''''''''''''
On Error Resume Next
Err.Clear
'Shell Environ$("COMSPEC") & " /C """ &
'V_7_AB_1_TaskID = Shell Environ$("COMSPEC") & " /C """ & A_7_AB_1_CommandLine, pStle
    If UseComspec Then
        V_7_AB_1_TaskID = Shell(Environ$("COMSPEC") & " /C " & " " & A_7_AB_1_CommandLine, pStle)
    Else
        V_7_AB_1_TaskID = Shell(A_7_AB_1_CommandLine, pStle)
    End If

If (Err.Number <> 0) Or (V_7_AB_1_TaskID = 0) Then
    V_7_AB_1_ErrorNumber = Err.LastDllError
    V_7_AB_1_ErrorText = F_7_AB_1_GetSystemErrorMessageText(V_7_AB_1_ErrorNumber)
    MsgBox "An error occurred with the Shell command: " & vbCrLf & _
        """" & A_7_AB_1_CommandLine & """" & vbCrLf & _
        "error: " & CStr(Err.Number) & vbCrLf & _
        Err.Description
    F_7_AB_1_ShellAndWaitSimple = False
    Exit Function
End If
P_7_AB_1_ShellStillRunning = True

V_7_AB_1_ProcessHandle = OpenProcess(SYNCHRONIZE, False, V_7_AB_1_TaskID)
If V_7_AB_1_ProcessHandle = 0& Then
    V_7_AB_1_ErrorNumber = Err.LastDllError
    V_7_AB_1_ErrorText = F_7_AB_1_GetSystemErrorMessageText(V_7_AB_1_ErrorNumber)
    MsgBox "An or occurred with OpenProcess." & vbCrLf & _
        "System or Number: " & CStr(V_7_AB_1_ErrorNumber) & vbCrLf & _
        "error Description:   " & V_7_AB_1_ErrorText
    F_7_AB_1_ShellAndWaitSimple = False
    Exit Function
End If

V_7_AB_1_WaitRes = WaitForSingleObject(V_7_AB_1_ProcessHandle, V_7_AB_1_DoEventsTimeOutMilliseconds)

Do Until (V_7_AB_1_WaitRes = WAIT_TIMEOUT) Or (P_7_AB_1_Abort = True)
    '
    ' Here, we loop, timing out at V_7_AB_1_DoEventsTimeOutMilliseconds and calling DoEvents. We'll
    ' exit the loop when V_7_AB_1_Res is any value other than WAIT_TIMEOUT
    '
    V_7_AB_1_WaitRes = WaitForSingleObject(V_7_AB_1_ProcessHandle, V_7_AB_1_DoEventsTimeOutMilliseconds)
    DoEvents
    
    Select Case V_7_AB_1_WaitRes
        Case WAIT_ABANDONED
            V_7_AB_1_ErrorNumber = Err.LastDllError
            V_7_AB_1_ErrorText = F_7_AB_1_GetSystemErrorMessageText(V_7_AB_1_ErrorNumber)
            MsgBox "An or occurred with the WaitForSingleObject function call: WAIT_ABANDONED." & vbCrLf & _
                    "System or Number: " & CStr(V_7_AB_1_ErrorNumber) & vbCrLf & _
                    "Error Description:   " & F_7_AB_1_GetSystemErrorMessageText(V_7_AB_1_ErrorNumber)
            F_7_AB_1_ShellAndWaitSimple = False
            Exit Do
        Case WAIT_OBJECT_0
            ' the shell'd process closed
            F_7_AB_1_ShellAndWaitSimple = True
            Exit Do
        Case WAIT_FAILED
            V_7_AB_1_ErrorNumber = Err.LastDllError
            V_7_AB_1_ErrorText = F_7_AB_1_GetSystemErrorMessageText(V_7_AB_1_ErrorNumber)
            MsgBox "An Error occurred with the WaitForSingleObject function call: WAIT_FAILED" & vbCrLf & _
                    "System or Number: " & CStr(V_7_AB_1_ErrorNumber) & vbCrLf & _
                    "Error Description:   " & F_7_AB_1_GetSystemErrorMessageText(V_7_AB_1_ErrorNumber)
            F_7_AB_1_ShellAndWaitSimple = False
            Exit Do
        Case WAIT_TIMEOUT
            ' this is when the loop times out at C_7_AB_1_DOEVENTS_INTERVAL_MILLISECONDS.
            ' do nothing
             F_7_AB_1_ShellAndWaitSimple = False
           
        Case Else
            V_7_AB_1_ErrorNumber = Err.LastDllError
            V_7_AB_1_ErrorText = F_7_AB_1_GetSystemErrorMessageText(V_7_AB_1_ErrorNumber)
            MsgBox "An unexpected or occurred with the WaitForSingleObject function call." & vbCrLf & _
                    "System or Number: " & CStr(V_7_AB_1_ErrorNumber) & vbCrLf & _
                    "or Description:   " & F_7_AB_1_GetSystemErrorMessageText(V_7_AB_1_ErrorNumber)
            F_7_AB_1_ShellAndWaitSimple = False
            Exit Do
    End Select
''''''''''''''''''''''
' END OF MAIN TIMER LOOP
''''''''''''''''''''''
Loop


''''''''''''''''''''''''''''''''''''''''
' close the process
''''''''''''''''''''''''''''''''''''''''
CloseHandle V_7_AB_1_ProcessHandle
P_7_AB_1_ShellStillRunning = False
P_7_AB_1_CommandLine = vbNullString
P_7_AB_1_Abort = False

End Function


Public Function F_7_AB_1_AbortShellAndWaitSimple()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' F_7_AB_1_AbortShellAndWaitSimple
' This procedure should be called from the form on which messages are displayed when the form is closed and
' the shell programming is still running. It will terminate the WaitForSingleObject loop, close the process
' handle, and Kill the timer. It will leave the shell'd program running.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    P_7_AB_1_Abort = True
End Function

Public Function F_7_AB_1_ShellSimpleStillRunning() As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' F_7_AB_1_ShellSimpleStillRunning
' This returns TRUE or FALSE indicating whether the Shell'd program is still running.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    F_7_AB_1_ShellSimpleStillRunning = P_7_AB_1_ShellStillRunning
End Function

Public Function F_7_AB_1_GetSimpleCommmandLine() As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' F_7_AB_1_GetSimpleCommmandLine
' This function returns the command line passed to the Shell function.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    F_7_AB_1_GetSimpleCommmandLine = P_7_AB_1_CommandLine
End Function



Private Function F_7_AB_1_GetSystemErrorMessageText(A_7_AB_1_ErrorNumber As Long) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' F_7_AB_1_GetSystemErrorMessageText
'
' This function gets the system or message text that corresponds to the or code returned by the
' GetLastor API function or the .LastDllor property. It may be used ONLY for these or codes. These are NOT the or
' numbers returned by err.Number (for these ors, use .Description to get the description of the message).
' The or number MUST be the value returned by GetLastor or .LastDLLor.
'
' In general, you should use .LastDllor rather than GetLastor because under some circumstances the value of
' GetLastor will be reset to 0 before the value is returned to VB. err.LastDllError will always reliably return
' the last or number raised in a DLL.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim V_7_AB_1_ErrorText As String
Dim V_7_AB_1_TextLen As Long
Dim V_7_AB_1_FormatMessageResult As Long
Dim V_7_AB_1_LangID As Long

' an or value of 0 indicates no or. Return an empty string rather than the
' default Windows system message text.
If A_7_AB_1_ErrorNumber = 0& Then
    F_7_AB_1_GetSystemErrorMessageText = ""
    Exit Function
End If



V_7_AB_1_LangID = 0&
V_7_AB_1_ErrorText = String$(C_7_AB_1_FORMAT_MESSAGE_TEXT_LEN, " ")
V_7_AB_1_TextLen = Len(V_7_AB_1_ErrorText)
On Error Resume Next
V_7_AB_1_FormatMessageResult = 0&

V_7_AB_1_FormatMessageResult = FormatMessage( _
                dwFlags:=FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
                lpSource:=0&, _
                dwMessageId:=A_7_AB_1_ErrorNumber, _
                 dwLanguageId:=0&, _
                lpBuffer:=V_7_AB_1_ErrorText, _
                nSize:=Len(V_7_AB_1_ErrorText), _
                Arguments:=0&)
   
   
On Error GoTo 0
If V_7_AB_1_FormatMessageResult > 0 Then
    ' FormatMessage returned some text. Take the left V_7_AB_1_FormatMessageResult characters
    ' and return that text.
    V_7_AB_1_ErrorText = Left$(V_7_AB_1_ErrorText, V_7_AB_1_FormatMessageResult)
    F_7_AB_1_GetSystemErrorMessageText = V_7_AB_1_ErrorText
Else
    ' Format message didn't return any text. there is no text description for the specified or.
    F_7_AB_1_GetSystemErrorMessageText = "NO OR DESCRIPTION AVAILABLE"
End If


End Function




Private Function F_7_AB_1_GetHWndOfFormByControlVB6(A_7_AB_1_FormControl As VB.Control) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' F_7_AB_1_GetHWndOfFormByControlVB6
' This function returns the HWnd of the parent form of A_7_AB_1_FormControl.
' It does this by starting with A_7_AB_1_FormControl and looping, getting the Parent of
' objects until it reaches an object whose TypeOf is Not a VB.Control. This object
' will be the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim V_7_AB_1_Obj As Object
    Dim V_7_AB_1_Done As Boolean
    Set V_7_AB_1_Obj = A_7_AB_1_FormControl
    Do Until V_7_AB_1_Done
        Set V_7_AB_1_Obj = V_7_AB_1_Obj.Parent
        If Not TypeOf V_7_AB_1_Obj Is VB.Control Then
            V_7_AB_1_Done = True
            F_7_AB_1_GetHWndOfFormByControlVB6 = V_7_AB_1_Obj.hWnd
        End If
    Loop
End Function

Private Function F_7_AB_1_GetHWndOfFormByHWndVB6(A_7_AB_1_HWnd As Long) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' F_7_AB_1_GetHWndOfFormByHWndVB6
' This function returns the HWnd of the parent form of A_7_AB_1_HWnd. It accomplishes this
' quite simply with the GetAncestor API function.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''
' ensure we have a valid window
''''''''''''''''''''''''''''''''''''''''''
If IsWindow(hWnd:=A_7_AB_1_HWnd) = False Then
    MsgBox "The HWnd in A_7_AB_1_HWnd to procedure F_7_AB_1_GetHWndOfFormByHWndVB6 is not a valid window."
    Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''
' ensure we the window is visible
''''''''''''''''''''''''''''''''''''''''''
If IsWindowVisible(hWnd:=A_7_AB_1_HWnd) = False Then
    MsgBox "The HWnd in A_7_AB_1_HWnd to procedure F_7_AB_1_GetHWndOfFormByHWndVB6 is not a visible window."
    Exit Function
End If

F_7_AB_1_GetHWndOfFormByHWndVB6 = GetAncestor(A_7_AB_1_HWnd, GA_ROOT)

End Function


