Attribute VB_Name = "modShellHidden"
Option Explicit

' --- Win32 API declarations ---
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread  As Long
    dwProcessId As Long
    dwThreadId  As Long
End Type

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpCommandLine As String, _
    ByVal lpProcessAttributes As Long, _
    ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" ( _
    ByVal hProcess As Long, _
    lpExitCode As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" ( _
    ByVal hProcess As Long, _
    ByVal uExitCode As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

' --- constants ---
Private Const NORMAL_PRIORITY_CLASS As Long = &H20&
Private Const CREATE_NO_WINDOW       As Long = &H8000000
Private Const STARTF_USESHOWWINDOW   As Long = &H1
Private Const SW_HIDE                As Long = 0

Private Const WAIT_OBJECT_0          As Long = 0
Private Const WAIT_TIMEOUT           As Long = &H102
Public Const SHELL_TIMEOUT_EXIT      As Long = 1460   ' conventional timeout code

' Run a command hidden, wait up to TimeoutMs, return process exit code (or 1460 on timeout).
Public Function ShellAndWaitHidden(ByVal CommandLine As String, _
                                   ByVal TimeoutMs As Long, _
                                   Optional ByVal CurrentDir As String = vbNullString) As Long
    Dim si As STARTUPINFO, pi As PROCESS_INFORMATION
    Dim ok As Long, waitRc As Long, exitCode As Long

    si.cb = Len(si)
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = SW_HIDE

    ok = CreateProcess(vbNullString, CommandLine, 0&, 0&, 0&, _
                       NORMAL_PRIORITY_CLASS Or CREATE_NO_WINDOW, _
                       0&, CurrentDir, si, pi)
    If ok = 0 Then
        ' Could not start; use a generic nonzero code
        ShellAndWaitHidden = 1
        Exit Function
    End If

    waitRc = WaitForSingleObject(pi.hProcess, TimeoutMs)
    If waitRc = WAIT_TIMEOUT Then
        ' kill it and return timeout code
        Call TerminateProcess(pi.hProcess, SHELL_TIMEOUT_EXIT)
        ShellAndWaitHidden = SHELL_TIMEOUT_EXIT
    Else
        Call GetExitCodeProcess(pi.hProcess, exitCode)
        ShellAndWaitHidden = exitCode
    End If

    Call CloseHandle(pi.hThread)
    Call CloseHandle(pi.hProcess)
End Function


