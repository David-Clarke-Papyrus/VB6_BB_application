Attribute VB_Name = "mExecuteCommand"
'The CreatePipe function creates an anonymous pipe,
'and returns handles to the read and write ends of the pipe.

Private Declare Function CreatePipe Lib "kernel32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As Any, _
    ByVal nSize As Long) As Long

'Used to read the the pipe filled by the process create
'with the CretaProcessA function
Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long


'Structure used by the CreateProcessA function
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'Structure used by the CreateProcessA function
Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
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

'Structure used by the CreateProcessA function
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

'This function launch the the commend and return the relative process
'into the PROCESS_INFORMATION structure
Private Declare Function CreateProcessA Lib "kernel32" ( _
    ByVal lpApplicationName As Long, _
    ByVal lpCommandLine As String, _
    lpProcessAttributes As SECURITY_ATTRIBUTES, _
    lpThreadAttributes As SECURITY_ATTRIBUTES, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long

'Close opened handle
Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hHandle As Long) As Long

Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, ByVal lpBuffer As Long, _
 ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long


'Consts for the above functions
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL As Long = 1


Private mCommand As String          'Private variable for the CommandLine property
Private mOutputs As String          'Private variable for the ReadOnly Outputs property


Public Function ExecuteCommand(ByVal CommandLine As String, Optional bShowWindow As Boolean = False, Optional sCurrentDir As String) As String
        Dim proc As PROCESS_INFORMATION     'Process info filled by CreateProcessA
        Dim ret As Long                     'long variable for get the return value of the
        'API functions
        Dim start As STARTUPINFO            'StartUp Info passed to the CreateProceeeA
        'function
        Dim sa As SECURITY_ATTRIBUTES       'Security Attributes passeed to the
        'CreateProcessA function
        Dim hReadPipe As Long               'Read Pipe handle created by CreatePipe
        Dim hWritePipe As Long              'Write Pite handle created by CreatePipe
        Dim lngBytesRead As Long            'Amount of byte read from the Read Pipe handle
        Dim strBuff As String * 256         'String buffer reading the Pipe


        'if the parameter is not empty update the CommandLine property
        On Error GoTo ExecuteCommand_Error

10        If Len(CommandLine) > 0 Then
20            mCommand = CommandLine
30        End If

        'if the command line is empty then exit whit a error message
40        If Len(CommandLine) = 0 Then
50            ErrSaveToFile "Command Line empty in procedure ExecuteCommand of module modPipedOutput."
60            Exit Function
70        End If

        'Create the Pipe
80        sa.nLength = Len(sa)
90        sa.bInheritHandle = 1&
100       sa.lpSecurityDescriptor = 0&
110       ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)

120       If ret = 0 Then
            'If an error occur during the Pipe creation exit
130           ErrSaveToFile "CreatePipe failed. Error: " & Err.LastDllError & " (" & Err.LastDllError & ") in procedure ExecuteCommand of module modPipedOutput."
140           Exit Function
150       End If

        '    ret = CreatePipe(hInReadPipe, hInWritePipe, sa, 0)


        'Launch the command line application
160       start.cb = Len(start)
170       start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW

        'set the StdOutput and the StdError output to the same Write Pipe handle
180       start.hStdOutput = hWritePipe
190       start.hStdError = hWritePipe
        '    start.hStdInput = hInReadPipe
200       If bShowWindow Then
210           start.wShowWindow = SW_SHOWNORMAL
220       Else
230           start.wShowWindow = SW_HIDE
240       End If

        'Execute the command
250       If Len(sCurrentDir) = 0 Then
260           ret& = CreateProcessA(0&, mCommand, sa, sa, 1&, _
                  NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)
270       Else
280           ret& = CreateProcessA(0&, mCommand, sa, sa, 1&, _
                  NORMAL_PRIORITY_CLASS, 0&, sCurrentDir, start, proc)
290       End If

300       If ret <> 1 Then
            'if the command is not found ....
310           ErrSaveToFile "File or command not found in procedure ExecuteCommand of module modPipedOutput."
320           Exit Function
330       End If

        'Now We can ... must close the hWritePipe
340       ret = CloseHandle(hWritePipe)
        '    ret = CloseHandle(hInReadPipe)
350       mOutputs = vbNullString

        'Read the ReadPipe handle
360       Do
370           ret = ReadFile(hReadPipe, strBuff, 256, lngBytesRead, 0&)

380           mOutputs = mOutputs & Left$(strBuff, lngBytesRead)
            'Send data to the object via ReceiveOutputs event

390       Loop While ret <> 0

        'Close the opened handles
400       ret = CloseHandle(proc.hProcess)
410       ret = CloseHandle(proc.hThread)
420       ret = CloseHandle(hReadPipe)

        'Return the Outputs property with the entire DOS output
430       ExecuteCommand = mOutputs

        On Error GoTo 0
440       Exit Function

ExecuteCommand_Error:

450       ErrSaveToFile "Error " & Err.Number & " (" & Err.Description & ") in line number " & Erl & " of procedure ExecuteCommand of Module modPipedOutput"
End Function


Function GetPipedOutput(hPipeHandle As Long) As String
        Dim lngBytesRead As Long
        Dim lngBytesAvail As Long
        Dim lngBytesLeft As Long
        Dim sOutBuffer As String * 256
        Dim ret As Long

10        GetPipedOutput = ""

20        PeekNamedPipe hPipeHandle, API_NULL, 0&, lngBytesRead, lngBytesAvail, lngBytesLeft
30        Do While lngBytesAvail > 0
40            lngBytesRead = 0
50            ret = ReadFile(hPipeHandle, sOutBuffer, 256, lngBytesRead, 0&)
60            If lngBytesRead > 0 Then

70                GetPipedOutput = GetPipedOutput & Left$(sOutBuffer, lngBytesRead)
80            Else
90                Exit Do
100           End If
110           PeekNamedPipe hPipeHandle, API_NULL, 0&, lngBytesRead, lngBytesAvail, lngBytesLeft
120       Loop
End Function


