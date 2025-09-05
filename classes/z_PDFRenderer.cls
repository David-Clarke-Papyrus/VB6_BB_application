VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_PDFRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========[ c_PdfRenderer.cls ]==============================================
' A tiny, self-contained facade for producing a PDF via FOP.
' You pass paths and options; it takes care of quoting, picking the right FOP
' folder, calling FOP.BAT, waiting, and basic logging.

Option Explicit

' Defaults tuned to your existing code paths
Private Const DEFAULT_TIMEOUT_MS As Long = 20000   ' 20 seconds
Private Const DEFAULT_FOP_DIR As String = "Executables\FOP\"
Private Const DEFAULT_FOP_DIR_XSLT As String = "Executables\FOP-1.1\"

' Simple result info you can read after Render returns
Private mLastCommand As String
Private mLastExitCode As Long
Private mLastError As String

Public Property Get LastCommand() As String: LastCommand = mLastCommand: End Property
Public Property Get LastExitCode() As Long:  LastExitCode = mLastExitCode: End Property
Public Property Get LastError() As String:   LastError = mLastError: End Property

Public Function Render( _
    ByVal PdfOutPath As String, _
    Optional ByVal XmlPath As String = "", _
    Optional ByVal XslPath As String = "", _
    Optional ByVal FoPath As String = "", _
    Optional ByVal TemplateName As String = "", _
    Optional ByVal DebugLogPath As String = "", _
    Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS, _
    Optional ByVal BaseLocalFolder As String = "", _
    Optional ByVal BaseSharedFolder As String = "" _
) As Boolean
    On Error GoTo fail

    mLastCommand = ""
    mLastExitCode = -1
    mLastError = ""

    ' ---- Validate inputs ---------------------------------------------------
    If LenB(PdfOutPath) = 0 Then
        mLastError = "PdfOutPath is required."
        GoTo done
    End If

    Dim useFo As Boolean
    useFo = (LenB(FoPath) > 0)

    If Not useFo Then
        If LenB(XmlPath) = 0 Or LenB(XslPath) = 0 Then
            mLastError = "Provide FoPath OR (XmlPath and XslPath)."
            GoTo done
        End If
    End If

    ' ---- Build FOP home based on template extension (preserves legacy) ----
    Dim fopHome As String
    fopHome = ResolveFopHome(BaseLocalFolder, TemplateName)

    ' ---- Compose command line (always via FOP.BAT to avoid jar/classpath churn)
    Dim fopBat As String, args As String, logRedir As String, jreBin As String, fopDir As String, javaexe As String
    Dim preamble As String
    Dim diagnostics As String
    
    fopBat = Combine(fopHome, "FOP.BAT")

    If useFo Then
        args = "-fo " & Quote(FoPath) & " -pdf " & Quote(PdfOutPath)
    Else
        args = "-xml " & Quote(XmlPath) & " -xsl " & Quote(XslPath) & " -pdf " & Quote(PdfOutPath)
    End If

    fopDir = TrimTrailingSlash(fopHome)   ' ensure no trailing "\"
    jreBin = "\\DAVID-PC\PBKS_S\Executables\Java-8\bin"
  
    
    ' Optional: if you ship a JRE on your share, uncomment next line and adjust the path:
    ' preamble = "pushd " & Quote(fopDir) & " && set ""NOPAUSE=on"" && set ""PATH=\\DAVID-PC\PBKS_S\Executables\Java\bin;%PATH%"" && "
    ' If you rely on machine Java, use this:
    preamble = "pushd " & Quote(fopDir) & " && set ""NOPAUSE=on"" && "
    
    ' Small breadcrumbs land in the log so you can see if Java is found:
    diagnostics = "where java & java -version & echo CWD: %CD% & "
    
    ' Build the full one-liner (NO "cmd.exe /c" here; ShellAndWaitCmd adds it)
    ' Optional: redirect stderr to a debug log (keeps your existing pattern)
    If LenB(DebugLogPath) > 0 Then
            logRedir = " > " & Quote(DebugLogPath) & " 2>&1"
    End If

    mLastCommand = _
       "cmd.exe /c (" _
     & "pushd " & Quote(fopDir) & " && " _
     & "set ""NOPAUSE=on"" && " _
     & "set ""PATH=" & jreBin & ";%PATH%"" && " _
     & "where java & java -version && " _
     & "call .\fop.bat " & args & " & " _
     & "set EC=%ERRORLEVEL% & popd & exit /b %EC%" _
     & ") > " & Quote(DebugLogPath) & " 2>&1"
    


    mLastExitCode = ShellAndWaitHidden(mLastCommand, TimeoutMs)


    ' ---- Confirm the PDF actually exists (your code already loops briefly; we do a final check)
    If FileExists(PdfOutPath) Then
        Render = True
    Else
        If mLastExitCode <> 0 And LenB(mLastError) = 0 Then
            mLastError = "FOP exited with code " & CStr(mLastExitCode) & " and PDF not found."
        ElseIf LenB(mLastError) = 0 Then
            mLastError = "PDF not created: " & PdfOutPath
        End If
        Render = False
    End If

    GoTo done

fail:
    mLastError = "Exception: " & Err.Number & " - " & Err.Description
    Render = False

done:
End Function

' ---- Helpers ---------------------------------------------------------------
Private Function TrimTrailingSlash(ByVal s As String) As String
    If LenB(s) > 0 And Right$(s, 1) = "\" Then
        TrimTrailingSlash = Left$(s, Len(s) - 1)
    Else
        TrimTrailingSlash = s
    End If
End Function
Private Function ResolveFopHome(ByVal baseLocal As String, ByVal TemplateName As String) As String
    Dim isXslt As Boolean
    Dim Ext As String: Ext = UCase$(Right$(TemplateName, 4))
    isXslt = (Ext = "XSLT")

    If LenB(baseLocal) = 0 Then baseLocal = App.path & "\" ' very conservative default

    If isXslt Then
        ResolveFopHome = Combine(baseLocal, DEFAULT_FOP_DIR_XSLT) ' ...\Executables\FOP-1.1\
    Else
        ResolveFopHome = Combine(baseLocal, DEFAULT_FOP_DIR)      ' ...\Executables\FOP\
    End If
End Function

Private Function ShellAndWaitCmd(ByVal fullCmd As String, ByVal TimeoutMs As Long) As Long
    ' Run via cmd.exe so redirection (2>) works reliably.
    Dim sh As Object, execObj As Object, started As Single
    Set sh = CreateObject("WScript.Shell")
    Set execObj = sh.exec(fullCmd)

    started = Timer
    Do While execObj.Status = 0
        DoEvents
        If TimeoutMs > 0 Then
            If (Timer - started) * 1000& > TimeoutMs Then
                On Error Resume Next
                execObj.Terminate
                On Error GoTo 0
                ShellAndWaitCmd = 1460      ' mimic WAIT_TIMEOUT
                Exit Function
            End If
        End If
    Loop

    ShellAndWaitCmd = execObj.exitCode
End Function

Private Function Quote(ByVal s As String) As String
    If LenB(s) = 0 Then Quote = "" Else Quote = """" & s & """"
End Function

Private Function FileExists(ByVal path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(path)
End Function

Private Function Combine(ByVal leftPart As String, ByVal rightPart As String) As String
    If LenB(leftPart) = 0 Then
        Combine = rightPart
    ElseIf Right$(leftPart, 1) = "\" Or Left$(rightPart, 1) = "\" Then
        Combine = leftPart & rightPart
    Else
        Combine = leftPart & "\" & rightPart
    End If
End Function
'===========================================================================


