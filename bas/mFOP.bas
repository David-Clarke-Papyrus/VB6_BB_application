Attribute VB_Name = "mFOP"
Option Explicit

'by adding a parameter that defauls false  and setting it true we can write a log file to debug this if necessary.
Function GetFOPCommandstring( _
    ByVal strFileIn As String, _
    ByVal strFileOut As String, _
    Optional ByVal writeLog As Boolean = False _
) As String
    Dim cp       As String
    Dim fopHome  As String
    Dim bat      As String
    
    '— build your classpath as before —
    fopHome = oPC.LocalFolder & "Executables\FOP\"
    cp = fopHome & "build\fop.jar;"
    cp = cp & fopHome & "lib\xml-apis.jar;"
    cp = cp & fopHome & "lib\xercesImpl-2.2.1.jar;"
    cp = cp & fopHome & "lib\xalan-2.4.1.jar;"
    cp = cp & fopHome & "lib\batik.jar;"
    cp = cp & fopHome & "lib\avalon-framework-cvs-20020806.jar;"
    cp = cp & fopHome & "lib\jimi-1.0.jar;"
    cp = cp & fopHome & "lib\jai_core.jar;"
    cp = cp & fopHome & "lib\jai_codec.jar"
    
    '— start the batch file —
    bat = "@echo off" & vbCrLf
    
    If writeLog Then
        '— define & timestamp the log file —
        bat = bat & "set LOG=""" & fopHome & "fop_run.log""" & vbCrLf
        bat = bat & "echo ==== %DATE% %TIME% ==== >> %LOG%" & vbCrLf & vbCrLf
        
        '— dump environment for diagnostics —
        bat = bat & "echo PATH=%PATH% >> %LOG%" & vbCrLf
        bat = bat & "echo JAVA_HOME=%JAVA_HOME% >> %LOG%" & vbCrLf & vbCrLf
        
        '— run FOP, redirecting all output into the log —
        bat = bat & "java -Xmx1024M -cp """ & cp & """ org.apache.fop.apps.Fop """ _
                   & strFileIn & """ """ & strFileOut & """ >> %LOG% 2>&1" & vbCrLf & vbCrLf
        
        '— record the exit code too —
        bat = bat & "echo EXIT CODE: %ERRORLEVEL% >> %LOG%" & vbCrLf
    Else
        '— simple, no-logging invocation —
        bat = bat & "java -Xmx1024M -cp """ & cp & """ org.apache.fop.apps.Fop """ _
                   & strFileIn & """ """ & strFileOut & """" & vbCrLf
    End If
    
    GetFOPCommandstring = bat
End Function


