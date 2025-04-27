Attribute VB_Name = "mFOP"
Option Explicit

Function GetFOPCommandstring(strFileIn As String, strFIleOut As String) As String
Dim cp As String
Dim fopHome As String
Dim strCommand As String
Dim fs As FileSystemObject

    fopHome = strLocalRootFolder & "\Executables\FOP\"
    cp = fopHome & "build\fop.jar;"
    cp = cp & fopHome & "\lib\xml-apis.jar;"
    cp = cp & fopHome & "\lib\xercesImpl-2.2.1.jar;"
    cp = cp & fopHome & "\lib\xalan-2.4.1.jar;"
    cp = cp & fopHome & "\lib\batik.jar;"
    cp = cp & fopHome & "\lib\avalon-framework-cvs-20020806.jar;"
    cp = cp & fopHome & "\lib\jimi-1.0.jar;"
    cp = cp & fopHome & "\lib\jai_core.jar;"
    cp = cp & fopHome & "\lib\jai_codec.jar"
    
    strCommand = "java -Xmx" & IIf(sJavaMemoryAllocation = "", "128", sJavaMemoryAllocation) & "M -cp " & cp & " org.apache.fop.apps.Fop " & strFileIn & " " & strFIleOut
    GetFOPCommandstring = strCommand
'Rem %~dp0 is the expanded pathname of the current script under NT
'set LOCAL_FOP_HOME=
'if "%OS%"=="Windows_NT" set LOCAL_FOP_HOME=%~dp0
'
'set LIBDIR=%LOCAL_FOP_HOME%lib
'set LOCALCLASSPATH=%LOCAL_FOP_HOME%build\fop.jar
'set LOCALCLASSPATH=%LOCALCLASSPATH%;%LIBDIR%\xml-apis.jar
'set LOCALCLASSPATH=%LOCALCLASSPATH%;%LIBDIR%\xercesImpl-2.2.1.jar
'set LOCALCLASSPATH=%LOCALCLASSPATH%;%LIBDIR%\xalan-2.4.1.jar
'set LOCALCLASSPATH=%LOCALCLASSPATH%;%LIBDIR%\batik.jar
'set LOCALCLASSPATH=%LOCALCLASSPATH%;%LIBDIR%\avalon-framework-cvs-20020806.jar
'set LOCALCLASSPATH=%LOCALCLASSPATH%;%LIBDIR%\jimi-1.0.jar
'set LOCALCLASSPATH=%LOCALCLASSPATH%;%LIBDIR%\jai_core.jar
'set LOCALCLASSPATH=%LOCALCLASSPATH%;%LIBDIR%\jai_codec.jar
'java -cp "%LOCALCLASSPATH%" org.apache.fop.apps.Fop %1 %2 %3 %4 %5 %6 %7 %8

End Function
