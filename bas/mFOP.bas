Attribute VB_Name = "mFOP"
Option Explicit

Function GetFOPCommandstring(strFileIn As String, strFIleOut As String) As String
Dim cp As String
Dim fopHome As String
Dim strCommand As String
Dim fs As FileSystemObject

    fopHome = oPC.LocalFolder & "Executables\FOP\"
    cp = fopHome & "build\fop.jar;"
    cp = cp & fopHome & "\lib\xml-apis.jar;"
    cp = cp & fopHome & "\lib\xercesImpl-2.2.1.jar;"
    cp = cp & fopHome & "\lib\xalan-2.4.1.jar;"
    cp = cp & fopHome & "\lib\batik.jar;"
    cp = cp & fopHome & "\lib\avalon-framework-cvs-20020806.jar;"
    cp = cp & fopHome & "\lib\jimi-1.0.jar;"
    cp = cp & fopHome & "\lib\jai_core.jar;"
    cp = cp & fopHome & "\lib\jai_codec.jar"
    
    strCommand = "java -Xmx1024M -cp " & cp & " org.apache.fop.apps.Fop " & strFileIn & " " & strFIleOut
    GetFOPCommandstring = strCommand

End Function
