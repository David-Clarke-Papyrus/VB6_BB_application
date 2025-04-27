Attribute VB_Name = "mMEC"

Function GetMECCommandstring(strFileIn As String, sFormatFilename As String, strFIleOut As String) As String
Dim cp As String
Dim MecHome As String
Dim strCommand As String
Dim fs As FileSystemObject
Dim strJH As String
Dim strLib As String

    strJH = Environ$("JAVA_HOME")
    strLib = Replace(Replace(strLib, """", ""), "bin", "lib")
    strJavaEXE = Replace(strJH, "\bin", "\bin\java.exe")
    MecHome = oPC.LocalFolder & "MEC\"
    cp = """" & ";eagle.jar;./eaglehelp;./jlib;./build/jlib;jlib/activation.jar;jlib/ant.jar;jlib/antlr.jar;jlib/bsf.jar;jlib/cryptix32-pgp.jar;" _
                    & "jlib/cryptix32.jar;jlib/fop_0_15_0.jar;jlib/hsql.jar;jlib/IBMParser.jar;jlib/imap.jar;jlib/jasper.jar;jlib/jhall.jar;" _
                    & "jlib/jnet.jar;jlib/jsse.jar;jlib/mail.jar;jlib/mailapi.jar;jlib/mbox.jar;jlib/mec_util.jar;jlib/merlot;jlib/Merlot.jar;" _
                    & "jlib/MerlotIcon.jar;jlib/mpEDIT.jar;jlib/NetComponents.jar;jlib/optional.jar;jlib/pop3.jar;jlib/servlet.jar;" _
                    & "jlib/smtp.jar;jlib/vssver.scc;jlib/w3c_svg4fop.jar;jlib/webserver.jar;jlib/xalan_1_2_1.jar;jlib/xerces_1_2_1.jar;" _
                    & "jlib/xml.jar;jlib/merlot;" & strLib & "/tools.jar" & """"
    
    strCommand = strJavaEXE & " -Xmx" & IIf(sJavaMemoryAllocation = "", "128", sJavaMemoryAllocation) _
                 & "M -cp " & cp & " de.mendelson.eagle.converter.edixml.EDIXMLConverter -debug " _
                 & " -ediin " & strFileIn _
                 & " -formatin " & sFormatFilename _
                 & " -xmlout " & strFIleOut _
                 & " -ncs -filter " & oPC.SharedFolderRoot & "\MEC\filter\PBKSCLEAN.filter > " & oPC.SharedFolderRoot & "\Mec_results.txt"
    GetMECCommandstring = strCommand
    LogSaveToFile (strCommand)
    
End Function


''

