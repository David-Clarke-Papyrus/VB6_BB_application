Attribute VB_Name = "zXML"
Option Explicit

Public Sub PrintInvoice(pFilename As String, pLogoPaht As String)
Dim oXML As MSXML2.DOMDocument30
Dim oXSL As MSXML2.DOMDocument30
Dim oXMLFO As MSXML2.DOMDocument30


    Set oXML = New MSXML2.DOMDocument30
    oXML.Load pFilename
'WRITE THE .HTML FILE
''''''    objXSL.async = False
''''''    objXSL.validateOnParse = False
''''''    objXSL.resolveExternals = False
''''''    strPath = oPC.SharedFolderRoot & "\Templates\IN_1.xslt"
''''''    Set fs = New FileSystemObject
''''''    If fs.FileExists(strPath) Then
''''''        objXSL.Load strPath
''''''    End If
''''''
''''''    If fs.FileExists(oPC.SharedFolderRoot & "\Emails\IN_" & Me.DocCode & ".HTML") Then
''''''        fs.DeleteFile oPC.SharedFolderRoot & "\Emails\IN_" & Me.DocCode & ".HTML", True
''''''    End If
''''''    oTF.OpenTextFileToAppend oPC.SharedFolderRoot & "\Emails\IN_" & Me.DocCode & ".HTML"
''''''    oTF.WriteToTextFile xMLDoc.docObject.transformNode(objXSL)
''''''    oTF.CloseTextFile
    
'WRITE THE .PDF FILE IF NECESSARY
'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file
        Set oXSL = Nothing
        Set oXSL = New MSXML2.DOMDocument30
        oXSL.async = False
        oXSL.validateOnParse = False
        oXSL.resolveExternals = False
        strPath = oPC.SharedFolderRoot & "\Templates\IN_FO_1.xsl"
        Set fs = New FileSystemObject
        If fs.FileExists(strPath) Then
            oXSL.Load strPath
        End If


        Set oXMLFO = New MSXML2.DOMDocument30
        oXMLFO.async = False
        oXMLFO.validateOnParse = False
        oXMLFO.resolveExternals = False
        oXML.transformNodeToObject oXSL, oXMLFO

        strFOFile = oPC.SharedFolderRoot & "\Emails\IN_" & Me.DocCode & ".FO"
        strPDFFile = oPC.SharedFolderRoot & "\Emails\IN_" & Me.DocCode & ".PDF"
        docWriteTostream strFOFile, oXMLFO, "UNICODE"
        
'Stage 2 Convert the .FO file to .PDF and clean up
        strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFile & " " & strPDFFile
        ShellAndWait strCommand, vbHide, False
        If fs.FileExists(strFOFile) Then
            fs.DeleteFile strFOFile
        End If
    End If
End Sub
