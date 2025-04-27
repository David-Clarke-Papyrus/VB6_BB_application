VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub PrintXML(pFilename As String, Optional pWorkingFolder As String, Optional pTemplateFolder As String, _
                    Optional pExecutablesFolder As String, Optional pUseXPDFPrint As Boolean, _
                    Optional pPDFFilepath As String, Optional pTemplateName As String)
    On Error GoTo errHandler
Dim oXML As MSXML2.DOMDocument60
Dim oXSL As MSXML2.DOMDocument60
Dim oXMLFO As MSXML2.DOMDocument60
Dim strPath As String
Dim strFOFile As String
Dim strPDFFile As String
Dim strCommand As String
Dim strTemplateName As String
Dim strExecutable As String
Dim strPrinter As String
Dim fs As FileSystemObject
Dim j As Integer
Dim xmlFile As New ujXML
Dim iCopyCount As Long
Dim strOldPrinter As String
Dim p As Printer
Dim pdf As XpdfPrint.XpdfPrint
Dim Res As Boolean

    
    xmlFile.docReadFromFile pFilename, "UNICODE"
    
    xmlFile.navTop
    On Error Resume Next
    xmlFile.docXslLocate ("TemplateName")
    strTemplateName = xmlFile.Element.text
    If pTemplateName > "" Then strTemplateName = pTemplateName
    xmlFile.navTop
    If xmlFile.docXslLocate("CopyCount") Then
        iCopyCount = CInt(xmlFile.Element.text)
    Else
        iCopyCount = 1
    End If
    
    xmlFile.navTop
    If xmlFile.docXslLocate("Printer") Then
        strPrinter = xmlFile.Element.text
    Else
        strPrinter = Printer.DeviceName
    End If
    Set oXML = New MSXML2.DOMDocument60
    oXML.Load pFilename
    
    
'WRITE THE .PDF FILE IF NECESSARY
'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file
        Set oXSL = Nothing
        Set oXSL = New MSXML2.DOMDocument60
        oXSL.async = False
        oXSL.ValidateOnParse = False
        oXSL.resolveExternals = False
        If pTemplateFolder = "" Then
         '   pTemplateFolder = strSharedServerFolder & "\Templates\"
            pTemplateFolder = oPC.SharedFolderRoot & "\Templates\"
        End If
        strPath = pTemplateFolder & strTemplateName
        Set fs = New FileSystemObject
        If fs.FileExists(strPath) Then
            oXSL.Load strPath
        End If

        Set oXMLFO = New MSXML2.DOMDocument60
        oXMLFO.async = False
        oXMLFO.ValidateOnParse = False
        oXMLFO.resolveExternals = False
        oXML.transformNodeToObject oXSL, oXMLFO
        If pWorkingFolder = "" Then
          '  pWorkingFolder = strSharedServerFolder & "\TEMP"
            pWorkingFolder = oPC.SharedFolderRoot & "\TEMP"
        End If
        If pWorkingFolder > "" Then
            strFOFile = pWorkingFolder & "\tmp" & ".FO"
            If pPDFFilepath > "" Then
                strPDFFile = pPDFFilepath
            Else
                strPDFFile = pWorkingFolder & "\tmp" & ".PDF"
            End If
        End If
        docWriteTostream strFOFile, oXMLFO, "UNICODE"
        
'Stage 2 Convert the .FO file to .PDF and clean up
        If pExecutablesFolder = "" Then
           ' pExecutablesFolder = strSharedServerFolder & "\Executables"
            pExecutablesFolder = oPC.SharedFolderRoot & "\Executables"
        End If
       ' strCommand = pExecutablesFolder & "\FOP\FOP.BAT" & " " & strFOFile & " " & strPDFFile
        
        strCommand = GetFOPCommandstring(strFOFile, strPDFFile)
        Res = F_7_AB_1_ShellAndWaitSimple(strCommand, vbHide, 10000)
        If fs.FileExists(strFOFile) Then
            fs.DeleteFile strFOFile
        End If

    If pUseXPDFPrint Then
        Set pdf = New XpdfPrint.XpdfPrint
        pdf.loadFile strPDFFile
        pdf.printPDF3 pdf.printSetPrinterFlag Or pdf.printSetCopiesFlag, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, iCopyCount, _
        IdentifyPrinter(strPrinter), 0, 0
        Set pdf = Nothing
    Else
        strOldPrinter = Printer.DeviceName
        SetDefaultPrinter strPrinter
        strExecutable = GetPDFExecutable(strPDFFile)
        If strExecutable = "" Then
            MsgBox "There is no application set on this computer to open the file: " & strPDFFile & ". The document cannot be displayed", vbOKOnly, "Can't do this"
        Else
            Shell strExecutable & " " & strPDFFile
        End If
'        Shell strExecutable, vbHide
'        Forms(0).Text1.LinkMode = vbLinkNone
'        Forms(0).Text1.LinkTopic = "acroview|control"
'        Forms(0).Text1.LinkMode = vbLinkManual
'        Forms(0).Text1.LinkTimeout = 32767
'        For j = 1 To iCopyCount
'            Forms(0).Text1.LinkExecute "[FilePrintSilent(""" & strPDFFile & """)]"
'        Next
'        Forms(0).Text1.LinkMode = vbLinkNone
'        Forms(0).Text1.LinkTopic = "acroview|control"
'        Forms(0).Text1.LinkMode = vbLinkManual
'        Forms(0).Text1.LinkTimeout = 32767
'        Forms(0).Text1.LinkExecute "[AppExit]"
'        SetDefaultPrinter strOldPrinter
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "zXML.PrintInvoice(pFilename)", Array(pFilename)
End Sub

Public Sub docWriteTostream(ByVal FilePath As String, obj As MSXML2.DOMDocument60, _
                Optional ByVal CharSet As String = "UNICODE")
    On Error GoTo errHandler
    Dim s As Object
    Set s = CreateObject("ADODB.Stream")
    With s
        If CharSet <> "" Then .CharSet = CharSet
        .Open
        .WriteText obj.xml
        .SaveToFile FilePath, 2 'adSaveCreateOverWrite
        .Close
    End With
    Exit Sub
errHandler:
    ErrorIn "ujXML.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
End Sub
'


