VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_XML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim XMLDoc As ujXML
Dim sXMLFIleName As String
Dim strPath As String
Dim strCommand As String
Dim cnt As Integer
Dim sHTMLFileName As String
Dim sPDFFilename As String

Public Property Get PDF_Filename() As String
    PDF_Filename = sPDFFilename
End Property

Public Sub GenerateXMLBranchReport(rs As ADODB.Recordset, pBranchCode As String, pBranchName As String, pContact As String, pEmailAddress As String)
Dim fs As New FileSystemObject

    Set XMLDoc = New ujXML
                        p 21
    With XMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "root"
            .chCreate "MessageType"
                .elText = "BranchCreditorErrors"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "TemplateName"
                .elText = "BCE"
            .elCreateSibling "DestinationEmail"
                .elText = pEmailAddress
            .elCreateSibling "BranchCode"
                .elText = pBranchCode
            .elCreateSibling "BranchName"
                .elText = pBranchName
            .elCreateSibling "Contact"
                .elText = pContact
            Do While Not rs.EOF
                .elCreateSibling "DetailLine", True
                .chCreate "DateReceived"
                    .elText = FNS(rs.Fields("P_DATERECEIVEDATHO"))
                .elCreateSibling "Acno", True
                    .elText = FNS(rs.Fields("P_Acno")) & Left(FNS(rs.Fields("P_TPNAME")), 50)
                .elCreateSibling "Description", True
                    .elText = FNS(rs.Fields("P_Descr"))
                .elCreateSibling "Ref", True
                    .elText = FNS(rs.Fields("P_Ref"))
                .elCreateSibling "JournalType", True
                    .elText = FNS(rs.Fields("P_JournalType"))
                .elCreateSibling "Amt", True
                    .elText = FNS(rs.Fields("P_Amt"))
                .navUP
                rs.MoveNext
            Loop
    End With
'FINALLY PRODUCE THE .XML FILE
    sXMLFIleName = oPC.LocalRoot & "\TEMP\" & "BCE_" & Format(Now(), "YYMMDDHHNN") & ".xml"
    With XMLDoc
        If fs.FileExists(sXMLFIleName) Then
            fs.DeleteFile sXMLFIleName
        End If
        .docWriteToFile (sXMLFIleName), False, "UNICODE", ""
    End With

End Sub

Public Sub GenerateXMLBranchMissingTransactionsReport(rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim fs As New FileSystemObject

    rs.MoveFirst
    Set XMLDoc = New ujXML
                        p 21
    With XMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "root"
            .chCreate "MessageType"
                .elText = "BranchMissingDocuments"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "TemplateName"
                .elText = "BMD_MSG"
            .elCreateSibling "DestinationEmail"
                .elText = CStr(rs.Fields("STORE_EMAIL"))
            .elCreateSibling "BranchCode"
                .elText = CStr(rs.Fields("STORE_CODE"))
            .elCreateSibling "BranchName"
                .elText = CStr(rs.Fields("STORE_NAME"))
            .elCreateSibling "Contact"
                .elText = CStr(rs.Fields("STORE_CONTACT"))
            .elCreateSibling "DateOfAudit"
                .elText = Format(Now(), "yyyy-mm-dd  HH:NN")
            Do While Not rs.EOF
                .elCreateSibling "DetailLine", True
                .chCreate "Type", True
                    .elText = FNS(rs.Fields("TRTYPE"))
                .elCreateSibling "DocCode", True
                    .elText = FNS(rs.Fields("MISSINGNUMBER"))
                .elCreateSibling "Note", True
                    .elText = FNS(rs.Fields("NOTE"))
                .navUP
                rs.MoveNext
            Loop
    End With
'FINALLY PRODUCE THE .XML FILE
    sXMLFIleName = oPC.LocalRoot & "\TEMP\" & "BMD_" & Format(Now(), "YYMMDDHHNN") & ".xml"
    With XMLDoc
        If fs.FileExists(sXMLFIleName) Then
            fs.DeleteFile sXMLFIleName
        End If
        .docWriteToFile (sXMLFIleName), False, "UTF-8", ""
    End With
    
    rs.MoveFirst

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_XML.GenerateXMLBranchMissingTransactionsReport(rs)", rs
End Sub

Public Sub CreateFiles(pType As String, pPDFFilename As String)
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim oTF As New z_TextFile

Dim objXSL As New MSXML2.DOMDocument30
Dim pWholeMessage As String
Dim sFOFileName As String

Dim opXMLDOC As New MSXML2.DOMDocument30
Dim strWorkingFolder As String

        objXSL.async = False
        objXSL.validateOnParse = False
        objXSL.resolveExternals = False
        strPath = oPC.SharedFolderRoot & "\Templates\" & pType & "_HTML.XSLT"
        Set fs = New FileSystemObject
        If fs.FileExists(strPath) Then
            objXSL.Load strPath
        End If
    
        sHTMLFileName = Replace(UCase(sXMLFIleName), ".XML", ".HTML")
        If fs.FileExists(sHTMLFileName) Then
            fs.DeleteFile sHTMLFileName, True
        End If
        oTF.OpenTextFileToAppend sHTMLFileName
        pWholeMessage = XMLDoc.docObject.transformNode(objXSL)
        oTF.WriteToTextFile pWholeMessage
        oTF.CloseTextFile
        
        
        Set objXSL = Nothing
        Set objXSL = New MSXML2.DOMDocument30
        objXSL.async = False
        objXSL.validateOnParse = False
        objXSL.resolveExternals = False
        strPath = oPC.SharedFolderRoot & "\Templates\" & pType & "_FO.XSL"
        Set fs = New FileSystemObject
        If fs.FileExists(strPath) Then
            objXSL.Load strPath
        End If
                        p 6
        Set opXMLDOC = New MSXML2.DOMDocument30
        opXMLDOC.async = False
        opXMLDOC.validateOnParse = False
        opXMLDOC.resolveExternals = False
        XMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC
        
        sFOFileName = strWorkingFolder & Replace(UCase(sXMLFIleName), ".XML", ".FO", , , vbTextCompare)
        sPDFFilename = strWorkingFolder & Replace(UCase(sXMLFIleName), ".XML", ".PDF", , , vbTextCompare)
                                p 7
        docWriteTostream sFOFileName, opXMLDOC, "UNICODE"
        
'Stage 2 Convert the .FO file to .PDF and clean up
       ''' strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & sFOFile & " " & sPDFFile
        strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & sFOFileName & " " & sPDFFilename
        F_7_AB_1_ShellAndWaitSimple strCommand
        'IF IT FAILS HERE CHECK THAT JAVA IS INSTALLED
        'We do the following because the batch file executes and is not waited for by the F_7_AB_1_ShellAndWaitSimple routine
        'so the FO file is sometimes deleted before the .PDF file is produced.
        cnt = 0
        Do While Not fs.FileExists(sPDFFilename) And cnt < 40
            MsgWaitObj 1000
            cnt = cnt + 1
        Loop
        If fs.FileExists(sFOFileName) Then
            fs.DeleteFile sFOFileName
        End If

        pPDFFilename = sPDFFilename
    Exit Sub
errHandler:
    ErrorIn "z_XML.CreateFiles(pType,pPDFFilename)", Array(pType, pPDFFilename)
End Sub
Public Sub docWriteTostream(ByVal FilePath As String, obj As MSXML2.DOMDocument30, _
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
    ErrorIn "a_PO.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
End Sub

