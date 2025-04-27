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
Dim strTemplatePath As String
Dim strFOFile As String
Dim strPDFFile As String
Dim strTemplateName As String
Dim strExecutable As String
Dim strPrinter As String
Dim oXML As MSXML2.DOMDocument30
Dim oXSL As MSXML2.DOMDocument30
Dim oXMLFO As MSXML2.DOMDocument30
Dim strCommand As String
Dim fs As New FileSystemObject
Dim bPrintingInvoice As Boolean
Dim bPrintPackingSlip As Boolean
Dim iCopyCount As Long
Dim strExecutablesFolder As String
Dim Res
Dim xmlFile As New ujXML
Public Function PrintXML(pFilename As String, Optional pWorkingFolder As String, Optional pTemplateFolder As String, Optional pExecutablesFolder As String, Optional PDFTool As String = "A") As Boolean
    On Error GoTo errHandler
 Dim bIsProForma As Boolean
 Dim Res As Boolean
 
    PrintXML = True
    strExecutablesFolder = pExecutablesFolder
    SetupFoldersandDefaultFilenames pTemplateFolder, pWorkingFolder
    xmlFile.docReadFromFile pFilename, "UNICODE"
    xmlFile.navTop
    xmlFile.docXslLocate ("TemplateName")
    strTemplateName = xmlFile.Element.Text
    xmlFile.navTop
    Res = xmlFile.docXslLocate("IsProforma")
    If Res Then
        bIsProForma = (xmlFile.Element.Text = "Y")
    Else
        bIsProForma = False
    End If
    strTemplatePath = pTemplateFolder & strTemplateName & "_FO.XSL"
    bPrintingInvoice = (Left(strTemplateName, 2) = "IN")
    
    xmlFile.navTop
    If xmlFile.docXslLocate("CopyCount") Then
        iCopyCount = CInt(xmlFile.Element.Text)
    Else
        iCopyCount = 1
    End If
    
    xmlFile.navTop
    If xmlFile.docXslLocate("Printer") Then
        strPrinter = xmlFile.Element.Text
    Else
        strPrinter = Printer.DeviceName
    End If
        
    
    Res = CreateFOUsingMSXML
                                If Res = False Then
                                    LogSaveToFile "PrintXML: res:" & CStr(Res)
                                    PrintXML = False
                                    Exit Function
                                End If
    
    Res = ConvertFOtoPDF
    If Not Res Then LogSaveToFile "Dispatcher: zXML.PrintXML: call to ConvertFOtoPDF failed"
    If Res Then PrintWithSelectedTool PDFTool
        
' PRINT PACKING SLIP IF NECESSARY
    iCopyCount = 1
    If bPrintingInvoice Then
        bPrintPackingSlip = (GetProperty("PrintPackingSlip") = "TRUE")
        If bPrintPackingSlip And (Not bIsProForma) Then
            strTemplateName = "PackingSlip"
            strTemplatePath = pTemplateFolder & strTemplateName & "_FO.XSL"
            SetupFoldersandDefaultFilenames pTemplateFolder, pWorkingFolder

            Res = CreateFOUsingMSXML
                                        If Res = False Then
                                            LogSaveToFile "PrintXML: res:" & CStr(Res)
                                            PrintXML = False
                                            Exit Function
                                        End If
                
            Res = ConvertFOtoPDF
            
            If Res Then PrintWithSelectedTool PDFTool
            
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "zXML.PrintXML(pFilename,pWorkingFolder,pTemplateFolder,pExecutablesFolder,PDFTool)", _
         Array(pFilename, pWorkingFolder, pTemplateFolder, pExecutablesFolder, PDFTool)
End Function

Public Sub docWriteTostream(ByVal FilePath As String, obj As MSXML2.DOMDocument30, _
                Optional ByVal CharSet As String = "UNICODE")
    On Error GoTo errHandler
    Dim s As Object
    Set s = CreateObject("ADODB.Stream")
    With s
        If CharSet <> "" Then .CharSet = CharSet
        .open
        .WriteText obj.xml
        .SaveToFile FilePath, 2 'adSaveCreateOverWrite
        .Close
    End With
    Exit Sub
errHandler:
    ErrorIn "ujXML.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
End Sub
'
Public Sub PrintWithFoxit(pPDFFile As String, pPrinter)
    On Error GoTo errHandler
Dim j As Integer
    strCommand = """" & oPC.LocalFolder & "\Executables\Foxit Reader.exe""" & " /t " & """" & pPDFFile & """  """ & pPrinter & """"
    If mDebugmodeOn Then
        MsgBox "In printWithFoxit: " & strCommand
        
    End If
    For j = 1 To iCopyCount
        F_7_AB_1_ShellAndWaitSimple strCommand, vbMinimizedFocus, 30000
    Next

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "zXML.PrintWithFoxit(pPDFFile,pPrinter)", Array(pPDFFile, pPrinter)
End Sub
Public Sub PrintWithXPDF(pPDFFile As String, pPrinter)
    On Error GoTo errHandler
Dim pdf As XpdfPrint.XpdfPrint

    Set pdf = New XpdfPrint.XpdfPrint
    pdf.loadFile pPDFFile
    pdf.printPDF3 pdf.printSetPrinterFlag Or pdf.printSetCopiesFlag, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, iCopyCount, _
    pPrinter, 0, 0
    Set pdf = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "zXML.PrintWithXPDF(pPDFFile,pPrinter)", Array(pPDFFile, pPrinter)
End Sub
Private Function PrintWithAdobeReader(pPDFFile As String, pPrinter As String) As Boolean
    On Error GoTo errHandler
Dim strOldPrinter As String
Dim i As Integer

    PrintWithAdobeReader = True
    strOldPrinter = Printer.DeviceName
    If mDebugmodeOn = True Then
        MsgBox "Old default printer:" & strOldPrinter
    End If
        
    
    SetDefaultPrinter pPrinter
    
    If mDebugmodeOn = True Then
        MsgBox "New default printer:" & Printer.DeviceName
    End If
    
    Res = Forms(0).AcroPDF.loadFile(pPDFFile)
    If Res = 0 Then
      LogSaveToFile "Dispatcher:zXML:PrintWithAdobeReader - LOADFILE fails - " & FNS(pPDFFile)
    Else
        If mDebugmodeOn = True Then LogSaveToFile "Print copy count = " & CStr(iCopyCount)
        For i = 1 To iCopyCount
            Forms(0).AcroPDF.printAll
        Next
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "zXML.PrintWithAdobeReader(pPDFFile,pPrinter)", Array(pPDFFile, pPrinter)
End Function
Public Function PrintWithGhostScript() As Boolean
                    'strPrinter = IdentifyPrinter(strPrinter)
                   ' strCommand = strGSPrintLocation & "\gsprint.exe -noquery -printer """ & strPrinter & """ -copies " & CStr(iCopyCount) & " " & strPDFFile
                   ' Shell strCommand

End Function

Private Sub SetupFoldersandDefaultFilenames(pTemplateFolder As String, pWorkingFolder As String)
    If pTemplateFolder = "" Then
        pTemplateFolder = strSharedServerFolder & "\Templates\"
    End If
    If pWorkingFolder = "" Then
        pWorkingFolder = strSharedServerFolder & "\TEMP"
    End If
    If pWorkingFolder > "" Then
        strFOFile = pWorkingFolder & "\tmp" & ".FO"
        strPDFFile = pWorkingFolder & "\tmp" & ".PDF"
    End If
    If strExecutablesFolder = "" Then
        strExecutablesFolder = strSharedServerFolder & "\Executables"
    End If
    
    
End Sub

Public Function CreateFOUsingMSXML() As Boolean
    On Error GoTo errHandler
    
    CreateFOUsingMSXML = True
    
    Set oXSL = Nothing
    Set oXSL = New MSXML2.DOMDocument30
    oXSL.async = False
    oXSL.validateOnParse = False
    oXSL.resolveExternals = False
    If fs.FileExists(strTemplatePath) Then
        Res = oXSL.Load(strTemplatePath)
    Else
        LogSaveToFile "CreateFOUsingMSXML: strTemplatePath does not exist (" & strTemplatePath & ")   res:" & CStr(Res)
        CreateFOUsingMSXML = False
        Exit Function
    End If
    Set oXMLFO = New MSXML2.DOMDocument30
    oXMLFO.async = False
    oXMLFO.validateOnParse = False
    oXMLFO.resolveExternals = False
    xmlFile.docObject.transformNodeToObject oXSL, oXMLFO
    docWriteTostream strFOFile, oXMLFO, "UNICODE"
    If Not fs.FileExists(strFOFile) Then
        LogSaveToFile "CreateFOUsingMSXML" & strFOFile & " has not been created"
    End If
            
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "zXML.CreateFOUsingMSXML"
End Function
Public Function ConvertFOtoPDF() As Boolean
    On Error GoTo errHandler
Dim Res As Boolean
    ConvertFOtoPDF = True
    
        ChDir "\PBKS\Executables\FOP\"
    
    strCommand = GetFOPCommandstring(strFOFile, strPDFFile)
    If mDebugmodeOn = True Then
        LogSaveToFile strCommand
        MsgBox " in ConvertFOtoPDF, strCommand = " & strCommand
        
    End If
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand, vbMinimizedFocus, 180000)
    
    If fs.FileExists(strFOFile) Then
        fs.DeleteFile strFOFile
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "zXML.ConvertFOtoPDF"
End Function

Public Sub PrintWithSelectedTool(PDFTool As String)
10        On Error GoTo errHandler
20        If mDebugmodeOn Then
30            LogSaveToFile "tool: " & PDFTool & ", Printer:" & strPrinter & ", File:" & strPDFFile
40        End If
50        If PDFTool = "X" Then
60            PrintWithXPDF strPDFFile, IdentifyPrinter(strPrinter)
70        Else
80            If PDFTool = "G" Then
90                PrintWithGhostScript
100           Else
110               If PDFTool = "F" Then
120                   PrintWithFoxit strPDFFile, IdentifyPrinter(strPrinter)
130               Else
140                       strPrinter = IdentifyPrinter(strPrinter)
                        ' strExecutable = GetPDFExecutable(strPDFFile)
150                       PrintWithAdobeReader strPDFFile, strPrinter
160               End If
170           End If
180       End If

190       Exit Sub
errHandler:
200       If ErrMustStop Then Debug.Assert False: Resume
210       ErrorIn "zXML.PrintWithSelectedTool(PDFTool)", PDFTool
End Sub

