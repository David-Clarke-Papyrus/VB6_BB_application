VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Printing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim iFilenum1 As Integer
Dim iISBN As Integer
Dim iTitle As Integer
Dim iRef As Integer
Dim iQty As Integer
Dim iPrice As Integer
Dim iDisc As Integer
Dim iDep As Integer
Dim iExt As Integer
Dim iEXVAT As Integer
Dim iCurRow As Integer
Dim iCurTotalRow As Integer
Dim bHideDiscountColumn As Boolean
Dim bHideReferenceColumn As Boolean
Dim bHideDepositColumn As Boolean
Dim bHideQtyColumn As Boolean
Dim rptDEL As arDEL_1
Dim rptPO As Object
Dim rptR As Object
Dim rptCO As arCO_1
Dim rptAP As Object
Dim rptAR As arAPPR_1
Dim rptAPS As arApproSlips
Dim rptCN As arCN
Dim rptTF As arApproSlips
Dim rptINV As Object
Dim LineArray() As String
Dim LineTotalArray() As String
Dim iTitleWidth As Integer
Dim iAuthorWidth As Integer
Dim iPublisherWidth As Integer
Dim lngLineCount As Long
Dim lngLineTotalCount As Long
Dim strTotalDescription As String
Dim strPrintingDevice As String
Dim p As Printer
Dim strTemplateFolder As String
Dim lngCopies As Long
Dim strStyle As String
Dim strPrintPreview As String
Dim strPrinter As String


Event STATUS(pMsg As String)

Public WithEvents xT As XTimer
Attribute xT.VB_VarHelpID = -1
Dim arToPRint() As String
Dim strLogoPath As String

'Controlling printing
Dim mbKeepCopies As Boolean
Dim mbPreview As Boolean

Public Property Let KeepCopies(Val As Boolean)
    mbKeepCopies = Val
End Property
Public Property Let Preview(Val As Boolean)
    mbPreview = Val
End Property

Public Sub EnableTimer(bOn As Boolean)
    xT.Enabled = bOn
End Sub

Private Sub xT_Tick()
    xT.Enabled = False
    Forms(0).Enabled = False
    
    GetFilesToPRint
    PrintAllWaiting
    ClearPrintedFiles
    RaiseEvent STATUS("COMPLETE")
    
    xT.Enabled = True
    Forms(0).Enabled = True
    
    Exit Sub

End Sub
Public Property Get TimerEnabled() As Boolean
    TimerEnabled = xT.Enabled
End Property
Sub PrintAllWaiting()
    On Error GoTo errHandler
Dim i As Integer
Dim strType As String
Dim strMode As String
Dim oXML As zXML

    i = 1
    Do While i <= UBound(arToPRint, 1)
        strType = arToPRint(i, 2)
        strMode = arToPRint(i, 3)
        Select Case strType
        Case "PR1"
          ''      Builddocument_INVCustom arToPRint(i, 1), strLogoPath, True
        Case "IN", "IN_", "QU_"
            If UCase(strMode) = "XML" Then
                Set oXML = New zXML
                oXML.PrintXML arToPRint(i, 1), , , , strPDFPrintTool
            Else
          ''      Builddocument_INVCustom arToPRint(i, 1), strLogoPath, False
            End If
        Case "CN1", "CN2"
            If UCase(strMode) = "XML" Then
                Set oXML = New zXML
                oXML.PrintXML arToPRint(i, 1), , , , strPDFPrintTool
            Else
           ''     Builddocument_CNCustom arToPRint(i, 1), strLogoPath
            End If
        Case "DEL", "GRN"
            If UCase(strMode) = "XML" Then
                Set oXML = New zXML
                oXML.PrintXML arToPRint(i, 1), , , , strPDFPrintTool
            Else
           ''     Builddocument_DEL arToPRint(i, 1)
            End If
        Case "APS"
            Builddocument_APS arToPRint(i, 1)
        Case "PO_"
                Set oXML = New zXML
                oXML.PrintXML arToPRint(i, 1), , , , strPDFPrintTool
        Case "CO_"
            If UCase(strMode) = "XML" Then
                Set oXML = New zXML
                oXML.PrintXML arToPRint(i, 1), , , , strPDFPrintTool
            Else
            ''    Builddocument_CO arToPRint(i, 1)
            End If
        Case "AP_", "APP"
            If UCase(strMode) = "XML" Then
                Set oXML = New zXML
                oXML.PrintXML arToPRint(i, 1), , , , strPDFPrintTool
            Else
            ''    Builddocument_AP arToPRint(i, 1), mApproLogoFilename
            End If
        Case "APR"
            Builddocument_AR arToPRint(i, 1)
        Case "TF_"
            If UCase(strMode) = "XML" Then
                Set oXML = New zXML
                oXML.PrintXML arToPRint(i, 1), , , , strPDFPrintTool
            Else
            ''    Builddocument_TF arToPRint(i, 1)
            End If
        Case "RT_"
            If UCase(strMode) = "XML" Then
                Set oXML = New zXML
                oXML.PrintXML arToPRint(i, 1), , , , strPDFPrintTool
            Else
            ''    Builddocument_R arToPRint(i, 1), mApproLogoFilename
            End If
        End Select
        i = i + 1
    Loop
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.PrintAllWaiting", , EA_NORERAISE
    LogError
End Sub
Function ClearPrintedFiles() As Boolean
    On Error GoTo errHandler
Dim i As Long
Dim iErrCount As Integer
Dim fs As New FileSystemObject
    i = 1
    iErrCount = 0
    Do While i <= UBound(arToPRint, 1)
        If mbKeepCopies Then
            fs.CopyFile arToPRint(i, 1), "C:\PBKS\BU\"
        End If
        fs.DeleteFile arToPRint(i, 1), True
        i = i + 1
    Loop
    If fs.FileExists(strSharedServerFolder & "\Printing\*.PDF") Then
        fs.DeleteFile strSharedServerFolder & "\Printing\*.PDF"
    End If
    ClearPrintedFiles = True
    Exit Function
errHandler:
    ErrPreserve
    If Err = 70 And iErrCount < 3 Then
        For i = 1 To 10000000
            i = i + 1
        Next
        iErrCount = iErrCount + 1
        Resume
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ClearPrintedFiles"
End Function
Sub GetFilesToPRint()
    On Error GoTo errHandler
Dim strFileName As String
Dim i As Integer
Dim iMax As Integer
Dim strTag As String
Dim arTemp(500, 3) As String
Dim iFilenum1 As Integer
Dim fs As New FileSystemObject
Dim f1
Dim folder
Dim fc
    Set folder = fs.GetFolder(strSharedServerFolder & "\Printing")
    Set fc = folder.files
    i = 1
    For Each f1 In fc
        iFilenum1 = FreeFile
        Open f1.Path For Input As #iFilenum1
        strTag = ""
        On Error Resume Next
        Line Input #iFilenum1, strTag
        If Err = 0 Then
            strTag = Left(f1.Name, 3)
            If strTag = "QU_" Or strTag = "IN1" Or strTag = "PR1" Or strTag = "GRN" Or strTag = "APS" Or strTag = "IN_" Or strTag = "TF_" Or strTag = "CN1" Or strTag = "CN2" Or strTag = "CO_" Or strTag = "DEL" Or strTag = "PO_" Or strTag = "CS_" Or strTag = "AP_" Or strTag = "APP" Or strTag = "APR" Or strTag = "RT_" Then
                arTemp(i, 1) = f1.Path
                arTemp(i, 2) = strTag
                arTemp(i, 3) = Right(f1.Name, Len(f1.Name) - InStr(1, f1.Name, "."))
                i = i + 1
            End If
            Close #iFilenum1
        Else
            Close #iFilenum1
            fs.DeleteFile f1.Path, True

        End If
    Next
    On Error GoTo errHandler
    iMax = i - 1
    ReDim arToPRint(iMax, 3)
    For i = 1 To iMax
        arToPRint(i, 1) = arTemp(i, 1)
        arToPRint(i, 2) = arTemp(i, 2)
        arToPRint(i, 3) = arTemp(i, 3)
    Next

    Exit Sub
errHandler:
    ErrPreserve
    If Err = 62 Then
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.GetFilesToPRint"
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set xT = New XTimer
    xT.Interval = 5000
    xT.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Printing.Class_Initialize"
End Sub



Private Sub FillatBookmark(pBM As String, pVal As String)
    On Error GoTo errHandler
 '   wm.FillBM pBM, pVal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.FillatBookmark(pBM,pVal)", Array(pBM, pVal)
End Sub
Private Sub ReplaceMarker(pMarker As String, pVal As String)
    On Error GoTo errHandler
'    wm.ReplaceMarker pMarker, pVal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.ReplaceMarker(pMarker,pVal)", Array(pMarker, pVal)
End Sub

'Public Sub Builddocument_DEL(pFilename As String)
'    On Error GoTo errHandler
'Dim strIP As String
'Dim ar() As String
'
'
'    iFilenum1 = FreeFile
'    Open pFilename For Input As #iFilenum1
'    iCurRow = 1
'    iCurTotalRow = 1
'
'    Line Input #iFilenum1, strIP
'    ar = Split(strIP, "|")
'
'    strPrinter = ar(1)
'    lngCopies = ar(4)
'    strStyle = ar(3)
'    strPrintPreview = ar(2)
'
''    If Not CheckPrinterExists(strPrinter, strPrinter) Then
''        strPrinter = Printer.DeviceName
''    End If
'
'    If strPrinter = "" Then
'        Close #iFilenum1
'        Exit Sub
'    End If
'
'    Select Case strStyle
'    Case "A"
'        Set rptDEL = New arDEL_1
'    End Select
'     ReDim LineArray(0)
'
'    Do Until EOF(iFilenum1)
'        Line Input #iFilenum1, strIP
'        ParseLine_DEL strIP
'    Loop
'    rptDEL.Component LineArray, LineTotalArray
'    Close #iFilenum1
'    If mbPreview Then
'        rptDEL.Show vbModal
'    Else
'        rptDEL.Printer.DeviceName = IdentifyPrinter(strPrinter)
'        rptDEL.Printer.Copies = lngCopies
'        rptDEL.PrintReport False
'    End If
'    Set rptDEL = Nothing
'
'    Exit Sub
'errHandler:
'    ErrPreserve
'    If ErrMustStop Then Debug.Assert False: Resume
'    Set rptDEL = Nothing
'    CloseFiles iFilenum1
'    ErrorIn "ManageDocument.Builddocument_DEL(pFilename)", Array(pFilename)
'End Sub
'
Public Sub Builddocument_APS(pFilename As String)
    On Error GoTo errHandler
Dim strIP As String
Dim ar() As String

    iFilenum1 = FreeFile
    Open pFilename For Input As #iFilenum1
    iCurRow = 1
    iCurTotalRow = 1
    
    Line Input #iFilenum1, strIP
    ar = Split(strIP, "|")
    
    strPrinter = ar(1)
    lngCopies = ar(4)
    strStyle = ar(3)
    strPrintPreview = ar(2)
    
'    If Not CheckPrinterExists(strPrinter, strPrinter) Then
'        WritetoErrors "Printer does not exist (printing Appro slips): " & strPrinter & "Filename = " & pFilename
'        strPrinter = Printer.DeviceName
'    End If
    
    If strPrinter = "" Then
        Close #iFilenum1
        Exit Sub
    End If
    ReDim LineArray(0)
    
    Select Case strStyle
    Case "A"
        Set rptAPS = New arApproSlips
    End Select
    
    Do Until EOF(iFilenum1)
        Line Input #iFilenum1, strIP
        ParseLine_APS strIP
    Loop
    rptAPS.Component LineArray
    Close #iFilenum1
    
    If mbPreview Then
        rptAPS.Show vbModal
    Else
        MsgBox "POS1" & strPrinter
        rptAPS.Printer.DeviceName = IdentifyPrinter(strPrinter)
        MsgBox "POS2" & rptAPS.Printer.DeviceName
        rptAPS.Printer.Copies = lngCopies
        rptAPS.PrintReport False
    End If
    Set rptAPS = Nothing
    
    Exit Sub
errHandler:
    ErrPreserve
    Set rptAPS = Nothing
    If ErrMustStop Then Debug.Assert False: Resume
    CloseFiles iFilenum1
    ErrorIn "ManageDocument.Builddocument_APS(pFilename)", Array(pFilename)
End Sub

Private Function GetFullDeviceNameFromName(pIn As String) As String
Dim i As Integer
Dim r As Long
Dim buffer As String
Dim ar1() As String
    ' Get the list of available printers from WIN.INI
    buffer = Space(8192)
    r = GetProfileString("PrinterPorts", vbNullString, "", _
       buffer, Len(buffer))

    ' Display the list of printer in the ListBox List1
    ParseList ar1, buffer

End Function
Private Sub ParseList(ar() As String, ByVal buffer As String)
    Dim i As Integer
    Dim s As String

    Do
        i = InStr(buffer, Chr(0))
        If i > 0 Then
            s = Left(buffer, i - 1)
            If Len(Trim(s)) Then
                ReDim ar(1, UBound(ar, 0) + 1)
                ar(UBound(ar)) = s
            End If
            buffer = Mid(buffer, i + 1)
        Else
            If Len(Trim(buffer)) Then
                ReDim ar(1, UBound(ar, 1) + 1)
                ar(UBound(ar)) = s
            End If
            buffer = ""
        End If
    Loop While i > 0
End Sub

'Private Sub ParseLine_DEL(pIn As String)
'    On Error GoTo errHandler
'
'Dim ar() As String
'Dim ar1() As String
'Dim Tag As String
'Dim tmp1 As String
'
'    ar = Split(pIn, ">")
'    Tag = ar(0)
'    Select Case Tag
''    Case "<LINECOUNT"
''        lngLineCount = CLng(ar(1))
''        ReDim LineArray(lngLineCount)
'    Case "<LINETOTALCOUNT"
'        lngLineTotalCount = CLng(ar(1))
'        ReDim LineTotalArray(lngLineTotalCount)
'    Case "<DOCUMENTCODE"
'        rptDEL.Sections("PAGEHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
'    Case "<PHONE"
'        rptDEL.Sections("PAGEHEADER").Controls("PHONE") = CStr(ar(1))
'    Case "<SINVREF"
'        rptDEL.fSINVREF = CStr(ar(1))
'    Case "<SINVDATE"
'        rptDEL.fSINVDATE = CStr(ar(1))
'    Case "<FAX"
'        rptDEL.Sections("PAGEHEADER").Controls("FAX") = CStr(ar(1))
'    Case "<OURADDRESS"
'        rptDEL.Sections("PAGEHEADER").Controls("fOURADDRESS") = UnpackText(CStr(ar(1)))
'    Case "<SM"
'        rptDEL.Sections("PAGEHEADER").Controls("fSM") = CStr(ar(1))
'    Case "<STORECODE"
'        rptDEL.Sections("PAGEHEADER").Controls("fStoreCode") = CStr(ar(1))
'    Case "<DATE"
'        rptDEL.Sections("PAGEHEADER").Controls("DATE") = CStr(ar(1))
'    Case "<SUPPLIER"
'        rptDEL.Sections("PAGEHEADER").Controls("SUPPLIER") = CStr(ar(1))
'    Case "<COMPANYNAME"
'        rptDEL.Sections("PAGEHEADER").Controls("fCompanyName") = CStr(ar(1))
'    Case "<MEMO"
'        rptDEL.Sections("PAGEHEADER").Controls("fMEMO") = CStr(ar(1))
'    Case "<Line"
'        ReDim Preserve LineArray(UBound(LineArray) + 1)
'        LineArray(iCurRow) = ar(1)
'        iCurRow = iCurRow + 1
'    Case "<TOTAL"
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    End Select
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ManageDocument.ParseLine_DEL(pIn)", pIn
'End Sub
'
'Public Sub Builddocument_PO(pFilename As String, Optional pLogofilename As String)
'    On Error GoTo errHandler
'Dim strIP As String
'Dim ar() As String
'
'
'    iFilenum1 = FreeFile
'    Open pFilename For Input As #iFilenum1
'    iCurRow = 1
'    iCurTotalRow = 1
'
'    Line Input #iFilenum1, strIP
'    ar = Split(strIP, "|")
'
'    strPrinter = ar(1)
'    lngCopies = ar(4)
'    strStyle = ar(3)
'    strPrintPreview = ar(2)
'
''    If Not CheckPrinterExists(strPrinter, strPrinter) Then
''        WritetoErrors "Printer does not exist (printing Purchase order): " & strPrinter & vbCrLf & "Filename = " & pFilename
''        strPrinter = Printer.DeviceName
''    End If
'
'    If strPrinter = "" Then
'        Close #iFilenum1
'        Exit Sub
'    End If
'
'    ReDim LineArray(0)
'
'    Select Case strStyle
'    Case "A"
'        Set rptPO = New arPO_1
'        Do Until EOF(iFilenum1)
'            Line Input #iFilenum1, strIP
'            ParseLine_PO_A strIP
'        Loop
'    Case "B"
'        Set rptPO = New arPO_2
'        LoadFormats_PO
'        Do Until EOF(iFilenum1)
'            Line Input #iFilenum1, strIP
'            ParseLine_PO_B strIP
'        Loop
'    End Select
'
'    rptPO.Component LineArray, LineTotalArray, pLogofilename
'  '  rptPO..Orientation = ddOLandscape
'    Close #iFilenum1
'    If mbPreview Then
'        rptPO.Printer.Orientation = ddOLandscape
'        rptPO.Show vbModal
'    Else
'        rptPO.Printer.DeviceName = IdentifyPrinter(strPrinter)
'        rptPO.Printer.Orientation = ddOLandscape
'        rptPO.Printer.PaperSize = 9
'        rptPO.pageleftmargin = 870
'        rptPO.PageRightMargin = 0
'        rptPO.Printer.Copies = lngCopies
'        rptPO.PrintReport False
'    End If
'    Set rptPO = Nothing
'    Exit Sub
'errHandler:
'    ErrPreserve
'    If ErrMustStop Then Debug.Assert False: Resume
'    Set rptPO = Nothing
'    CloseFiles iFilenum1
'    ErrorIn "ManageDocument.Builddocument_PO(pFilename)", Array(pFilename)
'End Sub
'Public Sub Builddocument_R(pFilename As String, Optional pLogofilename As String)
'    On Error GoTo errHandler
'Dim strIP As String
'Dim ar() As String
'
'
'    iFilenum1 = FreeFile
'    Open pFilename For Input As #iFilenum1
'    iCurRow = 1
'    iCurTotalRow = 1
'
'    Line Input #iFilenum1, strIP
'    ar = Split(strIP, "|")
'
'    strPrinter = ar(1)
'    lngCopies = ar(4)
'    strStyle = ar(3)
'    strPrintPreview = ar(2)
'
'    If strPrinter = "" Then
'        Close #iFilenum1
'        Exit Sub
'    End If
'
'    Select Case strStyle
'    Case "A"
'        Set rptR = New arR_2
'    Case "B"
'        Set rptR = New arR_1
'    End Select
'    ReDim LineArray(0)
'    Do Until EOF(iFilenum1)
'        Line Input #iFilenum1, strIP
'        ParseLine_R strIP
'    Loop
'    rptR.Component LineArray, LineTotalArray, pLogofilename
'   ' rptR
'    Close #iFilenum1
'    If mbPreview Then
'        rptR.Show vbModal
'    Else
'        rptR.Printer.DeviceName = IdentifyPrinter(strPrinter)
'        rptR.Printer.PaperSize = 9
'        rptR.pageleftmargin = 800
'        rptR.Printer.Copies = lngCopies
'        rptR.PrintReport False
'    End If
'    Set rptR = Nothing
'    Exit Sub
'errHandler:
'    ErrPreserve
'    If ErrMustStop Then Debug.Assert False: Resume
'    Set rptTF = Nothing
'    CloseFiles iFilenum1
'    ErrorIn "ManageDocument.Builddocument_R(pFilename)", Array(pFilename)
'End Sub
'Private Sub ParseLine_PO_A(pIn As String)
'    On Error GoTo errHandler
'
'Dim ar() As String
'Dim ar1() As String
'Dim ar2() As String
'Dim Tag As String
'Dim tmp1 As String
'
'    ar = Split(pIn, ">")
'    Tag = ar(0)
'    Select Case Tag
'    Case "<LINETOTALCOUNT"
'        lngLineTotalCount = CLng(ar(1))
'        ReDim LineTotalArray(lngLineTotalCount)
'    Case "<DOCUMENTCODE"
'        rptPO.Sections("REPORTHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
'        On Error Resume Next
'        rptPO.Sections("REPORTHEADER").Controls("fDOCUMENTCODEFOOT") = CStr(ar(1))
'            On Error GoTo errHandler
'    Case "<SENDER"
'        rptPO.Sections("REPORTHEADER").Controls("fSENDER") = CStr(ar(1))
'    Case "<OURADDRESS"
'        rptPO.Sections("REPORTHEADER").Controls("fOURADDRESS") = UnpackText(ar(1))
'    Case "<PHONE"
'        rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") = rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") & "     " & CStr(ar(1))
'    Case "<STATUS"
'        rptPO.Sections("REPORTHEADER").Controls("fStatus") = CStr(ar(1))
'    Case "<BILLTO"
'        rptPO.Sections("REPORTHEADER").Controls("BILLTO") = UnpackText(ar(1))
'    Case "<DELTO"
'        rptPO.Sections("REPORTHEADER").Controls("DELTO") = UnpackText(ar(1))
'    Case "<FAX"
'        rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") = rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") & "     " & CStr(ar(1))
'    Case "<ACCNUMBER"
'        rptPO.fields("ACCNUMBER") = CStr(ar(1))
'    Case "<SM"
'        rptPO.Sections("REPORTHEADER").Controls("fSM") = CStr(ar(1))
'    Case "<DATE"
'        rptPO.Sections("REPORTHEADER").Controls("DATE") = CStr(ar(1))
'    Case "<FOOT"
'        rptPO.Sections("PAGEFOOTER").Controls("fFOOT") = CStr(ar(1))
'    Case "<MEMO"
'        rptPO.Sections("REPORTHEADER").Controls("fMemo") = UnpackText(CStr(ar(1)))
'    Case "<SUPPLIER"
'        rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") = CStr(ar(1))
'    Case "<SUPPLIERADDRESS"
'    'This is only on PO layout "B"
'        On Error Resume Next
'        rptPO.Sections("REPORTHEADER").Controls("fSUPPLIERAddress") = UnpackText(CStr(ar(1)))
'        Err.Clear
'        On Error GoTo errHandler
'
'    Case "<MAINMSG"
'      '  tmp1 = Replace(CStr(ar(1)), "%", Chr(13))
'        rptPO.Sections("REPORTFOOTER").Controls("fORDERTEXT") = UnpackText(ar(1)) 'CStr(ar(1))
'    Case "<Line"
'        ReDim Preserve LineArray(UBound(LineArray) + 1)
'        LineArray(iCurRow) = ar(1)
'        iCurRow = iCurRow + 1
'    Case "<TOTAL"
'        LineTotalArray(iCurTotalRow) = UnpackText(ar(1))    'ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    End Select
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ManageDocument.ParseLine_PO_A(pIn)", pIn
'End Sub
'
'Private Sub ParseLine_PO_B(pIn As String)
'    On Error GoTo errHandler
'
'Dim ar() As String
'Dim ar1() As String
'Dim ar2() As String
'Dim Tag As String
'Dim tmp1 As String
'
'    ar = Split(pIn, ">")
'    Tag = ar(0)
'    Select Case Tag
'    Case "<LINECOUNT"
''        lngLineCount = CLng(ar(1))
''        ReDim LineArray(lngLineCount)
'    Case "<LINETOTALCOUNT"
'        lngLineTotalCount = CLng(ar(1))
'        ReDim LineTotalArray(lngLineTotalCount)
'    Case "<DOCUMENTCODE"
'        rptPO.Sections("REPORTHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
'        rptPO.Sections("PAGEFOOTER").Controls("fDOCUMENTCODEFOOT") = CStr(ar(1))
'    Case "<SENDER"
'        rptPO.Sections("REPORTHEADER").Controls("fSENDER") = CStr(ar(1))
'    Case "<PHONE"
'        rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") = rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") & "     " & CStr(ar(1))
'    Case "<STATUS"
'        rptPO.Sections("REPORTHEADER").Controls("fStatus") = CStr(ar(1))
'    Case "<OURADDRESS"
'        rptPO.Sections("REPORTHEADER").Controls("fOURADDRESS") = UnpackText(ar(1))
'    Case "<BILLTO"
'        rptPO.Sections("REPORTHEADER").Controls("BILLTO") = UnpackText(ar(1))
' '   Case "<FAX"
' '       rptPO.Sections("REPORTHEADER").Controls("fFAX") = ar(1)
'    Case "<DELTO"
'      '  tmp1 = Replace(ar(1), "%", Chr(13))
'        rptPO.Sections("REPORTHEADER").Controls("DELTO") = UnpackText(ar(1))
'    Case "<FAX"
'        rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") = rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") & "     " & CStr(ar(1))
'    Case "<ACCNUMBER"
'        rptPO.fields("ACCNUMBER") = CStr(ar(1))
'    Case "<SM"
'        rptPO.Sections("REPORTHEADER").Controls("fSM") = CStr(ar(1))
'    Case "<DATE"
'        rptPO.Sections("REPORTHEADER").Controls("DATE") = CStr(ar(1))
'    Case "<FOOT"
'        rptPO.Sections("PAGEFOOTER").Controls("fFOOT") = CStr(ar(1))
'    Case "<MEMO"
'        rptPO.Sections("REPORTHEADER").Controls("fMemo") = UnpackText(CStr(ar(1)))
'    Case "<SUPPLIER"
'        rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") = CStr(ar(1))
'    Case "<SUPPLIERADDRESS"
'    'This is only on PO layout "B"
'        On Error Resume Next
'        rptPO.Sections("REPORTHEADER").Controls("fSUPPLIERAddress") = UnpackText(CStr(ar(1)))
'        Err.Clear
'        On Error GoTo errHandler
'
'    Case "<MAINMSG"
'      '  tmp1 = Replace(CStr(ar(1)), "%", Chr(13))
'        rptPO.Sections("REPORTFOOTER").Controls("fORDERTEXT") = UnpackText(ar(1)) 'CStr(ar(1))
'    Case "<Line"
'        ReDim Preserve LineArray(UBound(LineArray) + 1)
'        LineArray(iCurRow) = ar(1)
'        iCurRow = iCurRow + 1
'    Case "<TOTAL"
'        LineTotalArray(iCurTotalRow) = UnpackText(ar(1))    'ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    End Select
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ManageDocument.ParseLine_PO_B(pIn)", pIn
'End Sub
'Private Sub ParseLine_R(pIn As String)
'    On Error GoTo errHandler
'
'Dim ar() As String
'Dim ar1() As String
'Dim ar2() As String
'Dim Tag As String
'Dim tmp1 As String
'
'    ar = Split(pIn, ">")
'    Tag = ar(0)
'    Select Case Tag
''    Case "<LINECOUNT"
''        lngLineCount = CLng(ar(1))
' '       ReDim LineArray(lngLineCount)
'    Case "<LINETOTALCOUNT"
'        lngLineTotalCount = CLng(ar(1))
'        ReDim LineTotalArray(lngLineTotalCount)
'    Case "<DOCUMENTCODE"
'        rptR.Sections("REPORTHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
'        rptR.Sections("PAGEFOOTER").Controls("FDOCNO") = CStr(ar(1))
'    Case "<PHONE"
'        rptR.Sections("REPORTHEADER").Controls("SUPPLIER") = rptR.Sections("REPORTHEADER").Controls("SUPPLIER") & "     " & CStr(ar(1))
'    Case "<STATUS"
'        rptR.Sections("REPORTHEADER").Controls("fStatus") = CStr(ar(1))
'    Case "<BILLTO"
'     '   tmp1 = Replace(ar(1), "%", Chr(13))
'        rptR.Sections("REPORTHEADER").Controls("BILLTO") = UnpackText(ar(1))
'    Case "<DELTO"
''        tmp1 = Replace(ar(1), "%", Chr(13))
''        rptR.Sections("PAGEHEADER").Controls("DELTO") = tmp1
'    Case "<SENDER"
'        On Error Resume Next
'        rptR.Sections("REPORTHEADER").Controls("fSENDER") = CStr(ar(1))
'        On Error GoTo errHandler
'    Case "<STORECODE"
'  '      rptR.fStoreCode = CStr(ar(1))
'    Case "<OURADDRESS"
'        rptR.Sections("REPORTHEADER").Controls("fOURADDRESS") = UnpackText(ar(1))
'    Case "<FAX"
'        rptR.Sections("REPORTHEADER").Controls("SUPPLIER") = rptR.Sections("REPORTHEADER").Controls("SUPPLIER") & "     " & CStr(ar(1))
'    Case "<SM"
'        rptR.Sections("REPORTHEADER").Controls("fSM") = CStr(ar(1))
'    Case "<ACCNUMBER"
'        rptR.fields("ACCNUMBER") = CStr(ar(1))
'    Case "<DATE"
'        rptR.Sections("REPORTHEADER").Controls("DATE") = CStr(ar(1))
'    Case "<MEMO"
'        rptR.Sections("REPORTHEADER").Controls("fMemo") = UnpackText(CStr(ar(1)))
'    Case "<SUPPLIER"
'        rptR.Sections("REPORTHEADER").Controls("SUPPLIER") = CStr(ar(1))
'    Case "<MAINMSG"
'        tmp1 = "Authorization: " & Replace(CStr(ar(1)), "%", Chr(13))
'        rptR.Sections("REPORTHEADER").Controls("fAuthorization") = tmp1 'CStr(ar(1))
'    Case "<Line"
'        ReDim Preserve LineArray(UBound(LineArray) + 1)
'        LineArray(iCurRow) = ar(1)
'        iCurRow = iCurRow + 1
'    Case "<TOTAL"
'        LineTotalArray(iCurTotalRow) = UnpackText(ar(1))    'ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    Case "<QTYTOTAL"
'        ar1() = Split(CStr(ar(1)), "|")
'        rptR.Sections("REPORTFOOTER").Controls("fTotalQty") = ar1(0) & "  " & ar1(1)                     'CStr(ar(1))
'    End Select
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ManageDocument.ParseLine_R(pIn)", pIn
'End Sub
'
'Private Sub ParseLine_INV2(pIn As String)
'    On Error GoTo errHandler
'
'Dim ar() As String
'Dim ar1() As String
'Dim Tag As String
'Dim tmp1 As String
'On Error Resume Next
'    ar = Split(pIn, ">")
'    Tag = ar(0)
'    Select Case Tag
'    Case "<LINECOUNT"
'        lngLineCount = CLng(ar(1))
'        ReDim LineArray(lngLineCount)
'    Case "<LINETOTALCOUNT"
'        lngLineTotalCount = CLng(ar(1))
'        ReDim LineTotalArray(lngLineTotalCount)
''    Case "<DOCUMENTCODE"
''        rptINV.Sections("PAGEHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
'    Case "<PHONE"
'        rptINV.Sections("PAGEHEADER").Controls("PHONE") = CStr(ar(1))
'    Case "<OURADDRESS"
'        rptINV.Sections("PAGEHEADER").Controls("fStore") = UnpackText(ar(1))
'    Case "<STATUS"
'        rptINV.fields("fSTATUS") = CStr(ar(1))
'    Case "<DOCUMENTCODE"
'        rptINV.Sections("PAGEHEADER").Controls("lblDocument").Caption = "TAX INVOICE No. " & CStr(ar(1))
'    Case "<BILLTO"
'      '  tmp1 = Replace(ar(1), "%", Chr(13))
'        rptINV.Sections("PAGEHEADER").Controls("BILLTO") = UnpackText(ar(1))
'    Case "<DELTO"
'     '  tmp1 = Replace(ar(1), "%", Chr(13))
'        rptINV.Sections("PAGEHEADER").Controls("DELTO") = UnpackText(ar(1))
'    Case "<FAX"
'        rptINV.Sections("PAGEHEADER").Controls("FAX") = CStr(ar(1))
'    Case "<SM"
'        rptINV.Sections("PageFOOTER").Controls("lblSM") = CStr(ar(1))
''    Case "<ACCNUMBER"
''        rptINV.Sections("PAGEHEADER").Controls("CUSTOMER") = "To: " & CStr(ar(1))
'    Case "<DATE"
'        rptINV.Sections("PAGEHEADER").Controls("DATE") = CStr(ar(1))
'    Case "<CUSTOMER"
'        rptINV.Sections("PAGEHEADER").Controls("CUSTOMER") = "To: " & CStr(ar(1))
'    Case "<Line"
'        LineArray(iCurRow) = ar(1)
'        iCurRow = iCurRow + 1
'    Case "<TOTAL"
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    End Select
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ManageDocument.ParseLine_INV2(pIn)", pIn
'End Sub
'Private Sub ParseLine_INVCUSTOM(pIn As String, pbProforma As Boolean)
'    On Error GoTo errHandler
'
'Dim ar() As String
'Dim ar1() As String
'Dim Tag As String
'Dim tmp1 As String
'On Error Resume Next
'    ar = Split(pIn, ">")
'    Tag = ar(0)
'    Select Case Tag
''    Case "<LINECOUNT"
''        lngLineCount = CLng(ar(1))
''        ReDim LineArray(lngLineCount)
''    Case "<LINETOTALCOUNT"
''        lngLineTotalCount = CLng(ar(1))
''        ReDim LineTotalArray(lngLineTotalCount)
'    Case "<PHONE"
'        rptINV.Sections("REPORTHEADER").Controls("fPHONE") = CStr(ar(1))
'    Case "<COREGNUMBER"
'        rptINV.Sections("PAGEFOOTER").Controls("fVAT") = rptINV.Sections("PAGEFOOTER").Controls("fVAT") & "Company Reg. No. " & CStr(ar(1))
'    Case "<VATNUMBER"
'        rptINV.Sections("PAGEFOOTER").Controls("fVAT") = rptINV.Sections("PAGEFOOTER").Controls("fVAT") & vbCrLf & "V.A.T. No. " & CStr(ar(1))
'    Case "<CUSTVATNUMBER"
'        If CStr(ar(1)) > "" Then
'           ' rptINV.Sections("REPORTHEADER").Controls("fCUSTOMER") = rptINV.Sections("REPORTHEADER").Controls("fCUSTOMER") & vbCrLf & "V.A.T. No. " & CStr(ar(1))
'            rptINV.Sections("REPORTHEADER").Controls("fCustVATNUMBER") = "To V.A.T. No. " & CStr(ar(1))
'        End If
'    Case "<OURADDRESS"
'        rptINV.Sections("REPORTHEADER").Controls("fStore") = UnpackText(ar(1))
'    Case "<STATUS"
'        rptINV.Sections("REPORTHEADER").Controls("fSTATUS") = CStr(ar(1))
'        rptINV.Sections("REPORTHEADER").Controls("fSTATUS").Visible = (CStr(ar(1)) > "")
'
'    Case "<DOCUMENTCODE"
'        If pbProforma = True Then
'            rptINV.Sections("REPORTHEADER").Controls("lblDocument").Caption = "PRO-FORMA INVOICE No. " & CStr(ar(1))
'            rptINV.Sections("PAGEHEADER").Controls("lblDocumentPage").Caption = "PRO-FORMA INVOICE No. " & CStr(ar(1))
'        Else
'            rptINV.Sections("REPORTHEADER").Controls("lblDocument").Caption = "TAX INVOICE No. " & CStr(ar(1))
'            rptINV.Sections("PAGEHEADER").Controls("lblDocumentPage").Caption = "TAX INVOICE No. " & CStr(ar(1))
'        End If
'    Case "<BILLTO"
'        rptINV.Sections("REPORTHEADER").Controls("fBILLTO") = UnpackText(ar(1))
'    Case "<DELTO"
'        rptINV.Sections("REPORTHEADER").Controls("fDELTO") = UnpackText(ar(1))
'    Case "<FAX"
'        rptINV.Sections("REPORTHEADER").Controls("fFAX") = CStr(ar(1))
'    Case "<MEMO"
'        If rptINV.Sections("REPORTHEADER").Controls("fMEMOHEADER").Visible = True Then
'            rptINV.Sections("REPORTHEADER").Controls("fMEMOHEADER") = CStr(ar(1))
'        Else
'            rptINV.Sections("REPORTFOOTER").Controls("fMEMOFOOTER") = CStr(ar(1))
'        End If
'    Case "<SM"
'        rptINV.Sections("REPORTHEADER").Controls("fSM") = CStr(ar(1))
'    Case "<ACCNUMBER"
'        rptINV.Sections("REPORTHEADER").Controls("fACNO") = CStr(ar(1))
'    Case "<DATE"
'        rptINV.Sections("REPORTHEADER").Controls("fDATE") = CStr(ar(1))
'    Case "<CUSTOMER"
'        rptINV.Sections("REPORTHEADER").Controls("fCUSTOMER") = "To: " & CStr(ar(1))
'    Case "<BANK"
'        rptINV.Sections("PAGEFOOTER").Controls("fBank") = UnpackText(CStr(ar(1)))
'    Case "<CONAME"
'        rptINV.Sections("REPORTHEADER").Controls("fSENDER") = CStr(ar(1))
'    Case "<Line"
'        ReDim Preserve LineArray(UBound(LineArray) + 1)
'        LineArray(iCurRow) = ar(1)
'        iCurRow = iCurRow + 1
'    Case "<TOTAL"
'        ReDim Preserve LineTotalArray(UBound(LineTotalArray) + 1)
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    Case "<TOTAL2"
'        ReDim Preserve LineTotalArray(UBound(LineTotalArray) + 1)
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    Case "<TOTAL3"
'        ReDim Preserve LineTotalArray(UBound(LineTotalArray) + 1)
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    Case "<TOTAL4"
'        ReDim Preserve LineTotalArray(UBound(LineTotalArray) + 1)
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    Case "<HIDEDISCOUNTCOLUMN"
'        bHideDiscountColumn = True
'    Case "<HIDEREFERENCECOLUMN"
'        bHideReferenceColumn = True
'    Case "<HIDEDEPOSITCOLUMN"
'        bHideDepositColumn = True
'    Case "<HIDEQTYCOLUMN"
'        bHideQtyColumn = True
'    End Select
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ManageDocument.ParseLine_INVCustom(pIn)", pIn
'End Sub
'Private Sub ParseLine_CNCUSTOM(pIn As String)
'    On Error GoTo errHandler
'
'Dim ar() As String
'Dim ar1() As String
'Dim Tag As String
'Dim tmp1 As String
'On Error Resume Next
'    ar = Split(pIn, ">")
'    Tag = ar(0)
'    Select Case Tag
'    Case "<COREGNUMBER"
'        rptCN.Sections("PAGEFOOTER").Controls("fVAT") = rptCN.Sections("PAGEFOOTER").Controls("fVAT") & vbCrLf & "Comany Reg. No. " & CStr(ar(1))
'    Case "<VATNUMBER"
'        rptCN.Sections("PAGEFOOTER").Controls("fVAT") = rptCN.Sections("PAGEFOOTER").Controls("fVAT") & vbCrLf & "V.A.T. No. " & CStr(ar(1))
'    Case "<CUSTVATNUMBER"
'        If CStr(ar(1)) > "" Then
''            rptCN.Sections("REPORTHEADER").Controls("fVATNUMBER") = "V.A.T. No. " & CStr(ar(1))
'            rptCN.Sections("REPORTHEADER").Controls("fCUSTOMER") = rptCN.Sections("REPORTHEADER").Controls("fCUSTOMER") & vbCrLf & "V.A.T. No. " & CStr(ar(1))
''            rptCN.Sections("REPORTHEADER").Controls("fDELTO") = rptCN.Sections("REPORTHEADER").Controls("fDELTO") & vbCrLf & "V.A.T. No. " & CStr(ar(1))
'        End If
'    Case "<PHONE"
'        rptCN.Sections("REPORTHEADER").Controls("fPHONE") = CStr(ar(1))
'    Case "<OURADDRESS"
'        rptCN.Sections("REPORTHEADER").Controls("fStore") = UnpackText(ar(1))
'    Case "<STATUS"
'    Case "<DOCUMENTCODE"
'        rptCN.Sections("REPORTHEADER").Controls("lblDocument").Caption = "TAX CREDIT NOTE No. " & CStr(ar(1))
'        rptCN.Sections("PAGEHEADER").Controls("lblDocumentPage").Caption = "TAX CREDIT NOTE No. " & CStr(ar(1))
'    Case "<BILLTO"
'        rptCN.Sections("REPORTHEADER").Controls("fBILLTO") = UnpackText(ar(1))
'    Case "<DELTO"
'        rptCN.Sections("REPORTHEADER").Controls("fDELTO") = UnpackText(ar(1))
'    Case "<FAX"
'        rptCN.Sections("REPORTHEADER").Controls("fFAX") = CStr(ar(1))
'    Case "<MEMO"
'        If rptCN.Sections("REPORTHEADER").Controls("fMEMOHEADER").Visible = True Then
'            rptCN.Sections("REPORTHEADER").Controls("fMEMOHEADER") = CStr(ar(1))
'        Else
'            rptCN.Sections("REPORTFOOTER").Controls("fMEMOFOOTER") = CStr(ar(1))
'        End If
'    Case "<SM"
'        rptCN.Sections("REPORTHEADER").Controls("fSM") = CStr(ar(1))
'    Case "<ACCNUMBER"
'        rptCN.Sections("REPORTHEADER").Controls("fACNO") = CStr(ar(1))
'    Case "<DATE"
'        rptCN.Sections("REPORTHEADER").Controls("fDATE") = CStr(ar(1))
'    Case "<CUSTOMER"
'        rptCN.Sections("REPORTHEADER").Controls("fCUSTOMER") = "To: " & CStr(ar(1))
'    Case "<CONAME"
'        rptCN.Sections("REPORTHEADER").Controls("fSENDER") = CStr(ar(1))
'    Case "<BANK"
'        rptCN.Sections("PAGEFOOTER").Controls("fBank") = UnpackText(CStr(ar(1)))
'    Case "<Line"
'        ReDim Preserve LineArray(UBound(LineArray) + 1)
'        LineArray(iCurRow) = ar(1)
'        iCurRow = iCurRow + 1
'    Case "<TOTAL"
'        ReDim Preserve LineTotalArray(UBound(LineTotalArray) + 1)
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    Case "<TOTAL2"
'        ReDim Preserve LineTotalArray(UBound(LineTotalArray) + 1)
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    Case "<TOTAL3"
'        ReDim Preserve LineTotalArray(UBound(LineTotalArray) + 1)
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    Case "<TOTAL4"
'        ReDim Preserve LineTotalArray(UBound(LineTotalArray) + 1)
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    Case "<HIDEDISCOUNTCOLUMN"
'        bHideDiscountColumn = True
'    Case "<HIDEREFERENCECOLUMN"
'        bHideReferenceColumn = True
'    Case "<HIDEDEPOSITCOLUMN"
'        bHideDepositColumn = True
'    Case "<HIDEQTYCOLUMN"
'        bHideQtyColumn = True
'    End Select
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ManageDocument.ParseLine_CNCustom(pIn)", pIn
'End Sub
'
''''''''''''
'Public Sub Builddocument_CO(pFilename As String)
'    On Error GoTo errHandler
'Dim strIP As String
'Dim ar() As String
'
'    iFilenum1 = FreeFile
'    Open pFilename For Input As #iFilenum1
'    iCurRow = 1
'    iCurTotalRow = 1
'    Line Input #iFilenum1, strIP
'
'    ar = Split(strIP, "|")
'
'    strPrinter = ar(1)
'    lngCopies = ar(4)
'    strStyle = ar(3)
'    strPrintPreview = ar(2)
'
''    If Not CheckPrinterExists(strPrinter, strPrinter) Then
''        WritetoErrors "Printer does not exist (printing order): " & strPrinter & ". Using activeprinter: " & "Filename = " & pFilename
''        strPrinter = Printer.DeviceName
''    End If
'
'    If strPrinter = "" Then
'        Close #iFilenum1
'        Exit Sub
'    End If
'    ReDim LineArray(0)
'
'    Select Case strStyle
'    Case "A"
'        Set rptCO = New arCO_1
'    End Select
'
'    Do Until EOF(iFilenum1)
'        Line Input #iFilenum1, strIP
'        ParseLine_CO strIP
'    Loop
'
'    rptCO.Component LineArray, LineTotalArray
'    Close #iFilenum1
'    If mbPreview Then
'        rptCO.Show vbModal
'    Else
'        rptCO.Printer.DeviceName = IdentifyPrinter(strPrinter)
'        rptCO.Printer.Copies = lngCopies
'        rptCO.PrintReport False
'    End If
'    Set rptCO = Nothing
'    Exit Sub
'
'errHandler:
'    ErrPreserve
'    Set rptCO = Nothing
'    If ErrMustStop Then Debug.Assert False: Resume
'    CloseFiles iFilenum1
'    ErrorIn "ManageDocument.Builddocument_CO(pFilename)", Array(pFilename)
'End Sub
'Private Sub ParseLine_CO(pIn As String)
'    On Error GoTo errHandler
'
'Dim ar() As String
'Dim ar1() As String
'Dim Tag As String
'Dim tmp1 As String
'
'    ar = Split(pIn, ">")
'    Tag = ar(0)
'    Select Case Tag
''    Case "<LINECOUNT"
''        lngLineCount = CLng(ar(1))
''        ReDim LineArray(lngLineCount)
'    Case "<LINETOTALCOUNT"
'        lngLineTotalCount = CLng(ar(1))
'        ReDim LineTotalArray(lngLineTotalCount)
'    Case "<DOCUMENTCODE"
'        rptCO.Sections("PAGEHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
'    Case "<PHONE"
'        rptCO.Sections("PAGEHEADER").Controls("PHONE") = CStr(ar(1))
'    Case "<STATUS"
'      '  rptPO.Fields("STATUS") = CStr(ar(1))
'    Case "<BILLTO"
'   '     tmp1 = Replace(ar(1), "%", Chr(13))
'        rptCO.Sections("PAGEHEADER").Controls("BILLTO") = UnpackText(ar(1))
'    Case "<DELTO"
'      '  tmp1 = Replace(ar(1), "%", Chr(13))
'        rptCO.Sections("PAGEHEADER").Controls("DELTO") = UnpackText(ar(1))
'    Case "<STORECODE"
'        rptCO.Sections("PAGEFOOTER").Controls("fStoreCode") = CStr(ar(1))
'        'rptCO.fStoreCode = CStr(ar(1))
'    Case "<FAX"
'        rptCO.Sections("PAGEHEADER").Controls("FAX") = CStr(ar(1))
'    Case "<ACCNUMBER"
'      '  rptCO.Fields("ACCNUMBER") = CStr(ar(1))
'    Case "<SM"
'        rptCO.fSM = CStr(ar(1))
'    Case "<MEMO"
'        rptCO.fMemo = "Note: " & UnpackText(CStr(ar(1)))
'    Case "<DATE"
'        rptCO.Sections("PAGEHEADER").Controls("DATE") = CStr(ar(1))
'    Case "<CUSTOMER"
'        rptCO.Sections("PAGEHEADER").Controls("CUSTOMER") = CStr(ar(1))
'    Case "<Line"
'        ReDim Preserve LineArray(UBound(LineArray) + 1)
'        LineArray(iCurRow) = ar(1)
'        iCurRow = iCurRow + 1
'    Case "<TOTAL"
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    End Select
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ManageDocument.ParseLine_CO(pIn)", pIn
'End Sub

'Public Sub Builddocument_AP(pFilename As String, Optional pLogofilename As String)
'    On Error GoTo errHandler
'Dim strIP As String
'Dim ar() As String
'
'    iFilenum1 = FreeFile
'    Open pFilename For Input As #iFilenum1
'    iCurRow = 1
'    iCurTotalRow = 1
'
'    Line Input #iFilenum1, strIP
'    ar = Split(strIP, "|")
'
'    strPrinter = ar(1)
'    lngCopies = ar(4)
'    strStyle = ar(3)
'    strPrintPreview = ar(2)
'
''    If Not CheckPrinterExists(strPrinter, strPrinter) Then
''        WritetoErrors "Printer does not exist (printing Appro): " & strPrinter & "Filename = " & pFilename
''        strPrinter = Printer.DeviceName
''    End If
'
'    If strPrinter = "" Then
'        Close #iFilenum1
'        Exit Sub
'    End If
'    ReDim LineArray(0)
'    ReDim LineTotalArray(0)
'
'    Select Case strStyle
'    Case "A"
'        Set rptAP = New arAPP_1
'    Case "B"
'        Set rptAP = New arAPP_2
'    End Select
'
'    Do Until EOF(iFilenum1)
'        Line Input #iFilenum1, strIP
'        ParseLine_AP strIP
'    Loop
'
'    rptAP.Component LineArray, LineTotalArray, pLogofilename
'    Close #iFilenum1
'
'    If mbPreview Then
'        rptAP.Show vbModal
'    Else
'            rptAP.Printer.DeviceName = IdentifyPrinter(strPrinter)
'            rptAP.Printer.Orientation = ddOLandscape
'            rptAP.Printer.PaperSize = 9
'            rptAP.pageleftmargin = 800
'            rptAP.Printer.Copies = lngCopies
'            rptAP.PrintReport False
'    End If
'    Set rptAP = Nothing
'
'    Exit Sub
'errHandler:
'    ErrPreserve
'    If ErrMustStop Then Debug.Assert False: Resume
'    Set rptAP = Nothing
'    CloseFiles iFilenum1
'    ErrorIn "ManageDocument.Builddocument_AP(pFilename)", Array(pFilename)
'End Sub
'Private Sub ParseLine_AP(pIn As String)
'    On Error GoTo errHandler
'
'Dim ar() As String
'Dim ar1() As String
'Dim Tag As String
'Dim tmp1 As String
'
'    ar = Split(pIn, ">")
'    Tag = ar(0)
'    Select Case Tag
''    Case "<LINECOUNT"
''        lngLineCount = CLng(ar(1))
''        ReDim LineArray(lngLineCount)
''    Case "<LINETOTALCOUNT"
''        lngLineTotalCount = CLng(ar(1))
''        ReDim LineTotalArray(lngLineTotalCount)
'    Case "<DOCUMENTCODE"
'        rptAP.Sections("REPORTHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
'        rptAP.Sections("PAGEFOOTER").Controls("fDOCNO") = CStr(ar(1))
'    Case "<COMPANYNAME"
'        rptAP.Sections("REPORTHEADER").Controls("FFROM") = UnpackText(CStr(ar(1)))
'    Case "<PHONE"
'        rptAP.Sections("REPORTHEADER").Controls("PHONE") = CStr(ar(1))
'    Case "<SENDER"
'        On Error Resume Next
'        rptAP.Sections("REPORTHEADER").Controls("fSENDER") = CStr(ar(1))
'        On Error GoTo errHandler
'    Case "<OURADDRESS"
'        rptAP.Sections("REPORTHEADER").Controls("fOURADDRESS") = UnpackText(ar(1))
'    Case "<STATUS"
'      '  rptPO.Fields("STATUS") = CStr(ar(1))
'    Case "<MEMO"
'        rptAP.Sections("REPORTHEADER").Controls("fMEMO") = CStr(ar(1))
'    Case "<STORECODE"
'        rptAP.Sections("PAGEFOOTER").Controls("fStoreCode") = CStr(ar(1))
'    Case "<BILLTO"
'     '  tmp1 = Replace(ar(1), "%", Chr(13))
'        rptAP.Sections("REPORTHEADER").Controls("BILLTO") = UnpackText(ar(1))
'    Case "<DELTO"
'     '   tmp1 = Replace(ar(1), "%", Chr(13))
'        rptAP.Sections("REPORTHEADER").Controls("DELTO") = UnpackText(ar(1))
'    Case "<FAX"
'        rptAP.Sections("REPORTHEADER").Controls("FAX") = CStr(ar(1))
'    Case "<SM"
'        rptAP.Sections("REPORTHEADER").Controls("fSM") = CStr(ar(1))
'    Case "<ACCNUMBER"
'      '  rptCO.Fields("ACCNUMBER") = CStr(ar(1))
'    Case "<DATE"
'        rptAP.Sections("REPORTHEADER").Controls("DATE") = CStr(ar(1))
'    Case "<CUSTOMER"
'        rptAP.Sections("REPORTHEADER").Controls("CUSTOMER") = CStr(ar(1))
'    Case "<Line"
'        ReDim Preserve LineArray(UBound(LineArray) + 1)
'        LineArray(iCurRow) = ar(1)
'        iCurRow = iCurRow + 1
'    Case "<TOTAL"
'        ReDim Preserve LineTotalArray(UBound(LineTotalArray) + 1)
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    End Select
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ManageDocument.ParseLine_AP(pIn)", pIn
'End Sub
'
Private Sub ParseLine_APS(pIn As String)
    On Error GoTo errHandler

Dim ar() As String
Dim ar1() As String
Dim Tag As String
Dim tmp1 As String

    ar = Split(pIn, ">")
    Tag = ar(0)
    Select Case Tag
    Case "<LINECOUNT"
'        lngLineCount = CLng(ar(1))
'        ReDim LineArray(lngLineCount)
    Case "<Line"
        ReDim Preserve LineArray(UBound(LineArray) + 1)
        LineArray(iCurRow) = ar(1)
        iCurRow = iCurRow + 1
    End Select
'ERRH:
'    MsgBox "Managedocument:ParseLine" & Error
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.ParseLine_APS(pIn)", pIn
End Sub



'Public Sub Builddocument_CN(pFilename As String, pTemplate As String, pstrLogoPath As String)
'    On Error GoTo errHandler
'Dim strIP As String
'Dim i As Integer
'Dim ar() As String
'Dim strPos As String
'Dim strFullDeviceName As String
'
''    bHideDiscountColumn = False
''''    Set oDoc = wm.StartNewDocument(pTemplate)
'    iFilenum1 = FreeFile
'    Open pFilename For Input As #iFilenum1
'    Line Input #iFilenum1, strIP
'    ar = Split(strIP, "|")
'    strFullDeviceName = ar(1)
'
'    If Not CheckPrinterExists(strFullDeviceName) Then
'        WritetoErrors "Printer does not exist (printing Credit note): " & strFullDeviceName & vbCrLf & "Filename = " & pFilename
'   '''     strFullDeviceName = wm.WORDApplication.ActivePrinter
'    End If
'
'    If strFullDeviceName = "" Then
'  '''      oDoc.Close 0
'        Close #iFilenum1
'        Exit Sub
'    End If
'
'    strPos = "Position 1"
'
'' '   FindColumnPositionInTable iISBN, iTitle, iRef, iQty, iDisc, iDep, iPrice, iExt, iEXVAT
''    If bHideDiscountColumn Then
''        oDoc.Tables(1).Columns(iTitle).Width = oDoc.Tables(1).Columns(iTitle).Width + oDoc.Tables(1).Columns(iDisc).Width
''        oDoc.Tables(1).Columns(iDisc).Delete
''    End If
' ''''   FindColumnPositionInTable iISBN, iTitle, iRef, iQty, iDisc, iDep, iPrice, iExt, iEXVAT
'
'    strPos = "Position 2"
'
'    iCurRow = 2
'    iCurTotalRow = 1
'    Do Until EOF(iFilenum1)
'        Line Input #iFilenum1, strIP
'        ParseLine_CN strIP, pstrLogoPath
'    Loop
'    strPos = "Position 2.B"
''    oDoc.Tables(1).Rows(oDoc.Tables(1).Rows.Count).Delete
''    oDoc.Tables(2).Rows(oDoc.Tables(2).Rows.Count).Delete
''    'format tables
''    oDoc.Tables(1).Rows(1).Borders(3).LineStyle = 7 'wdLineStyleDouble   'wdBorderBottom
''    oDoc.Tables(1).Rows(1).Borders(3).LineWidth = 2 'wdLineWidth025pt
''    oDoc.Tables(1).Rows(oDoc.Tables(1).Rows.Count).Borders(3).LineStyle = 13 'wdLineStyleThickThinMedGap
''    oDoc.Tables(1).Rows(oDoc.Tables(1).Rows.Count).Borders(3).LineWidth = 2 'wdLineWidth025pt
'    strPos = "Position 3"
''    FindColumnPositionInTable iISBN, iTitle, iRef, iQty, iDisc, iDep, iPrice, iExt, iEXVAT
''    If bHideReferenceColumn Then
''        oDoc.Tables(1).Columns(iRef - 1).Width = oDoc.Tables(1).Columns(iRef - 1).Width + oDoc.Tables(1).Columns(iRef).Width
''        oDoc.Tables(1).Columns(iRef).Delete '(oDoc.Tables(1).Columns(iDisc + 1).Cells)
''    End If
'    strPos = "Position 5"
'    ReplaceMarker "<STATUS>", ""
''    If wm.WordIsVisible Then
''        oDoc.PrintPreview
''        Forms(0).xT.Enabled = False
''        MsgBox "View document"
''        Forms(0).xT.Enabled = True
''    Else
''        strPos = "Position 6"
''        wm.WORDApplication.WordBasic.FILEPRINTSETUP Printer:=strFullDeviceName, DoNotSetAsSysDefault:=1
''        strPos = "Position 7"
''        oDoc.PrintOut Copies:=ar(4), Background:=False
''        strPos = "Position 9"
''    End If
'    strPos = "Position 10"
''    oDoc.Close 0
'    strPos = "Position 11"
'    Close #iFilenum1
'    Exit Sub
'errHandler:
'    ErrPreserve
''    If Not oDoc Is Nothing Then
''        On Error Resume Next
''        oDoc.Close 0
''    End If
'    If ErrMustStop Then Debug.Assert False: Resume
'    CloseFiles iFilenum1
'    ErrorIn "ManageDocument.Builddocument_CN(pFilename,pTemplate,pstrLogoPath)", _
'         Array(pFilename, pTemplate, pstrLogoPath), , , "Marker", strPos
'End Sub

'Private Sub ParseLine_CN(pIn As String, pLogoPath As String)
'    On Error GoTo errHandler
'Dim tmp As String
'Dim ar() As String
'Dim ar1() As String
'Dim ar2() As String
'Dim ar3() As String
'Dim Tag As String
'Dim tmp1 As String
'Dim j As Integer
'Dim iStart As Integer
'Dim iEnd As Integer
'Dim LogoLeft As Integer
'Dim LogoTop As Integer
'
'    ar = Split(pIn, ">")
'    Tag = ar(0)
'    Select Case Tag
'    Case "<PRINTINGSETTINGS"
'        ar2 = Split(CStr(ar(1)))
'        j = 0
'        Do While j <= UBound(ar2)
'            If Left(ar2(j), 4) = "CREDI" Then
'                iStart = InStr(1, ar2(j), "{") + 1
'                iEnd = InStr(1, ar2(j), "}")
'                tmp = Mid(ar2(j), iStart, iEnd - iStart)
'                ar3 = Split(tmp, ",")
'                LogoLeft = ar3(0)
'                LogoTop = ar3(1)
'            End If
'            j = j + 1
'        Loop
'    Case "<PHONE"
'        ReplaceMarker "<PHONE>", CStr(ar(1))
'    Case "<STATUS"
'        ReplaceMarker "<STATUS>", CStr(ar(1))
'    Case "<FAX"
'        ReplaceMarker "<FAX>", CStr(ar(1))
'    Case "<ACCNUMBER"
'        ReplaceMarker "<ACCNUMBER>", CStr(ar(1))
'    Case "<DATE"
'        ReplaceMarker "<DATE>", CStr(ar(1))
'    Case "<CONAME"
'        ReplaceMarker "<CONAME>", CStr(ar(1))
'    Case "<OURADDRESS"
'     '   tmp1 = Replace(ar(1), "%", Chr(13))
'        ReplaceMarker "<OURADDRESS>", UnpackText(ar(1))
'    Case "<DOCNUMBER"
'        ReplaceMarker "<DOCNUMBER>", CStr(ar(1))
'    Case "<CUSTOMER"
'     '   tmp1 = Replace(ar(1), "%", Chr(13))
'        ReplaceMarker "<CUSTOMER>", UnpackText(ar(1))
'    Case "<COREGNUMBER"
'        ReplaceMarker "<COREGNUMBER>", CStr(ar(1))
'    Case "<SM"
'        ReplaceMarker "<SM>", CStr(ar(1))
'    Case "<MEMO"
'        ReplaceMarker "<MEMO>", CStr(ar(1))
'    Case "<VATNUMBER"
'        ReplaceMarker "<VATNUMBER>", CStr(ar(1))
'    Case "<HIDEDISCOUNTCOLUMN"
'        bHideDiscountColumn = True
'    Case "<HIDEREFERENCECOLUMN"
'        bHideReferenceColumn = True
'    Case "<HIDEDEPOSITCOLUMN"
'        bHideDepositColumn = True
'    Case "<Line"
'        If iCurRow > 1 Then
'        End If
'        ar1 = Split(ar(1), "|")
'    End Select
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ManageDocument.ParseLine_CN(pIn,pLogoPath)", Array(pIn, pLogoPath)
'End Sub
'Public Sub Builddocument_TF(pFilename As String)
'    On Error GoTo errHandler
'Dim strIP As String
'Dim strTotalValues As String
'Dim i As Integer
'Dim ar() As String
'
'    iFilenum1 = FreeFile
'    Open pFilename For Input As #iFilenum1
'    iCurRow = 1
'    iCurTotalRow = 1
'
''Read first line containing printer information
'    Line Input #iFilenum1, strIP
'    ar = Split(strIP, "|")
'
'    strPrinter = ar(1)
'    lngCopies = ar(4)
'    strStyle = ar(3)
'    strPrintPreview = ar(2)
'
''    If Not CheckPrinterExists(strPrinter, strPrinter) Then
''        WritetoErrors "Printer does not exist (printing Transfers): " & strPrinter & vbCrLf & "Filename = " & pFilename
''        strPrinter = Printer.DeviceName
''    End If
'
''No output device exists, exit
'    If strPrinter = "" Then
'        Close #iFilenum1
'        Exit Sub
'    End If
'
''Set style
'    Select Case strStyle
'    Case "A"
'        Set rptTF = New arAPP_2     ''arTF_1
'    End Select
'    ReDim LineArray(0)
'
''Read each text file line
'    Do Until EOF(iFilenum1)
'        Line Input #iFilenum1, strIP
'        ParseLine_TF strIP
'    Loop
'
''Prepare total set
'    strTotalValues = ""
'    For i = 1 To UBound(LineTotalArray, 1)
'        strTotalValues = strTotalValues & IIf(strTotalValues > "", vbCrLf, "") & LineTotalArray(i)
'    Next
'    ''rptTF.Component LineArray, strTotalValues, strTotalDescription
'    Close #iFilenum1
'
''Open active report
'    If mbPreview Then
'        rptTF.Show vbModal
'    Else
'        rptTF.Printer.DeviceName = IdentifyPrinter(strPrinter)
'        rptTF.Printer.Copies = lngCopies
'        rptTF.PrintReport False
'    End If
'
'    Set rptTF = Nothing
'
'    Exit Sub
'errHandler:
'    ErrPreserve
'    If ErrMustStop Then Debug.Assert False: Resume
'    Set rptTF = Nothing
'    CloseFiles iFilenum1
'    ErrorIn "ManageDocument.Builddocument_TF(pFilename)", Array(pFilename)
'End Sub
'Private Sub ParseLine_TF(pIn As String)
'    On Error GoTo errHandler
'
'Dim ar() As String
'Dim ar1() As String
'Dim Tag As String
'Dim tmp1 As String
'
'    ar = Split(pIn, ">")
'    Tag = ar(0)
'    Select Case Tag
'
'    Case "<LINETOTALCOUNT"
'        lngLineTotalCount = CLng(ar(1))
'        ReDim LineTotalArray(lngLineTotalCount)
'    Case "<DOCUMENTCODE"
'        rptTF.Sections("REPORTHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
'    Case "<STATUS"
'        rptTF.Sections("REPORTHEADER").Controls("STATUS") = CStr(ar(1))
'    Case "<OURADDRESS"
'        rptTF.Sections("REPORTHEADER").Controls("OURADDRESS") = UnpackText(ar(1))
'    Case "<PHONE"
'        rptTF.Sections("REPORTHEADER").Controls("PHONE") = CStr(ar(1))
'    Case "<STATUS"
'    Case "<BILLTO"
'        rptTF.Sections("REPORTHEADER").Controls("BILLTO") = UnpackText(ar(1))
'    Case "<DELTO"
'        rptTF.Sections("REPORTHEADER").Controls("DELTO") = UnpackText(ar(1))
'    Case "<FAX"
'        rptTF.Sections("REPORTHEADER").Controls("FAX") = CStr(ar(1))
'    Case "<MEMO"
'        rptTF.Sections("REPORTHEADER").Controls("fMemo") = UnpackText(CStr(ar(1)))
'    Case "<STORECODE"
'        rptTF.Sections("PAGEFOOTER").Controls("fstorecode") = CStr(ar(1))
'    Case "<SM"
'        rptTF.fSM = CStr(ar(1))
'    Case "<ACCNUMBER"
'    Case "<DATE"
'        rptTF.Sections("REPORTHEADER").Controls("DATE") = CStr(ar(1))
'    Case "<CUSTOMER"
'        rptTF.Sections("REPORTHEADER").Controls("CUSTOMER") = CStr(ar(1))
'    Case "<Line"
'        ReDim Preserve LineArray(UBound(LineArray) + 1)
'        LineArray(iCurRow) = ar(1)
'        iCurRow = iCurRow + 1
'    Case "<TOTAL"
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    Case "<TOTAL2"
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    Case "<TOTAL3"
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    Case "<TOTAL1"
'        LineTotalArray(iCurTotalRow) = ar(1)
'        iCurTotalRow = iCurTotalRow + 1
'    Case "<TOTALDESCRIPTION"
'        strTotalDescription = UnpackText(ar(1))
'       ' strTotalDescription = Replace(strTotalDescription, "|", vbCrLf)
'    End Select
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    CloseFiles iFilenum1
'    ErrorIn "ManageDocument.ParseLine_TF(pIn)", pIn
'End Sub

Public Sub Builddocument_AR(pFilename As String)
    On Error GoTo errHandler
Dim strIP As String
Dim ar() As String

    iFilenum1 = FreeFile
    Open pFilename For Input As #iFilenum1
    iCurRow = 1
    iCurTotalRow = 1
    
    Line Input #iFilenum1, strIP
    ar = Split(strIP, "|")
    
    strPrinter = ar(1)
    lngCopies = ar(4)
    strStyle = ar(3)
    strPrintPreview = ar(2)
    
'    If Not CheckPrinterExists(strPrinter, strPrinter) Then
'        WritetoErrors "Printer does not exist (printing appro return): " & strPrinter & vbCrLf & "Filename = " & pFilename
'        strPrinter = Printer.DeviceName
'    End If
    
    If strPrinter = "" Then
        Close #iFilenum1
        Exit Sub
    End If
    ReDim LineArray(0)
    
    Select Case strStyle
    Case "A"
        Set rptAR = New arAPPR_1
    End Select
    
    Do Until EOF(iFilenum1)
        Line Input #iFilenum1, strIP
        ParseLine_AR strIP
    Loop
    
    rptAR.Component LineArray, LineTotalArray
    Close #iFilenum1
    If mbPreview Then
        rptAR.Show vbModal
    Else
        rptAR.Printer.DeviceName = IdentifyPrinter(strPrinter)
        rptAR.Printer.Copies = lngCopies
        rptAR.PrintReport False
    End If
    Set rptAR = Nothing
    
    Exit Sub
errHandler:
    ErrPreserve
    If ErrMustStop Then Debug.Assert False: Resume
    Set rptAR = Nothing
    CloseFiles iFilenum1
    ErrorIn "ManageDocument.Builddocument_AR(pFilename)", Array(pFilename)
End Sub
Private Sub ParseLine_AR(pIn As String)
    On Error GoTo errHandler

Dim ar() As String
Dim ar1() As String
Dim Tag As String
Dim tmp1 As String

    ar = Split(pIn, ">")
    Tag = ar(0)
    Select Case Tag
'    Case "<LINECOUNT"
'        lngLineCount = CLng(ar(1))
'        ReDim LineArray(lngLineCount)
    Case "<LINETOTALCOUNT"
        lngLineTotalCount = CLng(ar(1))
        ReDim LineTotalArray(lngLineTotalCount)
    Case "<DOCUMENTCODE"
        rptAR.Sections("PAGEHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
    Case "<COMPANYNAME"
        rptAR.Sections("PAGEHEADER").Controls("FFROM") = CStr(ar(1))
    Case "<PHONE"
        rptAR.Sections("PAGEHEADER").Controls("PHONE") = CStr(ar(1))
    Case "<MEMO"
        rptAR.Sections("PAGEHEADER").Controls("fMemo") = UnpackText(CStr(ar(1)))
    Case "<BILLTO"
     '  tmp1 = Replace(ar(1), "%", Chr(13))
        rptAP.Sections("PAGEHEADER").Controls("BILLTO") = UnpackText(ar(1))
    Case "<DELTO"
     '   tmp1 = Replace(ar(1), "%", Chr(13))
        rptAR.Sections("PAGEHEADER").Controls("DELTO") = UnpackText(ar(1))
    Case "<STORECODE"
        rptAR.Sections("PAGEFOOTER").Controls("fstorecode") = CStr(ar(1))
    Case "<FAX"
        rptAR.Sections("PAGEHEADER").Controls("FAX") = CStr(ar(1))
    Case "<SM"
        rptAR.fSM = CStr(ar(1))
    Case "<ACCNUMBER"
      '  rptCO.Fields("ACCNUMBER") = CStr(ar(1))
    Case "<DATE"
        rptAR.Sections("PAGEHEADER").Controls("DATE") = CStr(ar(1))
    Case "<CUSTOMER"
        rptAR.Sections("PAGEHEADER").Controls("CUSTOMER") = "Customer: " & CStr(ar(1))
    Case "<Line"
        ReDim Preserve LineArray(UBound(LineArray) + 1)
        LineArray(iCurRow) = ar(1)
        iCurRow = iCurRow + 1
    Case "<TOTAL"
        LineTotalArray(iCurTotalRow) = ar(1)
        iCurTotalRow = iCurTotalRow + 1
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.ParseLine_AR(pIn)", pIn
End Sub

'Private Function CheckPrinterExists(pDevice As String, pFullDeviceName As String) As Boolean
'Dim pdf As XpdfPrint.XpdfPrint
'Set pdf = New XpdfPrint.XpdfPrint
'Dim nPrinters As Long
'Dim i As Integer
'Dim printerName As String
'
'    On Error GoTo errHandler
'    CheckPrinterExists = False
'    If pDevice > "" Then
'        For Each p In Printers
'            If InStr(1, pDevice, ParseDeviceName(p.DeviceName)) > 0 Then
'                pFullDeviceName = p.DeviceName
'                CheckPrinterExists = True
'                Exit For
'            End If
'        Next
'    End If
'    nPrinters = pdf.getNumPrinters
'    For i = 0 To nPrinters - 1
'        printerName = pdf.getPrinterName(i)
'        MsgBox "printer " & i & ": " & printerName
'    Next i
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ManageDocument.CheckPrinterExists(pDevice)", pDevice
'End Function


Private Sub LoadFormats_PO()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fc, f, fol

    Set fol = fs.GetFolder(strTemplateFolder)
    Set fc = fol.files
    For Each f In fc
        If UCase(f.Name) = "LOGO.BMP" Or UCase(f.Name) = "LOGO.JPG" Then
            rptPO.Sections("REPORTHEADER").Controls("fLOGO").Picture = LoadPicture(f.Path)
            rptPO.Sections("REPORTHEADER").Controls("fLOGO").PictureAlignment = ddPATopRight
        End If
    Next
   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.LoadFormats_PO"
End Sub

'Private Sub LoadFormats_INV()
'    On Error GoTo errHandler
'Dim fs As New FileSystemObject
'Dim fc, f, fol
'
'    Set fol = fs.GetFolder(strTemplateFolder)
'    Set fc = fol.files
'    For Each f In fc
'        Select Case UCase(f.Name)
'        Case "INVOICE.LAY"
'            LoadInvoiceFormat f.Path
'        End Select
'    Next
'    For Each f In fc
'        If UCase(f.Name) = "LOGO.BMP" Or UCase(f.Name) = "LOGO.JPG" Then
'            rptINV.Sections("REPORTHEADER").Controls("fLOGO").Picture = LoadPicture(f.Path)
'        '    rptINVCUSTOM.Sections("REPORTHEADER").Height = rptINVCUSTOM.Sections("REPORTHEADER").Controls("fLOGO").Height + 200
'        End If
'    Next
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ManageDocument.LoadFormats_INV"
'End Sub
'Private Sub LoadFormats_CN()
'    On Error GoTo errHandler
'Dim fs As New FileSystemObject
'Dim fc, f, fol
'
'    Set fol = fs.GetFolder(strTemplateFolder)
'    Set fc = fol.files
'    For Each f In fc
'
'        Select Case UCase(f.Name)
'        Case "CN.LAY"
'            Set rptCN = New arCN
'            LoadCNFormat f.Path
'        End Select
'    Next
'    For Each f In fc
'        If UCase(f.Name) = "LOGO.BMP" Or UCase(f.Name) = "LOGO.JPG" Then
'            rptCN.Sections("REPORTHEADER").Controls("fLOGO").Picture = LoadPicture(f.Path)
'           ' rptCN.Sections("PAGEHEADER").Controls("fLOGO").PictureAlignment = ddPATopRight
'
'        End If
'    Next
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ManageDocument.LoadFormats"
'End Sub
'Private Sub LoadCNFormat(pFilename As String)
'    On Error GoTo errHandler
'Dim oTF As z_TextFile
'Dim ar() As String
'Dim ard() As String
'Dim arprop() As String
'Dim strLine As String
'Dim i As Integer
'
'    Set oTF = New z_TextFile
'
'    oTF.OpenTextFileToRead pFilename
'    Do While Not oTF.IsEOF
'        strLine = oTF.ReadLinefromTextFile
'        If InStr(1, strLine, "|") > 0 Then
'            ar = Split(strLine, "|")
'            Select Case ar(0)
'            Case "ORIENTATION"
'                If ar(1) = "LANDSCAPE" Then
'                    rptCN.Printer.Orientation = ddOLandscape
'                Else
'                    rptCN.Printer.Orientation = ddOPortrait
'                End If
'            Case "REPORTHEADERHEIGHT"
'                rptCN.Sections("REPORTHEADER").Height = ar(1)
'            Case "PAGEHEADERHEIGHT"
'                rptCN.Sections("PAGEHEADER").Height = ar(1)
'            Case "PAGENUMBER"
'                rptCN.Sections("PAGEFOOTER").Controls("fPagenumber").Left = ar(1)
'                rptCN.Sections("PAGEFOOTER").Controls("label2").Left = ar(1) + 800
'                rptCN.Sections("PAGEFOOTER").Controls("fPageMax").Left = ar(1) + 1200
'            Case "LEFTMARGIN"
'
'                rptCN.PageSettings.LeftMargin ar(1)
'            Case "TOPMARGIN"
'                rptCN.PageSettings.TopMargin = ar(1)
'            Case "MEMOPOSITION"
'                Select Case ar(1)
'                Case "TOP"
'                    rptCN.Sections("REPORTFOOTER").Controls("fMEMOFOOTER").Visible = False
'                    rptCN.Sections("REPORTHEADER").Controls("fMEMOHEADER").Visible = True
'                Case "BOTTOM"
'                    rptCN.Sections("REPORTHEADER").Controls("fMEMOHEADER").Visible = False
'                    rptCN.Sections("REPORTFOOTER").Controls("fMEMOFOOTER").Visible = True
'                End Select
'            Case "LINE1Y"
'                ard = Split(ar(1), ",")
'                rptCN.Sections("PAGEHEADER").Controls("Line1").Y1 = ard(0)
'                rptCN.Sections("PAGEHEADER").Controls("Line1").Y2 = ard(0)
'                rptCN.Sections("PAGEHEADER").Controls("Line1").X2 = ard(1)
'            Case "LINE4Y"
'                ard = Split(ar(1), ",")
'                rptCN.Sections("PAGEFOOTER").Controls("Line4").Y1 = ard(0)
'                rptCN.Sections("PAGEFOOTER").Controls("Line4").Y2 = ard(0)
'                rptCN.Sections("PAGEFOOTER").Controls("Line4").X2 = ard(1)
'            Case "LINE3Y"
'                ard = Split(ar(1), ",")
'                rptCN.Sections("PAGEHEADER").Controls("Line3").Y1 = ard(0)
'                rptCN.Sections("PAGEHEADER").Controls("Line3").Y2 = ard(0)
'                rptCN.Sections("PAGEHEADER").Controls("Line3").X2 = ard(1)
'            Case "LINE2Y"
'                ard = Split(ar(1), ",")
'                rptCN.Sections("REPORTFOOTER").Controls("Line2").Y1 = ard(0)
'                rptCN.Sections("REPORTFOOTER").Controls("Line2").Y2 = ard(0)
'                rptCN.Sections("REPORTFOOTER").Controls("Line2").X2 = ard(1)
'            Case "PAGENUMBER"
'                rptCN.Sections("PAGEFOOTER").Controls("fPagenumber").Left = ar(1)
'                rptCN.Sections("PAGEFOOTER").Controls("label2").Left = ar(1) + 800
'                rptCN.Sections("PAGEFOOTER").Controls("fPageMax").Left = ar(1) + 1200
'            Case "LEFTMARGIN"
'                rptCN.PageSettings.LeftMargin = ar(1)
'            Case "TOPMARGIN"
'                rptCN.PageSettings.TopMargin = ar(1)
'            Case Else
'                If InStr(1, ar(2), ",") > 0 Then
'                    ard = Split(ar(2), ",")
'                    For i = 0 To UBound(ard)
'                        If InStr(1, ard(i), ":") > 0 Then
'                            arprop = Split(ard(i), ":")
'                            Select Case arprop(0)
'                            Case "TOP"
'                                rptCN.Sections(ar(0)).Controls(ar(1)).TOP = arprop(1)
'                            Case "LEFT"
'                                 rptCN.Sections(ar(0)).Controls(ar(1)).Left = arprop(1)
'                            Case "WIDTH"
'                                 rptCN.Sections(ar(0)).Controls(ar(1)).Width = arprop(1)
'                            Case "HEIGHT"
'                                 rptCN.Sections(ar(0)).Controls(ar(1)).Height = arprop(1)
'                            Case "ALIGN"
'                                    On Error Resume Next
'                                 rptCN.Sections(ar(0)).Controls(ar(1)).Alignment = IIf(UCase(arprop(1)) = "LEFT", 0, IIf(UCase(arprop(1)) = "RIGHT", 1, 2))
'                                    On Error GoTo errHandler
'                            Case "VISIBLE"
'                                 rptCN.Sections(ar(0)).Controls(ar(1)).Visible = arprop(1)
'                            Case "BOLD"
'                                 rptCN.Sections(ar(0)).Controls(ar(1)).Font.Bold = IIf(arprop(1), 1, 0)
'                            Case "SIZE"
'                                 rptCN.Sections(ar(0)).Controls(ar(1)).Font.Size = arprop(1)
'                            End Select
'                        End If
'                    Next i
'                End If
'            End Select
'        End If
'    Loop
'    oTF.CloseTextFile
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ManageDocument.LoadCNFormat(pFilename)", pFilename
'End Sub

'Private Sub LoadInvoiceFormat(pFilename As String)
'    On Error GoTo errHandler
'Dim oTF As z_TextFile
'Dim ar() As String
'Dim ard() As String
'Dim arprop() As String
'Dim strLine As String
'Dim arINV As arINV_Custom1
'Dim i As Integer
'
'    Set oTF = New z_TextFile
'    oTF.OpenTextFileToRead pFilename
'    Do While Not oTF.IsEOF
'        strLine = oTF.ReadLinefromTextFile
'        If InStr(1, strLine, "|") > 0 Then
'            ar = Split(strLine, "|")
'          '  If ar(1) = "fTOTAL" Then MsgBox "here"
'            Select Case ar(0)
'            Case "ORIENTATION"
'                If ar(1) = "LANDSCAPE" Then
'                    rptINV.Printer.Orientation = ddOLandscape
'                Else
'                    rptINV.Printer.Orientation = ddOPortrait
'                End If
'            Case "PRINTWIDTH"
'                rptINV.PrintWidth = ar(1)
'            Case "REPORTHEADERHEIGHT"
'                rptINV.Sections("REPORTHEADER").Height = ar(1)
'            Case "PAGEHEADERHEIGHT"
'                rptINV.Sections("PAGEHEADER").Height = ar(1)
'            Case "PAGENUMBER"
'                rptINV.Sections("PAGEFOOTER").Controls("fPagenumber").Left = ar(1)
'                rptINV.Sections("PAGEFOOTER").Controls("lblOf").Left = ar(1) + 800
'                rptINV.Sections("PAGEFOOTER").Controls("fPageMax").Left = ar(1) + 1200
'            Case "LEFTMARGIN"
'                rptINV.PageSettings.LeftMargin = ar(1)
'            Case "HIDESSCOLUMN"
'                If ar(1) = "TRUE" Then
'                    rptINV.Sections("DETAIL").Controls("fSS").Visible = False
'                    rptINV.Sections("PAGEHEADER").Controls("lblSS").Visible = False
'                  '  rptINVCUSTOM.Sections("DETAIL").Controls("fFirm").Left = rptINVCUSTOM.Sections("DETAIL").Controls("fFirm").Left + 450
'                  '  rptINVCUSTOM.Sections("PAGEHEADER").Controls("lblFirm").Left = rptINVCUSTOM.Sections("PAGEHEADER").Controls("lblFirm").Left + 450
'                    rptINV.Sections("PAGEHEADER").Controls("lblFirm").Caption = "Qty"
'                End If
'            Case "TOPMARGIN"
'                rptINV.PageSettings.TopMargin = ar(1)
'            Case "MEMOPOSITION"
'                Select Case ar(1)
'                Case "TOP"
'                    rptINV.Sections("REPORTFOOTER").Controls("fMEMOFOOTER").Visible = False
'                    rptINV.Sections("REPORTHEADER").Controls("fMEMOHEADER").Visible = True
'                Case "BOTTOM"
'                    rptINV.Sections("REPORTHEADER").Controls("fMEMOHEADER").Visible = False
'                    rptINV.Sections("REPORTFOOTER").Controls("fMEMOFOOTER").Visible = True
'                End Select
'            Case "LINE1Y"
'                ard = Split(ar(1), ",")
'                rptINV.Sections("PAGEHEADER").Controls("Line1").Y1 = ard(0)
'                rptINV.Sections("PAGEHEADER").Controls("Line1").Y2 = ard(0)
'                rptINV.Sections("PAGEHEADER").Controls("Line1").X2 = ard(1)
'            Case "LINE4Y"
'                ard = Split(ar(1), ",")
'                rptINV.Sections("PAGEFOOTER").Controls("Line4").Y1 = ard(0)
'                rptINV.Sections("PAGEFOOTER").Controls("Line4").Y2 = ard(0)
'                rptINV.Sections("PAGEFOOTER").Controls("Line4").X2 = ard(1)
'            Case "LINE3Y"
'                ard = Split(ar(1), ",")
'                rptINV.Sections("PAGEHEADER").Controls("Line3").Y1 = ard(0)
'                rptINV.Sections("PAGEHEADER").Controls("Line3").Y2 = ard(0)
'                rptINV.Sections("PAGEHEADER").Controls("Line3").X2 = ard(1)
'            Case "LINE2Y"
'                ard = Split(ar(1), ",")
'                rptINV.Sections("REPORTFOOTER").Controls("Line2").Y1 = ard(0)
'                rptINV.Sections("REPORTFOOTER").Controls("Line2").Y2 = ard(0)
'                rptINV.Sections("REPORTFOOTER").Controls("Line2").X2 = ard(1)
'            Case Else
'
'                If InStr(1, ar(2), ",") > 0 Then
'                    ard = Split(ar(2), ",")
'                    For i = 0 To UBound(ard)
'                        If InStr(1, ard(i), ":") > 0 Then
'                            arprop = Split(ard(i), ":")
'                            Select Case UCase(arprop(0))
'                            Case "TOP"
'                                rptINV.Sections(ar(0)).Controls(ar(1)).TOP = arprop(1)
'                            Case "LEFT"
'                                 rptINV.Sections(ar(0)).Controls(ar(1)).Left = arprop(1)
'                            Case "RIGHT"
'                                 rptINV.Sections(ar(0)).Controls(ar(1)).Left = arprop(1) - rptINV.Sections(ar(0)).Controls(ar(1)).Width
'                            Case "WIDTH"
'                                 rptINV.Sections(ar(0)).Controls(ar(1)).Width = arprop(1)
'                            Case "HEIGHT"
'                                 rptINV.Sections(ar(0)).Controls(ar(1)).Height = arprop(1)
'                            Case "ALIGN"
'                                    On Error Resume Next
'                                If ar(1) = "fLOGO" Then
'                                    rptINV.Sections("REPORTHEADER").Controls("fLOGO").PictureAlignment = IIf(UCase(arprop(1)) = "LEFT", ddPATopLeft, IIf(UCase(arprop(1)) = "RIGHT", ddPATopRight, 2))
'                                Else
'                                    rptINV.Sections(ar(0)).Controls(ar(1)).Alignment = IIf(UCase(arprop(1)) = "LEFT", 0, IIf(UCase(arprop(1)) = "RIGHT", 1, 2))
'                                End If
'                                    On Error GoTo errHandler
'                            Case "VISIBLE"
'                                 rptINV.Sections(ar(0)).Controls(ar(1)).Visible = arprop(1)
'                            Case "BOLD"
'                                 rptINV.Sections(ar(0)).Controls(ar(1)).Font.Bold = IIf(arprop(1), 1, 0)
'                            Case "SIZE"
'                                 rptINV.Sections(ar(0)).Controls(ar(1)).Font.Size = arprop(1)
'                            End Select
'                        End If
'                    Next i
'                End If
'            End Select
'        End If
'    Loop
'    oTF.CloseTextFile
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ManageDocument.LoadInvoiceFormat(pFilename)", pFilename
'End Sub

Public Sub InitializeManager(pTemplateFolder)
    
    strTemplateFolder = pTemplateFolder
  '  LoadFormats

End Sub
'Public Sub Builddocument_INVCustom(pFilename As String, pLogoPath As String, pbProforma As Boolean)
'    On Error GoTo errHandler
'Dim strIP As String
'Dim ar() As String
'Dim i As Integer
'
'    iFilenum1 = FreeFile
'    Open pFilename For Input As #iFilenum1
'    iCurRow = 1
'    iCurTotalRow = 1
'
'
'    Line Input #iFilenum1, strIP
'    ar = Split(strIP, "|")
'
'    strPrinter = ar(1)
'    lngCopies = ar(4)
'    strStyle = ar(3)
'    strPrintPreview = ar(2)
'
'
'    bHideDiscountColumn = False
'    bHideReferenceColumn = False
'    bHideDepositColumn = False
'    bHideQtyColumn = False
'
'    For i = 1 To 5
'        Line Input #iFilenum1, strIP
'        ParseLine_INVCUSTOM strIP, pbProforma
'    Next
'
'    ReDim LineArray(0)
'    ReDim LineTotalArray(0)
'
'    Select Case strStyle
'    Case "A"
'        Set rptINV = New arINV_Custom1
'        LoadFormats_INV
'        Do Until EOF(iFilenum1)
'            Line Input #iFilenum1, strIP
'            ParseLine_INVCUSTOM strIP, pbProforma
'        Loop
'        rptINV.Component LineArray, LineTotalArray
'    Case "B"
'       Set rptINV = New arINV_Custom2
'        LoadFormats_INV
'        Do Until EOF(iFilenum1)
'            Line Input #iFilenum1, strIP
'            ParseLine_INVCUSTOM strIP, pbProforma
'        Loop
'        rptINV.Component LineArray, LineTotalArray
'    End Select
'
'    Close #iFilenum1
'
''If Billto and Delto are the same only show one of them
'    If rptINV.Sections("ReportHeader").Controls("fBillto").Text = rptINV.Sections("ReportHeader").Controls("fDelto").Text Then
'        rptINV.Sections("ReportHeader").Controls("fDelto").Visible = False
'        rptINV.Sections("ReportHeader").Controls("lblDelto").Visible = False
'    End If
''Hide appropriate columns
'    If bHideDiscountColumn Then
'        rptINV.Sections("Detail").Controls("fDiscount").Visible = False
'        rptINV.Sections("PAGEHEADER").Controls("lblDiscount").Visible = False
'    End If
'    If bHideQtyColumn Then
'        rptINV.Sections("Detail").Controls("fFirm").Visible = False
'        rptINV.Sections("PAGEHEADER").Controls("lblFirm").Visible = False
'        rptINV.Sections("Detail").Controls("fSS").Visible = False
'        rptINV.Sections("PAGEHEADER").Controls("lblSS").Visible = False
'        rptINV.Sections("Detail").Controls("fDescription").Width = rptINV.Sections("Detail").Controls("fDescription").Width + 1200
'    End If
'    If bHideReferenceColumn Then
'        rptINV.Sections("Detail").Controls("fRef").Visible = False
'        rptINV.Sections("PAGEHEADER").Controls("lblRef").Visible = False
'        rptINV.Sections("Detail").Controls("fDescription").Width = rptINV.Sections("Detail").Controls("fDescription").Width + 1000
'    End If
'
'
'    If mbPreview Then
'        rptINV.Show vbModal
'    Else
''    If CheckPrinterExists(strPrinter, strPrinter) Then
'        'Note: orientation gets reset when device name etc is set so we must remember and reset
'            i = rptINV.Printer.Orientation
'            rptINV.Printer.DeviceName = IdentifyPrinter(strPrinter)
'            rptINV.Printer.Copies = lngCopies
'            rptINV.Printer.Orientation = i
'            rptINV.PrintReport False
''        ElseIf Not Printer Is Nothing Then  'Set default printer
''            i = rptINV.Printer.Orientation
''            rptINV.Printer.DeviceName = IdentifyPrinter(strPrinter)
''            rptINV.Printer.Copies = lngCopies
''            rptINV.Printer.Orientation = i
''            rptINV.PrintReport False
''        End If
'    End If
'    Set rptINV = Nothing
'  '  Set rptINVCUSTOM = Nothing
'    Exit Sub
'errHandler:
'    ErrPreserve
'    If ErrMustStop Then Debug.Assert False: Resume
'    Set rptINV = Nothing
'    CloseFiles iFilenum1
'    ErrorIn "ManageDocument.Builddocument_INVCustom(pFilename)", Array(pFilename)
'End Sub
'Public Sub Builddocument_CNCustom(pFilename As String, pLogoPath As String)
'    On Error GoTo errHandler
'Dim strIP As String
'Dim ar() As String
'Dim i As Integer
'
'    iFilenum1 = FreeFile
'    Open pFilename For Input As #iFilenum1
'    iCurRow = 1
'    iCurTotalRow = 1
'
'    Line Input #iFilenum1, strIP
'    ar = Split(strIP, "|")
'
'    strPrinter = ar(1)
'    lngCopies = ar(4)
'    strStyle = ar(3)
'    strPrintPreview = ar(2)
'
'    bHideDiscountColumn = False
'    bHideReferenceColumn = False
'    bHideDepositColumn = False
'    bHideQtyColumn = False
'
'    For i = 1 To 5
'        Line Input #iFilenum1, strIP
'        ParseLine_CNCUSTOM strIP
'    Next
'    ReDim LineArray(0)
'    ReDim LineTotalArray(0)
'
'    Select Case strStyle
'    Case "A", "B"
'        LoadFormats_CN
'        Do Until EOF(iFilenum1)
'            Line Input #iFilenum1, strIP
'            ParseLine_CNCUSTOM strIP
'        Loop
'        rptCN.Component LineArray, LineTotalArray
'    End Select
'
'    Close #iFilenum1
'
''Hide appropriate columns
''If Billto and Delto are the same only show one of them
'    If rptCN.Sections("ReportHeader").Controls("fBillto").Text = rptCN.Sections("ReportHeader").Controls("fDelto").Text Then
'        rptCN.Sections("ReportHeader").Controls("fDelto").Visible = False
'        rptCN.Sections("ReportHeader").Controls("lblDelto").Visible = False
'    End If
''Hide appropriate columns
'    If bHideDiscountColumn Then
'        rptCN.Sections("Detail").Controls("fDiscount").Visible = False
'        rptCN.Sections("PAGEHEADER").Controls("lblDiscount").Visible = False
'    End If
'    If bHideQtyColumn Then
'        rptCN.Sections("Detail").Controls("fFirm").Visible = False
'        rptCN.Sections("PAGEHEADER").Controls("lblFirm").Visible = False
'        rptCN.Sections("Detail").Controls("fSS").Visible = False
'        rptCN.Sections("PAGEHEADER").Controls("lblSS").Visible = False
'        rptCN.Sections("Detail").Controls("fDescription").Width = rptCN.Sections("Detail").Controls("fDescription").Width + 1200
'    End If
'    If bHideReferenceColumn Then
'        rptCN.Sections("Detail").Controls("fRef").Visible = False
'        rptCN.Sections("PAGEHEADER").Controls("lblRef").Visible = False
'        rptCN.Sections("Detail").Controls("fDescription").Width = rptCN.Sections("Detail").Controls("fDescription").Width + 1000
'    End If
'
'    If mbPreview Then
'        rptCN.Show vbModal
'    Else
'            i = rptCN.Printer.Orientation
'            rptCN.Printer.DeviceName = IdentifyPrinter(strPrinter)
'            rptCN.Printer.Copies = lngCopies
'            rptCN.Printer.Orientation = i
'            rptCN.PrintReport False
'    End If
'    Set rptCN = Nothing
'    Exit Sub
'errHandler:
'    ErrPreserve
'    If ErrMustStop Then Debug.Assert False: Resume
'    Set rptCN = Nothing
'    CloseFiles iFilenum1
'    ErrorIn "ManageDocument.Builddocument_CNCustom(pFilename)", Array(pFilename)
'End Sub
'
Private Sub ClearControls()
    With rptINV
    .Sections("REPORTHEADER").Controls("fBILLTO").Text = ""
    .Sections("REPORTHEADER").Controls("fDELTO") = ""
    .Sections("REPORTHEADER").Controls("fDATE") = ""
    .Sections("REPORTHEADER").Controls("fMEMOHEADER") = ""
    .Sections("REPORTFOOTER").Controls("fMEMOFOOTER") = ""
    .Sections("REPORTHEADER").Controls("fSTORE") = ""
    .Sections("DETAIL").Controls("fCODE") = ""
    .Sections("DETAIL").Controls("fDESCRIPTION") = ""
    .Sections("DETAIL").Controls("fREF") = ""
    .Sections("DETAIL").Controls("fFIRM") = ""
    .Sections("DETAIL").Controls("fSS") = ""
    .Sections("DETAIL").Controls("fPRICE") = ""
    .Sections("DETAIL").Controls("fDISCOUNT") = ""
    .Sections("DETAIL").Controls("fExtension") = ""
    .Sections("PAGEFOOTER").Controls("fBank") = ""
    .Sections("PAGEFOOTER").Controls("fVAT") = ""
    End With
End Sub




