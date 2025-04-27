VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ManageDocument"
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
Dim rptTF As arTF_1
Dim rptINV As Object
Dim rptINVCUSTOM As arINV_2
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

Public Sub Builddocument_INV(pFilename As String, pTemplate As String, pstrLogoPath As String)
    On Error GoTo errHandler
Dim strIP As String
Dim i As Integer
Dim ar() As String
Dim strPos As String
Dim strFullDeviceName As String

    bHideDiscountColumn = False
    bHideReferenceColumn = False
    bHideDepositColumn = False
    bHideQtyColumn = False
    Set oDoc = wm.StartNewDocument(pTemplate)
    strPos = "Position 1"
    iFilenum1 = FreeFile
    
'Read first set of lines fist to set up columns etc.
    Open pFilename For Input As #iFilenum1
    For i = 1 To 5
        Line Input #iFilenum1, strIP
        ParseLine_INV strIP, pstrLogoPath
    Next
    
    Line Input #iFilenum1, strIP
    ar = Split(strIP, "|")
    strFullDeviceName = ar(1)
    If Not CheckPrinterExists(strFullDeviceName) Then
        WritetoErrors "Printer does not exist (printing invoice): " & strFullDeviceName & ". Using activeprinter: " & wm.WORDApplication.ActivePrinter & vbCrLf & "Filename = " & pFilename
        strFullDeviceName = wm.WORDApplication.ActivePrinter
    End If
    
    strPos = "Position 2"
    If strFullDeviceName = "" Then
        oDoc.Close 0
        Close #iFilenum1
        Exit Sub
    End If
    strPos = "Position 3"
    
    FindColumnPositionInTable iISBN, iTitle, iRef, iQty, iDisc, iDep, iPrice, iExt, iEXVAT
    If bHideDepositColumn Then
        oDoc.Tables(1).Columns(iTitle).Width = oDoc.Tables(1).Columns(iTitle).Width + oDoc.Tables(1).Columns(iDep).Width
        oDoc.Tables(1).Columns(iDep).Delete '(oDoc.Tables(1).Columns(iDisc + 1).Cells)
    End If
    strPos = "Position 4"
    FindColumnPositionInTable iISBN, iTitle, iRef, iQty, iDisc, iDep, iPrice, iExt, iEXVAT
    If bHideDiscountColumn Then
        oDoc.Tables(1).Columns(iTitle).Width = oDoc.Tables(1).Columns(iTitle).Width + oDoc.Tables(1).Columns(iDisc).Width
        oDoc.Tables(1).Columns(iDisc).Delete '(oDoc.Tables(1).Columns(iDisc + 1).Cells)
    End If
    strPos = "Position 5"
    FindColumnPositionInTable iISBN, iTitle, iRef, iQty, iDisc, iDep, iPrice, iExt, iEXVAT
    If bHideQtyColumn Then
        oDoc.Tables(1).Columns(iTitle).Width = oDoc.Tables(1).Columns(iTitle).Width + oDoc.Tables(1).Columns(iQty).Width
        oDoc.Tables(1).Columns(iQty).Delete '(oDoc.Tables(1).Columns(iDisc + 1).Cells)
        FindColumnPositionInTable iISBN, iTitle, iRef, iQty, iDisc, iDep, iPrice, iExt, iEXVAT
        oDoc.Tables(1).Columns(iTitle).Width = oDoc.Tables(1).Columns(iTitle).Width + oDoc.Tables(1).Columns(iPrice).Width
        oDoc.Tables(1).Columns(iPrice).Delete '(oDoc.Tables(1).Columns(iDisc + 1).Cells)
    End If
    strPos = "Position 6"
    FindColumnPositionInTable iISBN, iTitle, iRef, iQty, iDisc, iDep, iPrice, iExt, iEXVAT
    If bHideReferenceColumn Then
        oDoc.Tables(1).Columns(iTitle).Width = oDoc.Tables(1).Columns(iTitle).Width + oDoc.Tables(1).Columns(iRef).Width
        oDoc.Tables(1).Columns(iRef).Delete '(oDoc.Tables(1).Columns(iDisc + 1).Cells)
    End If
    FindColumnPositionInTable iISBN, iTitle, iRef, iQty, iDisc, iDep, iPrice, iExt, iEXVAT
    
'Read next lines for content
    strPos = "Position 7"
    iCurRow = 2
    iCurTotalRow = 1
    Do Until EOF(iFilenum1)
        Line Input #iFilenum1, strIP
        If strIP > "" Then ParseLine_INV strIP, pstrLogoPath
    Loop
    strPos = "Position 8"
    oDoc.Tables(1).Rows(oDoc.Tables(1).Rows.Count).Delete
    oDoc.Tables(2).Rows(oDoc.Tables(2).Rows.Count).Delete
    'format tables
    strPos = "Position 9"
    oDoc.Tables(1).Rows(1).Borders(3).LineStyle = 7 'wdLineStyleDouble   'wdBorderBottom
    oDoc.Tables(1).Rows(1).Borders(3).LineWidth = 2 'wdLineWidth025pt
    oDoc.Tables(1).Rows(oDoc.Tables(1).Rows.Count).Borders(3).LineStyle = 13 'wdLineStyleThickThinMedGap
    oDoc.Tables(1).Rows(oDoc.Tables(1).Rows.Count).Borders(3).LineWidth = 2 'wdLineWidth025pt
    ReplaceMarker "<STATUS>", ""
    If wm.WordIsVisible Then
        oDoc.PrintPreview
        Forms(0).objT.Enabled = False
        MsgBox "View document"
        Forms(0).objT.Enabled = True
    Else
        wm.WORDApplication.WordBasic.FILEPRINTSETUP Printer:=strFullDeviceName, DoNotSetAsSysDefault:=1
        oDoc.PrintOut Copies:=ar(4), Background:=False
    End If
    oDoc.Close 0
    Close #iFilenum1
    
    Exit Sub
errHandler:
    ErrPreserve
    If Not oDoc Is Nothing Then
        On Error Resume Next
        oDoc.Close 0
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrCloseFiles iFilenum1
    ErrorIn "ManageDocument.Builddocument_INV(pFilename,pTemplate,pstrLogoPath)", _
         Array(pFilename, pTemplate, pstrLogoPath), , , "Marker", strPos
End Sub

Private Sub ParseLine_INV(pIn As String, pLogoPath As String)
    On Error GoTo errHandler
Dim tmp As String
Dim ar() As String
Dim ar1() As String
Dim Tag As String
Dim tmp1 As String
Dim LogoLeft As Integer
Dim LogoTop As Integer
Dim ar2() As String
Dim ar3() As String
Dim j As Integer
Dim iStart As Integer
Dim iEnd As Integer

    ar = Split(pIn, ">")
    Tag = ar(0)
    Select Case Tag
    Case "<PRINTINGSETTINGS"
        ar2 = Split(CStr(ar(1)))
        j = 0
        Do While j <= UBound(ar2)
            If Left(ar2(j), 4) = "INVO" Then
                iStart = InStr(1, ar2(j), "{") + 1
                iEnd = InStr(1, ar2(j), "}")
                tmp = Mid(ar2(j), iStart, iEnd - iStart)
                ar3 = Split(tmp, ",")
                LogoLeft = ar3(0)
                LogoTop = ar3(1)
            End If
            j = j + 1
        Loop
    Case "<PHONE"
        ReplaceMarker "<PHONE>", CStr(ar(1))
    Case "<BILLTO"
        ReplaceMarker "<BILLTO>", UnpackText(CStr(ar(1)))
    Case "<DELTO"
        ReplaceMarker "<DELTO>", UnpackText(CStr(ar(1)))
    Case "<STATUS"
        ReplaceMarker "<STATUS>", CStr(ar(1))
    Case "<FAX"
        ReplaceMarker "<FAX>", CStr(ar(1))
    Case "<ACCNUMBER"
        ReplaceMarker "<ACCNUMBER>", CStr(ar(1))
    Case "<DATE"
        ReplaceMarker "<DATE>", CStr(ar(1))
    Case "<MEMO"
        ReplaceMarker "<MEMO>", IIf(Trim(ar(1)) > "", "Note: ", "") & UnpackText(CStr(ar(1)))
    Case "<OURADDRESS"
   '     tmp1 = Replace(ar(1), "%", Chr(13))
        ReplaceMarker "<OURADDRESS>", UnpackText(ar(1))
    Case "<DOCUMENTCODE"
        ReplaceMarker "<INVOICENUMBER>", CStr(ar(1))
    Case "<SM"
        ReplaceMarker "<SM>", CStr(ar(1))
    Case "<CONAME"
        ReplaceMarker "<CONAME>", CStr(ar(1))
    Case "<CUSTOMER"
      '  tmp1 = Replace(ar(1), "%", Chr(13))
        ReplaceMarker "<CUSTOMER>", UnpackText(ar(1))
    Case "<COREGNUMBER"
        ReplaceMarker "<COREGNUMBER>", CStr(ar(1))
    Case "<VATNUMBER"
        ReplaceMarker "<VATNUMBER>", "V.A.T. No. " & CStr(ar(1))
    Case "<HIDEDISCOUNTCOLUMN"
        bHideDiscountColumn = True
    Case "<HIDEREFERENCECOLUMN"
        bHideReferenceColumn = True
    Case "<HIDEDEPOSITCOLUMN"
        bHideDepositColumn = True
    Case "<HIDEQTYCOLUMN"
        bHideQtyColumn = True
    Case "<Line"
        If iCurRow > 1 Then
            oDoc.Tables(1).Rows(iCurRow).Select
            wm.WORDApplication.Selection.InsertRows 1
   '         Selection.Font.Bold = False
        End If
        ar1 = Split(ar(1), "|")
        If iISBN > 0 And UBound(ar1) > -1 Then
            oDoc.Tables(1).Cell(iCurRow, iISBN).range = CStr(ar1(0))
        End If
        If iRef > 0 And UBound(ar1) > 0 Then
            oDoc.Tables(1).Cell(iCurRow, iRef).range = CStr(ar1(1))
        End If
        If iQty > 0 And UBound(ar1) > 1 Then
            oDoc.Tables(1).Cell(iCurRow, iQty).range = CStr(ar1(2))
        End If
        If iTitle > 0 And UBound(ar1) > 3 Then
            oDoc.Tables(1).Cell(iCurRow, iTitle).range = Left(CStr(ar1(3)), iTitleWidth) & IIf(Len(ar1(3)) > iTitleWidth, ".../", "  ") & Left(CStr(ar1(4)), iTitleWidth) '& IIf(Len(ar1(4)) > iAuthorWidth, ".../", "  ") & Left(CStr(ar1(5)), iTitleWidth)
        End If
        If iPrice > 0 And UBound(ar1) > 5 Then
            oDoc.Tables(1).Cell(iCurRow, iPrice).range = CStr(ar1(6))
        End If
        If iExt > 0 And UBound(ar1) > 6 Then
            oDoc.Tables(1).Cell(iCurRow, iExt).range = CStr(ar1(7))
        End If
        If iDep > 0 And UBound(ar1) > 7 Then
            oDoc.Tables(1).Cell(iCurRow, iDep).range = CStr(ar1(8))
        End If
        If iDisc > 0 And UBound(ar1) > 8 Then
            oDoc.Tables(1).Cell(iCurRow, iDisc).range = CStr(ar1(9))
        End If
        If iEXVAT > 0 And UBound(ar1) > 9 Then
            oDoc.Tables(1).Cell(iCurRow, iEXVAT).range = CStr(ar1(10))
        End If
        If UBound(ar1) > 10 Then
            If ar1(11) > " " Then
                wm.WORDApplication.Selection.ParagraphFormat.SpaceAfter = 0
                iCurRow = iCurRow + 1
                oDoc.Tables(1).Rows(iCurRow).Select
                wm.WORDApplication.Selection.InsertRows 1
                wm.WORDApplication.Selection.ParagraphFormat.SpaceBefore = 0
                oDoc.Tables(1).Cell(iCurRow, iTitle).range = CStr(ar1(11))
                iCurRow = iCurRow + 1
                oDoc.Tables(1).Rows(3).Select
            Else
                iCurRow = iCurRow + 1
            End If
        Else
                iCurRow = iCurRow + 1
        End If
   Case "<TOTAL"
        If iCurTotalRow > 0 Then
            oDoc.Tables(2).Rows(iCurTotalRow).Select
            wm.WORDApplication.Selection.InsertRows 1
        End If
        ar1 = Split(ar(1), "|")
        oDoc.Tables(2).Cell(iCurTotalRow, 1).range = CStr(ar1(0))
        oDoc.Tables(2).Cell(iCurTotalRow, 2).range = CStr(ar1(1))
        iCurTotalRow = iCurTotalRow + 1
   Case "<TOTAL2"
        If iCurTotalRow > 0 Then
            oDoc.Tables(2).Rows(iCurTotalRow).Select
            wm.WORDApplication.Selection.InsertRows 1
        End If
        ar1 = Split(ar(1), "|")
        oDoc.Tables(2).Cell(iCurTotalRow, 1).range = CStr(ar1(0))
        oDoc.Tables(2).Cell(iCurTotalRow, 2).range = CStr(ar1(1))
        oDoc.Tables(2).Rows(iCurTotalRow).range.Font.Bold = False
        oDoc.Tables(2).Rows(iCurTotalRow).range.Font.Italic = True
        oDoc.Tables(2).Cell(iCurTotalRow, 2).Borders(-1).LineStyle = 7 'wdLineStyleSingle 1 'wdBorderTop -1
       ' oDoc.Tables(2).Cell(iCurTotalRow, 2).Borders(-3).LineWidth = 4 'wdLineWidth050pt
        oDoc.Tables(2).Rows(iCurTotalRow).Borders(-3).LineStyle = 0
        iCurTotalRow = iCurTotalRow + 1
   Case "<TOTAL3"
        If iCurTotalRow > 0 Then
            oDoc.Tables(2).Rows(iCurTotalRow).Select
            wm.WORDApplication.Selection.InsertRows 1
        End If
        ar1 = Split(ar(1), "|")
        oDoc.Tables(2).Cell(iCurTotalRow, 1).range = CStr(ar1(0))
        oDoc.Tables(2).Cell(iCurTotalRow, 2).Borders(-3).LineStyle = 1 'wdLineStyleDouble 7 (does not work)  'wdBorderBottom -3
        oDoc.Tables(2).Cell(iCurTotalRow, 2).Borders(-3).LineWidth = 4 'wdLineWidth050pt
        oDoc.Tables(2).Cell(iCurTotalRow, 2).range = CStr(ar1(1))
        iCurTotalRow = iCurTotalRow + 1
   Case "<TOTAL4"
        If iCurTotalRow > 0 Then
            oDoc.Tables(2).Rows(iCurTotalRow).Select
            wm.WORDApplication.Selection.InsertRows 1
        End If
        ar1 = Split(ar(1), "|")
        oDoc.Tables(2).Rows(iCurTotalRow).range.Font.Bold = False
        oDoc.Tables(2).Cell(iCurTotalRow, 1).range = CStr(ar1(0))
        oDoc.Tables(2).Cell(iCurTotalRow, 2).range = CStr(ar1(1))
        iCurTotalRow = iCurTotalRow + 1
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.ParseLine_INV(pIn,pLogoPath)", Array(pIn, pLogoPath)
End Sub
Private Sub ParseLine_wordversion_9(pIn As String)
    On Error GoTo errHandler
Dim ar() As String
Dim ar1() As String
Dim Tag As String
Dim tmp1 As String

    ar = Split(pIn, ">")
    Tag = ar(0)
    Select Case Tag
    Case "<LOGO"
        oDoc.Frames(1).Select
        oDoc.InlineShapes.AddPicture CStr(ar(1)), True, True
    Case "<PRINTINGSETTINGS"
    Case "<PHONE"
        ReplaceMarker "<PHONE>", CStr(ar(1))
    Case "<FAX"
        ReplaceMarker "<FAX>", CStr(ar(1))
    Case "<ACCNUMBER"
        ReplaceMarker "<ACCNUMBER>", CStr(ar(1))
    Case "<DATE"
        ReplaceMarker "<DATE>", CStr(ar(1))
    Case "<OURADDRESS"
   '     tmp1 = Replace(ar(1), "%", Chr(13))
        ReplaceMarker "<OURADDRESS>", UnpackText(ar(1))
    Case "<INVOICENUMBER"
        ReplaceMarker "<INVOICENUMBER>", CStr(ar(1))
    Case "<CUSTOMER"
    '    tmp1 = Replace(ar(1), "%", Chr(13))
        ReplaceMarker "<CUSTOMER>", UnpackText(ar(1))
    Case "<COREGNUMBER"
        ReplaceMarker "<COREGNUMBER>", CStr(ar(1))
    Case "<VATNUMBER"
        ReplaceMarker "<VATNUMBER>", CStr(ar(1))
    Case "<HIDEDISCOUNTCOLUMN"
        bHideDiscountColumn = True
    Case "<HIDEDEPOSITCOLUMN"
        bHideDepositColumn = True
    Case "<Line"
        If iCurRow > 1 Then
            oDoc.Tables(1).Rows(iCurRow).Select
    '        Selection.InsertRowsBelow 1
        End If
        iCurRow = iCurRow + 1
        ar1 = Split(ar(1), "|")
        If iISBN > 0 Then
            oDoc.Tables(1).Cell(iCurRow, iISBN).range = CStr(ar1(0))
        End If
        If iRef > 0 Then
            oDoc.Tables(1).Cell(iCurRow, iRef).range = CStr(ar1(1))
        End If
        If iQty > 0 Then
            oDoc.Tables(1).Cell(iCurRow, iQty).range = CStr(ar1(2))
        End If
        If iTitle > 0 Then
            oDoc.Tables(1).Cell(iCurRow, iTitle).range = Left(CStr(ar1(3)), iTitleWidth) & ".../" & Left(CStr(ar1(4)), iTitleWidth) & ".../" & Left(CStr(ar1(5)), iTitleWidth)
        End If
        If iPrice > 0 Then
            oDoc.Tables(1).Cell(iCurRow, iPrice).range = CStr(ar1(6))
        End If
        If iExt > 0 Then
            oDoc.Tables(1).Cell(iCurRow, iExt).range = CStr(ar1(7))
        End If
        If iDep > 0 Then
            oDoc.Tables(1).Cell(iCurRow, iDep).range = CStr(ar1(8))
        End If
        If iDisc > 0 Then
            oDoc.Tables(1).Cell(iCurRow, iDisc).range = CStr(ar1(9))
        End If
    Case "<TOTAL"
        If iCurTotalRow > 0 Then
            oDoc.Tables(2).Rows(iCurTotalRow).Select
   '         Selection.InsertRowsBelow 1
        End If
        iCurTotalRow = iCurTotalRow + 1
        ar1 = Split(ar(1), "|")
        oDoc.Tables(2).Cell(iCurTotalRow, 1).range = CStr(ar1(0))
        oDoc.Tables(2).Cell(iCurTotalRow, 2).range = CStr(ar1(1))
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.ParseLine_wordversion_9(pIn)", pIn
End Sub

Private Sub FillatBookmark(pBM As String, pVal As String)
    On Error GoTo errHandler
    wm.FillBM pBM, pVal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.FillatBookmark(pBM,pVal)", Array(pBM, pVal)
End Sub
Private Sub ReplaceMarker(pMarker As String, pVal As String)
    On Error GoTo errHandler
    wm.ReplaceMarker pMarker, pVal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.ReplaceMarker(pMarker,pVal)", Array(pMarker, pVal)
End Sub
Private Sub FindColumnPositionInTable(iISBN As Integer, iTitle As Integer, iRef As Integer, iQty As Integer, iDisc As Integer, iDep As Integer, iPrice As Integer, iExt As Integer, iEXVAT As Integer)
    On Error GoTo errHandler
Dim iCol As Integer
Dim ar() As String
Dim tmp As String
Dim iEnd As Integer
Dim iBegin As Integer

    iISBN = 0
    iTitle = 0
    iRef = 0
    iQty = 0
    iDisc = 0
    iDep = 0
    iPrice = 0
    iExt = 0
    iEXVAT = 0
    
    For iCol = 1 To oDoc.Tables(1).Columns.Count
        With oDoc
            Set range = .Tables(1).Cell(2, iCol).range
            If InStr(1, range.Text, "<ISBN>") Then
                iISBN = iCol
            ElseIf InStr(1, range.Text, "<TITLE>") Then
                iTitle = iCol
                iBegin = InStr(1, range.Text, "{")
                iEnd = InStr(1, range.Text, "}")
                iTitleWidth = 200
                iAuthorWidth = 200
                iPublisherWidth = 200

                If iBegin > 0 Then
                    tmp = Mid(range.Text, iBegin + 1, iEnd - iBegin - 1)
                    ar = Split(tmp, ",")
                    If UBound(ar) > 0 Then
                        iTitleWidth = CInt(ar(0))
                        iAuthorWidth = CInt(ar(1))
                        iPublisherWidth = CInt(ar(2))
                    End If
                End If
            ElseIf InStr(1, range.Text, "<REF>") Then
                iRef = iCol
            ElseIf InStr(1, range.Text, "<QTY>") Then
                iQty = iCol
            ElseIf InStr(1, range.Text, "<PRICE>") Then
                iPrice = iCol
            ElseIf InStr(1, range.Text, "<DEP>") Then
                iDep = iCol
            ElseIf InStr(1, range.Text, "<DISC>") Then
                iDisc = iCol
            ElseIf InStr(1, range.Text, "<EXT>") Then
                iExt = iCol
            ElseIf InStr(1, range.Text, "<EVAT>") Then
                iEXVAT = iCol
            End If
        End With
    Next iCol
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.FindColumnPositionInTable(iISBN,iTitle,iRef,iQty,iDisc,iDep,iPrice,iExt," & _
        "iEXVAT)", Array(iISBN, iTitle, iRef, iQty, iDisc, iDep, iPrice, iExt, iEXVAT)
End Sub

Public Sub Builddocument_DEL(pFilename As String)
    On Error GoTo errHandler
Dim strIP As String
Dim strStyle As String
Dim ar() As String
Dim strFullDeviceName As String
    
    
    iFilenum1 = FreeFile
    Open pFilename For Input As #iFilenum1
    iCurRow = 1
    iCurTotalRow = 1
    
    Line Input #iFilenum1, strIP
    ar = Split(strIP, "|")
   
    strFullDeviceName = ar(1)
    If Not CheckPrinterExists(strFullDeviceName) Then
        strFullDeviceName = Printer.DeviceName
    End If
    
    If strFullDeviceName = "" Then
        Close #iFilenum1
        Exit Sub
    End If
    
    
    strStyle = ar(3)
    
    Select Case strStyle
    Case "A"
        Set rptDEL = New arDEL_1
    End Select
    
    Do Until EOF(iFilenum1)
        Line Input #iFilenum1, strIP
        ParseLine_DEL strIP
    Loop
    rptDEL.Component LineArray, LineTotalArray
    Close #iFilenum1
    If mbPreview Then
        rptDEL.Show vbModal
    Else
        rptDEL.Printer.DeviceName = strFullDeviceName  'ar(1) ') 'ar(1) GetFullDeviceNameFromName(
        rptDEL.Printer.Copies = ar(4)
        rptDEL.PrintReport False
    End If
    Set rptDEL = Nothing
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrCloseFiles iFilenum1
    ErrorIn "ManageDocument.Builddocument_DEL(pFilename)", Array(pFilename)
End Sub

Public Sub Builddocument_APS(pFilename As String)
    On Error GoTo errHandler
Dim strIP As String
Dim strStyle As String
Dim ar() As String
Dim strFullDeviceName As String

    iFilenum1 = FreeFile
    Open pFilename For Input As #iFilenum1
    iCurRow = 1
    iCurTotalRow = 1
    
    Line Input #iFilenum1, strIP
    ar = Split(strIP, "|")
    
    strFullDeviceName = ar(1)
    If Not CheckPrinterExists(strFullDeviceName) Then
        WritetoErrors "Printer does not exist (printing Appro slips): " & strFullDeviceName & ". Using activeprinter: " & wm.WORDApplication.ActivePrinter & vbCrLf & "Filename = " & pFilename
        strFullDeviceName = Printer.DeviceName
    End If
    
    If strFullDeviceName = "" Then
        Close #iFilenum1
        Exit Sub
    End If
    
    strStyle = ar(3)
    
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
        rptAPS.Printer.DeviceName = strFullDeviceName    'ar(1) ') 'ar(1) GetFullDeviceNameFromName(
        rptAPS.Printer.Copies = ar(4)
        rptAPS.PrintReport False
    End If
    Set rptAPS = Nothing
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrCloseFiles iFilenum1
    ErrorIn "ManageDocument.Builddocument_APS(pFilename)", Array(pFilename)
End Sub

Private Function GetFullDeviceNameFromName(pIn As String) As String
Dim i As Integer
Dim r As Long
Dim Buffer As String
Dim ar1() As String
    ' Get the list of available printers from WIN.INI
    Buffer = Space(8192)
    r = GetProfileString("PrinterPorts", vbNullString, "", _
       Buffer, Len(Buffer))

    ' Display the list of printer in the ListBox List1
    ParseList ar1, Buffer

End Function
Private Sub ParseList(ar() As String, ByVal Buffer As String)
    Dim i As Integer
    Dim s As String

    Do
        i = InStr(Buffer, Chr(0))
        If i > 0 Then
            s = Left(Buffer, i - 1)
            If Len(Trim(s)) Then
                ReDim ar(1, UBound(ar, 0) + 1)
                ar(UBound(ar)) = s
            End If
            Buffer = Mid(Buffer, i + 1)
        Else
            If Len(Trim(Buffer)) Then
                ReDim ar(1, UBound(ar, 1) + 1)
                ar(UBound(ar)) = s
            End If
            Buffer = ""
        End If
    Loop While i > 0
End Sub

Private Sub ParseLine_DEL(pIn As String)
    On Error GoTo errHandler

Dim ar() As String
Dim ar1() As String
Dim Tag As String
Dim tmp1 As String

    ar = Split(pIn, ">")
    Tag = ar(0)
    Select Case Tag
    Case "<LINECOUNT"
        lngLineCount = CLng(ar(1))
        ReDim LineArray(lngLineCount)
    Case "<LINETOTALCOUNT"
        lngLineTotalCount = CLng(ar(1))
        ReDim LineTotalArray(lngLineTotalCount)
    Case "<DOCUMENTCODE"
        rptDEL.Sections("PAGEHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
    Case "<PHONE"
        rptDEL.Sections("PAGEHEADER").Controls("PHONE") = CStr(ar(1))
    Case "<SINVREF"
        rptDEL.fSINVREF = "Your Ref. " & CStr(ar(1))
    Case "<SINVDATE"
        rptDEL.fSINVDATE = "Dated: " & CStr(ar(1))
    Case "<FAX"
        rptDEL.Sections("PAGEHEADER").Controls("FAX") = CStr(ar(1))
    Case "<SM"
        rptDEL.fSM = CStr(ar(1))
    Case "<DATE"
        rptDEL.Sections("PAGEHEADER").Controls("DATE") = CStr(ar(1))
    Case "<SUPPLIER"
        rptDEL.Sections("PAGEHEADER").Controls("SUPPLIER") = CStr(ar(1))
    Case "<MEMO"
        rptDEL.Sections("PAGEHEADER").Controls("fMEMO") = CStr(ar(1))
    Case "<Line"
        LineArray(iCurRow) = ar(1)
        iCurRow = iCurRow + 1
    Case "<TOTAL"
        LineTotalArray(iCurTotalRow) = ar(1)
        iCurTotalRow = iCurTotalRow + 1
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.ParseLine_DEL(pIn)", pIn
End Sub

Public Sub Builddocument_PO(pFilename As String, Optional pLogoFilename As String)
    On Error GoTo errHandler
Dim strIP As String
Dim ar() As String
Dim strStyle As String
Dim strFullDeviceName As String

    
    iFilenum1 = FreeFile
    Open pFilename For Input As #iFilenum1
    iCurRow = 1
    iCurTotalRow = 1
    
    Line Input #iFilenum1, strIP
    ar = Split(strIP, "|")
    
    strFullDeviceName = ar(1)
    If Not CheckPrinterExists(strFullDeviceName) Then
        WritetoErrors "Printer does not exist (printing Purchase order): " & strFullDeviceName & ". Using activeprinter: " & wm.WORDApplication.ActivePrinter & vbCrLf & "Filename = " & pFilename
        strFullDeviceName = Printer.DeviceName
    End If
    
    If strFullDeviceName = "" Then
        Close #iFilenum1
        Exit Sub
    End If
    
    
    strStyle = ar(3)
    
    Select Case strStyle
    Case "A"
        Set rptPO = New arPO_1
        Do Until EOF(iFilenum1)
            Line Input #iFilenum1, strIP
            ParseLine_PO_A strIP
        Loop
    Case "B"
        Set rptPO = New arPO_2
        Do Until EOF(iFilenum1)
            Line Input #iFilenum1, strIP
            ParseLine_PO_B strIP
        Loop
    End Select
    
    rptPO.Component LineArray, LineTotalArray, pLogoFilename
  '  rptPO..Orientation = ddOLandscape
    Close #iFilenum1
    If mbPreview Then
        rptPO.Printer.Orientation = ddOLandscape
        rptPO.Show vbModal
    Else
        rptPO.Printer.DeviceName = strFullDeviceName       'p.DeviceName  ' & " on " & p.Port
        rptPO.Printer.Orientation = ddOLandscape
        rptPO.Printer.PaperSize = 9
        rptPO.PageLeftMargin = 800
        rptPO.PageRightMargin = 0
        rptPO.Printer.Copies = CInt(ar(4))
        rptPO.PrintReport False
    End If
    Set rptPO = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrCloseFiles iFilenum1
    ErrorIn "ManageDocument.Builddocument_PO(pFilename)", Array(pFilename)
End Sub
Public Sub Builddocument_R(pFilename As String, Optional pLogoFilename As String)
    On Error GoTo errHandler
Dim strIP As String
Dim ar() As String
Dim strStyle As String
Dim strFullDeviceName As String

    
    iFilenum1 = FreeFile
    Open pFilename For Input As #iFilenum1
    iCurRow = 1
    iCurTotalRow = 1
    
    Line Input #iFilenum1, strIP
    ar = Split(strIP, "|")
    
    strFullDeviceName = ar(1)
    If Not CheckPrinterExists(strFullDeviceName) Then
        WritetoErrors "Printer does not exist (printing Return): " & strFullDeviceName & ". Using activeprinter: " & wm.WORDApplication.ActivePrinter & vbCrLf & "Filename = " & pFilename
        strFullDeviceName = Printer.DeviceName
    End If
    
    If strFullDeviceName = "" Then
        Close #iFilenum1
        Exit Sub
    End If
    
    
    
    strStyle = ar(3)
    
    Select Case strStyle
    Case "A"
        Set rptR = New arR_2
    Case "B"
        Set rptR = New arR_1
    End Select
    
    Do Until EOF(iFilenum1)
        Line Input #iFilenum1, strIP
        ParseLine_R strIP
    Loop
    rptR.Component LineArray, LineTotalArray, pLogoFilename
    Close #iFilenum1
    If mbPreview Then
        rptR.Show vbModal
    Else
        rptR.Printer.DeviceName = strFullDeviceName      'ar(1) ') 'ar(1) GetFullDeviceNameFromName(
        rptR.Printer.PaperSize = 9
        rptR.PageLeftMargin = 800
        rptR.Printer.Copies = ar(4)
        rptR.PrintReport False
    End If
    Set rptR = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrCloseFiles iFilenum1
    ErrorIn "ManageDocument.Builddocument_R(pFilename)", Array(pFilename)
End Sub
Private Sub ParseLine_PO_A(pIn As String)
    On Error GoTo errHandler

Dim ar() As String
Dim ar1() As String
Dim ar2() As String
Dim Tag As String
Dim tmp1 As String

    ar = Split(pIn, ">")
    Tag = ar(0)
    Select Case Tag
    Case "<LINECOUNT"
        lngLineCount = CLng(ar(1))
        ReDim LineArray(lngLineCount)
    Case "<LINETOTALCOUNT"
        lngLineTotalCount = CLng(ar(1))
        ReDim LineTotalArray(lngLineTotalCount)
    Case "<DOCUMENTCODE"
        rptPO.Sections("REPORTHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
        On Error Resume Next
        rptPO.Sections("REPORTHEADER").Controls("fDOCUMENTCODEFOOT") = CStr(ar(1))
            On Error GoTo errHandler
    Case "<SENDER"
        rptPO.Sections("REPORTHEADER").Controls("fSENDER") = CStr(ar(1))
    Case "<OURADDRESS"
        rptPO.Sections("REPORTHEADER").Controls("fOURADDRESS") = UnpackText(ar(1))
    Case "<PHONE"
        rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") = rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") & "     " & CStr(ar(1))
    Case "<STATUS"
        rptPO.Sections("REPORTHEADER").Controls("fStatus") = CStr(ar(1))
    Case "<BILLTO"
        rptPO.Sections("REPORTHEADER").Controls("BILLTO") = UnpackText(ar(1))
    Case "<DELTO"
      '  tmp1 = Replace(ar(1), "%", Chr(13))
        rptPO.Sections("REPORTHEADER").Controls("DELTO") = UnpackText(ar(1))
    Case "<FAX"
        rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") = rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") & "     " & CStr(ar(1))
    Case "<ACCNUMBER"
        rptPO.Fields("ACCNUMBER") = CStr(ar(1))
    Case "<SM"
        rptPO.Sections("REPORTHEADER").Controls("fSM") = CStr(ar(1))
    Case "<DATE"
        rptPO.Sections("REPORTHEADER").Controls("DATE") = CStr(ar(1))
    Case "<FOOT"
        rptPO.Sections("PAGEFOOTER").Controls("fFOOT") = CStr(ar(1))
    Case "<SUPPLIER"
        rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") = CStr(ar(1))
    Case "<SUPPLIERADDRESS"
    'This is only on PO layout "B"
        On Error Resume Next
        rptPO.Sections("REPORTHEADER").Controls("fSUPPLIERAddress") = UnpackText(CStr(ar(1)))
        Err.Clear
        On Error GoTo errHandler

    Case "<MAINMSG"
      '  tmp1 = Replace(CStr(ar(1)), "%", Chr(13))
        rptPO.Sections("REPORTFOOTER").Controls("fORDERTEXT") = UnpackText(ar(1)) 'CStr(ar(1))
    Case "<Line"
        LineArray(iCurRow) = ar(1)
        iCurRow = iCurRow + 1
    Case "<TOTAL"
        LineTotalArray(iCurTotalRow) = UnpackText(ar(1))    'ar(1)
        iCurTotalRow = iCurTotalRow + 1
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.ParseLine_PO_A(pIn)", pIn
End Sub

Private Sub ParseLine_PO_B(pIn As String)
    On Error GoTo errHandler

Dim ar() As String
Dim ar1() As String
Dim ar2() As String
Dim Tag As String
Dim tmp1 As String

    ar = Split(pIn, ">")
    Tag = ar(0)
    Select Case Tag
    Case "<LINECOUNT"
        lngLineCount = CLng(ar(1))
        ReDim LineArray(lngLineCount)
    Case "<LINETOTALCOUNT"
        lngLineTotalCount = CLng(ar(1))
        ReDim LineTotalArray(lngLineTotalCount)
    Case "<DOCUMENTCODE"
        rptPO.Sections("REPORTHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
        rptPO.Sections("PAGEFOOTER").Controls("fDOCUMENTCODEFOOT") = CStr(ar(1))
    Case "<SENDER"
        rptPO.Sections("REPORTHEADER").Controls("fSENDER") = CStr(ar(1))
    Case "<PHONE"
        rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") = rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") & "     " & CStr(ar(1))
    Case "<STATUS"
        rptPO.Sections("REPORTHEADER").Controls("fStatus") = CStr(ar(1))
    Case "<OURADDRESS"
        rptPO.Sections("REPORTHEADER").Controls("fOURADDRESS") = UnpackText(ar(1))
    Case "<BILLTO"
        rptPO.Sections("REPORTHEADER").Controls("BILLTO") = UnpackText(ar(1))
    Case "<DELTO"
      '  tmp1 = Replace(ar(1), "%", Chr(13))
        rptPO.Sections("REPORTHEADER").Controls("DELTO") = UnpackText(ar(1))
    Case "<FAX"
        rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") = rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") & "     " & CStr(ar(1))
    Case "<ACCNUMBER"
        rptPO.Fields("ACCNUMBER") = CStr(ar(1))
    Case "<SM"
        rptPO.Sections("REPORTHEADER").Controls("fSM") = CStr(ar(1))
    Case "<DATE"
        rptPO.Sections("REPORTHEADER").Controls("DATE") = CStr(ar(1))
    Case "<FOOT"
        rptPO.Sections("PAGEFOOTER").Controls("fFOOT") = CStr(ar(1))
    Case "<SUPPLIER"
        rptPO.Sections("REPORTHEADER").Controls("SUPPLIER") = CStr(ar(1))
    Case "<SUPPLIERADDRESS"
    'This is only on PO layout "B"
        On Error Resume Next
        rptPO.Sections("REPORTHEADER").Controls("fSUPPLIERAddress") = UnpackText(CStr(ar(1)))
        Err.Clear
        On Error GoTo errHandler

    Case "<MAINMSG"
      '  tmp1 = Replace(CStr(ar(1)), "%", Chr(13))
        rptPO.Sections("REPORTFOOTER").Controls("fORDERTEXT") = UnpackText(ar(1)) 'CStr(ar(1))
    Case "<Line"
        LineArray(iCurRow) = ar(1)
        iCurRow = iCurRow + 1
    Case "<TOTAL"
        LineTotalArray(iCurTotalRow) = UnpackText(ar(1))    'ar(1)
        iCurTotalRow = iCurTotalRow + 1
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.ParseLine_PO_B(pIn)", pIn
End Sub
Private Sub ParseLine_R(pIn As String)
    On Error GoTo errHandler

Dim ar() As String
Dim ar1() As String
Dim ar2() As String
Dim Tag As String
Dim tmp1 As String

    ar = Split(pIn, ">")
    Tag = ar(0)
    Select Case Tag
    Case "<LINECOUNT"
        lngLineCount = CLng(ar(1))
        ReDim LineArray(lngLineCount)
    Case "<LINETOTALCOUNT"
        lngLineTotalCount = CLng(ar(1))
        ReDim LineTotalArray(lngLineTotalCount)
    Case "<DOCUMENTCODE"
        rptR.Sections("REPORTHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
        rptR.Sections("PAGEFOOTER").Controls("FDOCNO") = CStr(ar(1))
    Case "<PHONE"
        rptR.Sections("REPORTHEADER").Controls("SUPPLIER") = rptR.Sections("REPORTHEADER").Controls("SUPPLIER") & "     " & CStr(ar(1))
    Case "<STATUS"
        rptR.Sections("REPORTHEADER").Controls("fStatus") = CStr(ar(1))
    Case "<BILLTO"
     '   tmp1 = Replace(ar(1), "%", Chr(13))
        rptR.Sections("REPORTHEADER").Controls("BILLTO") = UnpackText(ar(1))
    Case "<DELTO"
'        tmp1 = Replace(ar(1), "%", Chr(13))
'        rptR.Sections("PAGEHEADER").Controls("DELTO") = tmp1
    Case "<SENDER"
        On Error Resume Next
        rptR.Sections("REPORTHEADER").Controls("fSENDER") = CStr(ar(1))
        On Error GoTo errHandler
    Case "<OURADDRESS"
        rptR.Sections("REPORTHEADER").Controls("fOURADDRESS") = UnpackText(ar(1))
    Case "<FAX"
        rptR.Sections("REPORTHEADER").Controls("SUPPLIER") = rptR.Sections("REPORTHEADER").Controls("SUPPLIER") & "     " & CStr(ar(1))
    Case "<SM"
        rptR.Sections("REPORTHEADER").Controls("fSM") = CStr(ar(1))
    Case "<ACCNUMBER"
        rptR.Fields("ACCNUMBER") = CStr(ar(1))
    Case "<DATE"
        rptR.Sections("REPORTHEADER").Controls("DATE") = CStr(ar(1))
    Case "<SUPPLIER"
        rptR.Sections("REPORTHEADER").Controls("SUPPLIER") = CStr(ar(1))
'    Case "<MAINMSG"
'        tmp1 = Replace(CStr(ar(1)), "%", Chr(13))
'        rptR.Sections("REPORTFOOTER").Controls("fORDERTEXT") = tmp1 'CStr(ar(1))
    Case "<Line"
        LineArray(iCurRow) = ar(1)
        iCurRow = iCurRow + 1
    Case "<TOTAL"
        LineTotalArray(iCurTotalRow) = UnpackText(ar(1))    'ar(1)
        iCurTotalRow = iCurTotalRow + 1
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.ParseLine_R(pIn)", pIn
End Sub

Public Sub Builddocument_INV2(pFilename As String, pLogoPath As String)
    On Error GoTo errHandler
Dim strIP As String
Dim ar() As String
Dim strStyle As String

    iFilenum1 = FreeFile
    Open pFilename For Input As #iFilenum1
    iCurRow = 1
    iCurTotalRow = 1
    
    Line Input #iFilenum1, strIP
    ar = Split(strIP, "|")
    strStyle = ar(3)
    
    Select Case strStyle
    Case "A"
        Set rptINV = New arINV_1
    Case "B"
        Set rptINV = rptINVCUSTOM
    End Select
    
    Do Until EOF(iFilenum1)
        Line Input #iFilenum1, strIP
        ParseLine_INV2 strIP
    Loop
    rptINV.Component LineArray, LineTotalArray
    Close #iFilenum1
    If mbPreview Then
        rptINV.Show vbModal
    Else
        If CheckPrinterExists(ar(1)) Then
            rptINV.Printer.DeviceName = ar(1) ') 'ar(1) GetFullDeviceNameFromName(
            rptINV.Printer.Copies = ar(4)
            rptINV.PrintReport False
        ElseIf Not Printer Is Nothing Then  'Set default printer
            rptINV.Printer.DeviceName = Printer.DeviceName
            rptINV.Printer.Copies = ar(4)
            rptINV.PrintReport False
        End If
    End If
    Set rptINV = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrCloseFiles iFilenum1
    ErrorIn "ManageDocument.Builddocument_INV2(pFilename)", Array(pFilename)
End Sub
Private Sub ParseLine_INV2(pIn As String)
    On Error GoTo errHandler

Dim ar() As String
Dim ar1() As String
Dim Tag As String
Dim tmp1 As String
On Error Resume Next
    ar = Split(pIn, ">")
    Tag = ar(0)
    Select Case Tag
    Case "<LINECOUNT"
        lngLineCount = CLng(ar(1))
        ReDim LineArray(lngLineCount)
    Case "<LINETOTALCOUNT"
        lngLineTotalCount = CLng(ar(1))
        ReDim LineTotalArray(lngLineTotalCount)
'    Case "<DOCUMENTCODE"
'        rptINV.Sections("PAGEHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
    Case "<PHONE"
        rptINV.Sections("PAGEHEADER").Controls("PHONE") = CStr(ar(1))
    Case "<OURADDRESS"
        rptINV.Sections("PAGEHEADER").Controls("fStore") = UnpackText(ar(1))
    Case "<STATUS"
      '  rptPO.Fields("STATUS") = CStr(ar(1))
    Case "<DOCUMENTCODE"
        rptINV.Sections("PAGEHEADER").Controls("lblDocument").Caption = "TAX INVOICE No. " & CStr(ar(1))
    Case "<BILLTO"
      '  tmp1 = Replace(ar(1), "%", Chr(13))
        rptINV.Sections("PAGEHEADER").Controls("BILLTO") = UnpackText(ar(1))
    Case "<DELTO"
     '  tmp1 = Replace(ar(1), "%", Chr(13))
        rptINV.Sections("PAGEHEADER").Controls("DELTO") = UnpackText(ar(1))
    Case "<FAX"
        rptINV.Sections("PAGEHEADER").Controls("FAX") = CStr(ar(1))
    Case "<SM"
        rptINV.Sections("PageFOOTER").Controls("lblSM") = CStr(ar(1))
'    Case "<ACCNUMBER"
'        rptINV.Sections("PAGEHEADER").Controls("CUSTOMER") = "To: " & CStr(ar(1))
    Case "<DATE"
        rptINV.Sections("PAGEHEADER").Controls("DATE") = CStr(ar(1))
    Case "<CUSTOMER"
        rptINV.Sections("PAGEHEADER").Controls("CUSTOMER") = "To: " & CStr(ar(1))
    Case "<Line"
        LineArray(iCurRow) = ar(1)
        iCurRow = iCurRow + 1
    Case "<TOTAL"
        LineTotalArray(iCurTotalRow) = ar(1)
        iCurTotalRow = iCurTotalRow + 1
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.ParseLine_INV2(pIn)", pIn
End Sub

'''''''''''
Public Sub Builddocument_CO(pFilename As String)
    On Error GoTo errHandler
Dim strIP As String
Dim strStyle As String
Dim ar() As String
Dim strFullDeviceName As String
    
    iFilenum1 = FreeFile
    Open pFilename For Input As #iFilenum1
    iCurRow = 1
    iCurTotalRow = 1

    Line Input #iFilenum1, strIP
    ar = Split(strIP, "|")
    
    strFullDeviceName = ar(1)
    If Not CheckPrinterExists(strFullDeviceName) Then
        WritetoErrors "Printer does not exist (printing order): " & strFullDeviceName & ". Using activeprinter: " & wm.WORDApplication.ActivePrinter & vbCrLf & "Filename = " & pFilename
        strFullDeviceName = Printer.DeviceName
    End If
    
    If strFullDeviceName = "" Then
        Close #iFilenum1
        Exit Sub
    End If


    strStyle = ar(3)
    
    Select Case strStyle
    Case "A"
        Set rptCO = New arCO_1
    End Select
    
    
    Do Until EOF(iFilenum1)
        Line Input #iFilenum1, strIP
        ParseLine_CO strIP
    Loop
  '  ReplaceMarker "<STATUS>", ""
    rptCO.Component LineArray, LineTotalArray
    Close #iFilenum1
    If mbPreview Then
        rptCO.Show vbModal
    Else
        rptCO.Printer.DeviceName = strFullDeviceName     'ar(1) ') 'ar(1) GetFullDeviceNameFromName(
        rptCO.Printer.Copies = ar(4)
        rptCO.PrintReport False
    End If
    Set rptCO = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrCloseFiles iFilenum1
    ErrorIn "ManageDocument.Builddocument_CO(pFilename)", Array(pFilename)
End Sub
Private Sub ParseLine_CO(pIn As String)
    On Error GoTo errHandler

Dim ar() As String
Dim ar1() As String
Dim Tag As String
Dim tmp1 As String

    ar = Split(pIn, ">")
    Tag = ar(0)
    Select Case Tag
    Case "<LINECOUNT"
        lngLineCount = CLng(ar(1))
        ReDim LineArray(lngLineCount)
    Case "<LINETOTALCOUNT"
        lngLineTotalCount = CLng(ar(1))
        ReDim LineTotalArray(lngLineTotalCount)
    Case "<DOCUMENTCODE"
        rptCO.Sections("PAGEHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
    Case "<PHONE"
        rptCO.Sections("PAGEHEADER").Controls("PHONE") = CStr(ar(1))
    Case "<STATUS"
      '  rptPO.Fields("STATUS") = CStr(ar(1))
    Case "<BILLTO"
   '     tmp1 = Replace(ar(1), "%", Chr(13))
        rptCO.Sections("PAGEHEADER").Controls("BILLTO") = UnpackText(ar(1))
    Case "<DELTO"
      '  tmp1 = Replace(ar(1), "%", Chr(13))
        rptCO.Sections("PAGEHEADER").Controls("DELTO") = UnpackText(ar(1))
    Case "<FAX"
        rptCO.Sections("PAGEHEADER").Controls("FAX") = CStr(ar(1))
    Case "<ACCNUMBER"
      '  rptCO.Fields("ACCNUMBER") = CStr(ar(1))
    Case "<SM"
        rptCO.fSM = CStr(ar(1))
    Case "<MEMO"
        rptCO.fMemo = "Note: " & UnpackText(CStr(ar(1)))
    Case "<DATE"
        rptCO.Sections("PAGEHEADER").Controls("DATE") = CStr(ar(1))
    Case "<CUSTOMER"
        rptCO.Sections("PAGEHEADER").Controls("CUSTOMER") = CStr(ar(1))
    Case "<Line"
        LineArray(iCurRow) = ar(1)
        iCurRow = iCurRow + 1
    Case "<TOTAL"
        LineTotalArray(iCurTotalRow) = ar(1)
        iCurTotalRow = iCurTotalRow + 1
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.ParseLine_CO(pIn)", pIn
End Sub

Public Sub Builddocument_AP(pFilename As String, Optional pLogoFilename As String)
    On Error GoTo errHandler
Dim strIP As String
Dim ar() As String
Dim strStyle As String
Dim strFullDeviceName As String

    iFilenum1 = FreeFile
    Open pFilename For Input As #iFilenum1
    iCurRow = 1
    iCurTotalRow = 1
    
    Line Input #iFilenum1, strIP
    ar = Split(strIP, "|")
    
    strFullDeviceName = ar(1)
    If Not CheckPrinterExists(strFullDeviceName) Then
        WritetoErrors "Printer does not exist (printing Appro): " & strFullDeviceName & ". Using activeprinter: " & wm.WORDApplication.ActivePrinter & vbCrLf & "Filename = " & pFilename
        strFullDeviceName = Printer.DeviceName
    End If
    
    If strFullDeviceName = "" Then
        Close #iFilenum1
        Exit Sub
    End If
    
    strStyle = ar(3)
    
    Select Case strStyle
    Case "A"
        Set rptAP = New arAPP_1
    Case "B"
        Set rptAP = New arAPP_2
    End Select
    
    Do Until EOF(iFilenum1)
        Line Input #iFilenum1, strIP
        ParseLine_AP strIP
    Loop
  '  ReplaceMarker "<STATUS>", ""
    rptAP.Component LineArray, LineTotalArray, pLogoFilename
    Close #iFilenum1
    If mbPreview Then
        rptAP.Show vbModal
    Else
            rptAP.Printer.DeviceName = strFullDeviceName   'ar(1) ') 'ar(1) GetFullDeviceNameFromName(
            rptAP.Printer.Orientation = ddOLandscape
            rptAP.Printer.PaperSize = 9
            rptAP.PageLeftMargin = 800
            rptAP.Printer.Copies = ar(4)
            rptAP.PrintReport False
    End If
    Set rptAP = Nothing
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrCloseFiles iFilenum1
    ErrorIn "ManageDocument.Builddocument_AP(pFilename)", Array(pFilename)
End Sub
Private Sub ParseLine_AP(pIn As String)
    On Error GoTo errHandler

Dim ar() As String
Dim ar1() As String
Dim Tag As String
Dim tmp1 As String

    ar = Split(pIn, ">")
    Tag = ar(0)
    Select Case Tag
    Case "<LINECOUNT"
        lngLineCount = CLng(ar(1))
        ReDim LineArray(lngLineCount)
    Case "<LINETOTALCOUNT"
        lngLineTotalCount = CLng(ar(1))
        ReDim LineTotalArray(lngLineTotalCount)
    Case "<DOCUMENTCODE"
        rptAP.Sections("REPORTHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
        rptAP.Sections("PAGEFOOTER").Controls("fDOCNO") = CStr(ar(1))
    Case "<COMPANYNAME"
        rptAP.Sections("REPORTHEADER").Controls("FFROM") = UnpackText(CStr(ar(1)))
    Case "<PHONE"
        rptAP.Sections("REPORTHEADER").Controls("PHONE") = CStr(ar(1))
    Case "<SENDER"
        On Error Resume Next
        rptAP.Sections("REPORTHEADER").Controls("fSENDER") = CStr(ar(1))
        On Error GoTo errHandler
    Case "<OURADDRESS"
        rptAP.Sections("REPORTHEADER").Controls("fOURADDRESS") = UnpackText(ar(1))
    Case "<STATUS"
      '  rptPO.Fields("STATUS") = CStr(ar(1))
    Case "<MEMO"
        rptAP.Sections("REPORTHEADER").Controls("fMEMO") = CStr(ar(1))
    Case "<BILLTO"
     '  tmp1 = Replace(ar(1), "%", Chr(13))
        rptAP.Sections("REPORTHEADER").Controls("BILLTO") = UnpackText(ar(1))
    Case "<DELTO"
     '   tmp1 = Replace(ar(1), "%", Chr(13))
        rptAP.Sections("REPORTHEADER").Controls("DELTO") = UnpackText(ar(1))
    Case "<FAX"
        rptAP.Sections("REPORTHEADER").Controls("FAX") = CStr(ar(1))
    Case "<SM"
        rptAP.Sections("REPORTHEADER").Controls("fSM") = CStr(ar(1))
    Case "<ACCNUMBER"
      '  rptCO.Fields("ACCNUMBER") = CStr(ar(1))
    Case "<DATE"
        rptAP.Sections("REPORTHEADER").Controls("DATE") = CStr(ar(1))
    Case "<CUSTOMER"
        rptAP.Sections("REPORTHEADER").Controls("CUSTOMER") = CStr(ar(1))
    Case "<Line"
        LineArray(iCurRow) = ar(1)
        iCurRow = iCurRow + 1
    Case "<TOTAL"
        LineTotalArray(iCurTotalRow) = ar(1)
        iCurTotalRow = iCurTotalRow + 1
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.ParseLine_AP(pIn)", pIn
End Sub

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
        lngLineCount = CLng(ar(1))
        ReDim LineArray(lngLineCount)
    Case "<Line"
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



Public Sub Builddocument_CN(pFilename As String, pTemplate As String, pstrLogoPath As String)
    On Error GoTo errHandler
Dim strIP As String
Dim i As Integer
Dim ar() As String
Dim strPos As String
Dim strFullDeviceName As String

'    bHideDiscountColumn = False
    Set oDoc = wm.StartNewDocument(pTemplate)
    iFilenum1 = FreeFile
    Open pFilename For Input As #iFilenum1
    Line Input #iFilenum1, strIP
    ar = Split(strIP, "|")
    strFullDeviceName = ar(1)
    
    If Not CheckPrinterExists(strFullDeviceName) Then
        WritetoErrors "Printer does not exist (printing Credit note): " & strFullDeviceName & ". Using activeprinter: " & wm.WORDApplication.ActivePrinter & vbCrLf & "Filename = " & pFilename
        strFullDeviceName = wm.WORDApplication.ActivePrinter
    End If
    
    If strFullDeviceName = "" Then
        oDoc.Close 0
        Close #iFilenum1
        Exit Sub
    End If
    
    strPos = "Position 1"
    
' '   FindColumnPositionInTable iISBN, iTitle, iRef, iQty, iDisc, iDep, iPrice, iExt, iEXVAT
'    If bHideDiscountColumn Then
'        oDoc.Tables(1).Columns(iTitle).Width = oDoc.Tables(1).Columns(iTitle).Width + oDoc.Tables(1).Columns(iDisc).Width
'        oDoc.Tables(1).Columns(iDisc).Delete
'    End If
    FindColumnPositionInTable iISBN, iTitle, iRef, iQty, iDisc, iDep, iPrice, iExt, iEXVAT
    
    strPos = "Position 2"
    
    iCurRow = 2
    iCurTotalRow = 1
    Do Until EOF(iFilenum1)
        Line Input #iFilenum1, strIP
        ParseLine_CN strIP, pstrLogoPath
    Loop
    strPos = "Position 2.B"
    oDoc.Tables(1).Rows(oDoc.Tables(1).Rows.Count).Delete
    oDoc.Tables(2).Rows(oDoc.Tables(2).Rows.Count).Delete
    'format tables
    oDoc.Tables(1).Rows(1).Borders(3).LineStyle = 7 'wdLineStyleDouble   'wdBorderBottom
    oDoc.Tables(1).Rows(1).Borders(3).LineWidth = 2 'wdLineWidth025pt
    oDoc.Tables(1).Rows(oDoc.Tables(1).Rows.Count).Borders(3).LineStyle = 13 'wdLineStyleThickThinMedGap
    oDoc.Tables(1).Rows(oDoc.Tables(1).Rows.Count).Borders(3).LineWidth = 2 'wdLineWidth025pt
    strPos = "Position 3"
'    FindColumnPositionInTable iISBN, iTitle, iRef, iQty, iDisc, iDep, iPrice, iExt, iEXVAT
'    If bHideReferenceColumn Then
'        oDoc.Tables(1).Columns(iRef - 1).Width = oDoc.Tables(1).Columns(iRef - 1).Width + oDoc.Tables(1).Columns(iRef).Width
'        oDoc.Tables(1).Columns(iRef).Delete '(oDoc.Tables(1).Columns(iDisc + 1).Cells)
'    End If
    strPos = "Position 5"
    ReplaceMarker "<STATUS>", ""
    If wm.WordIsVisible Then
        oDoc.PrintPreview
        Forms(0).objT.Enabled = False
        MsgBox "View document"
        Forms(0).objT.Enabled = True
    Else
        strPos = "Position 6"
        wm.WORDApplication.WordBasic.FILEPRINTSETUP Printer:=strFullDeviceName, DoNotSetAsSysDefault:=1
        strPos = "Position 7"
        oDoc.PrintOut Copies:=ar(4), Background:=False
        strPos = "Position 9"
    End If
    strPos = "Position 10"
    oDoc.Close 0
    strPos = "Position 11"
    Close #iFilenum1
    Exit Sub
errHandler:
    ErrPreserve
    If Not oDoc Is Nothing Then
        On Error Resume Next
        oDoc.Close 0
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrCloseFiles iFilenum1
    ErrorIn "ManageDocument.Builddocument_CN(pFilename,pTemplate,pstrLogoPath)", _
         Array(pFilename, pTemplate, pstrLogoPath), , , "Marker", strPos
End Sub

Private Sub ParseLine_CN(pIn As String, pLogoPath As String)
    On Error GoTo errHandler
Dim tmp As String
Dim ar() As String
Dim ar1() As String
Dim ar2() As String
Dim ar3() As String
Dim Tag As String
Dim tmp1 As String
Dim j As Integer
Dim iStart As Integer
Dim iEnd As Integer
Dim LogoLeft As Integer
Dim LogoTop As Integer

    ar = Split(pIn, ">")
    Tag = ar(0)
    Select Case Tag
    Case "<PRINTINGSETTINGS"
        ar2 = Split(CStr(ar(1)))
        j = 0
        Do While j <= UBound(ar2)
            If Left(ar2(j), 4) = "CREDI" Then
                iStart = InStr(1, ar2(j), "{") + 1
                iEnd = InStr(1, ar2(j), "}")
                tmp = Mid(ar2(j), iStart, iEnd - iStart)
                ar3 = Split(tmp, ",")
                LogoLeft = ar3(0)
                LogoTop = ar3(1)
            End If
            j = j + 1
        Loop
    Case "<PHONE"
        ReplaceMarker "<PHONE>", CStr(ar(1))
    Case "<STATUS"
        ReplaceMarker "<STATUS>", CStr(ar(1))
    Case "<FAX"
        ReplaceMarker "<FAX>", CStr(ar(1))
    Case "<ACCNUMBER"
        ReplaceMarker "<ACCNUMBER>", CStr(ar(1))
    Case "<DATE"
        ReplaceMarker "<DATE>", CStr(ar(1))
    Case "<CONAME"
        ReplaceMarker "<CONAME>", CStr(ar(1))
    Case "<OURADDRESS"
     '   tmp1 = Replace(ar(1), "%", Chr(13))
        ReplaceMarker "<OURADDRESS>", UnpackText(ar(1))
    Case "<DOCNUMBER"
        ReplaceMarker "<DOCNUMBER>", CStr(ar(1))
    Case "<CUSTOMER"
     '   tmp1 = Replace(ar(1), "%", Chr(13))
        ReplaceMarker "<CUSTOMER>", UnpackText(ar(1))
    Case "<COREGNUMBER"
        ReplaceMarker "<COREGNUMBER>", CStr(ar(1))
    Case "<SM"
        ReplaceMarker "<SM>", CStr(ar(1))
    Case "<MEMO"
        ReplaceMarker "<MEMO>", CStr(ar(1))
    Case "<VATNUMBER"
        ReplaceMarker "<VATNUMBER>", CStr(ar(1))
    Case "<HIDEDISCOUNTCOLUMN"
        bHideDiscountColumn = True
    Case "<HIDEREFERENCECOLUMN"
        bHideReferenceColumn = True
    Case "<HIDEDEPOSITCOLUMN"
        bHideDepositColumn = True
    Case "<Line"
        If iCurRow > 1 Then
            oDoc.Tables(1).Rows(iCurRow).Select
            wm.WORDApplication.Selection.InsertRows 1
   '         Selection.Font.Bold = False
        End If
        ar1 = Split(ar(1), "|")
        If iISBN > 0 And UBound(ar1) > -1 Then
            oDoc.Tables(1).Cell(iCurRow, iISBN).range = CStr(ar1(0))
        End If
        If iRef > 0 And UBound(ar1) > 7 Then
            oDoc.Tables(1).Cell(iCurRow, iRef).range = CStr(ar1(8))
        End If
        If iQty > 0 And UBound(ar1) > 1 Then
            oDoc.Tables(1).Cell(iCurRow, iQty).range = CStr(ar1(2))
        End If
        If iTitle > 0 And UBound(ar1) > 2 Then
            oDoc.Tables(1).Cell(iCurRow, iTitle).range = CStr(ar1(3))
        End If
        If iPrice > 0 And UBound(ar1) > 3 Then
            oDoc.Tables(1).Cell(iCurRow, iPrice).range = CStr(ar1(4))
        End If
        If iExt > 0 And UBound(ar1) > 4 Then
            oDoc.Tables(1).Cell(iCurRow, iExt).range = CStr(ar1(5))
        End If
'        If iDep > 0 Then
'            oDoc.Tables(1).Cell(iCurRow, iDep).range = CStr(ar1(6))
'        End If
        If iDisc > 0 And UBound(ar1) > 5 Then
            oDoc.Tables(1).Cell(iCurRow, iDisc).range = CStr(ar1(6))
        End If
        If iEXVAT > 0 And UBound(ar1) > 6 Then
            oDoc.Tables(1).Cell(iCurRow, iEXVAT).range = CStr(ar1(7))
        End If
        If UBound(ar1) > 8 Then
            If ar1(9) > " " Then
                Selection.ParagraphFormat.SpaceAfter = 0
                iCurRow = iCurRow + 1
                oDoc.Tables(1).Rows(iCurRow).Select
                wm.WORDApplication.Selection.InsertRows 1
                Selection.ParagraphFormat.SpaceBefore = 0
                oDoc.Tables(1).Cell(iCurRow, iTitle).range = CStr(ar1(9))
                iCurRow = iCurRow + 1
                oDoc.Tables(1).Rows(iCurRow).Select
            Else
                iCurRow = iCurRow + 1
            End If
        Else
            iCurRow = iCurRow + 1
        End If
   Case "<TOTAL"
        If iCurTotalRow > 0 Then
            oDoc.Tables(2).Rows(iCurTotalRow).Select
            wm.WORDApplication.Selection.InsertRows 1
        End If
        ar1 = Split(ar(1), "|")
        oDoc.Tables(2).Cell(iCurTotalRow, 1).range = CStr(ar1(0))
        oDoc.Tables(2).Cell(iCurTotalRow, 2).range = CStr(ar1(1))
        iCurTotalRow = iCurTotalRow + 1
   Case "<TOTAL2"
        If iCurTotalRow > 0 Then
            oDoc.Tables(2).Rows(iCurTotalRow).Select
            wm.WORDApplication.Selection.InsertRows 1
        End If
        ar1 = Split(ar(1), "|")
        oDoc.Tables(2).Cell(iCurTotalRow, 1).range = CStr(ar1(0))
        oDoc.Tables(2).Rows(iCurTotalRow).range.Font.Bold = False
        oDoc.Tables(2).Cell(iCurTotalRow, 2).Borders(-1).LineStyle = 1 'wdLineStyleSingle 1 'wdBorderTop -1
        oDoc.Tables(2).Cell(iCurTotalRow, 2).Borders(-1).LineWidth = 4 'wdLineWidth050pt
        oDoc.Tables(2).Cell(iCurTotalRow, 2).range = CStr(ar1(1))
        iCurTotalRow = iCurTotalRow + 1
   Case "<TOTAL3"
        If iCurTotalRow > 0 Then
            oDoc.Tables(2).Rows(iCurTotalRow).Select
            wm.WORDApplication.Selection.InsertRows 1
        End If
        ar1 = Split(ar(1), "|")
        oDoc.Tables(2).Cell(iCurTotalRow, 1).range = CStr(ar1(0))
        oDoc.Tables(2).Cell(iCurTotalRow, 2).Borders(-3).LineStyle = 1 'wdLineStyleDouble 7 (does not work)  'wdBorderBottom -3
        oDoc.Tables(2).Cell(iCurTotalRow, 2).Borders(-3).LineWidth = 4 'wdLineWidth050pt
        oDoc.Tables(2).Cell(iCurTotalRow, 2).range = CStr(ar1(1))
        iCurTotalRow = iCurTotalRow + 1
   Case "<TOTAL4"
        If iCurTotalRow > 0 Then
            oDoc.Tables(2).Rows(iCurTotalRow).Select
            wm.WORDApplication.Selection.InsertRows 1
        End If
        ar1 = Split(ar(1), "|")
        oDoc.Tables(2).Rows(iCurTotalRow).range.Font.Bold = False
        oDoc.Tables(2).Cell(iCurTotalRow, 1).range = CStr(ar1(0))
        oDoc.Tables(2).Cell(iCurTotalRow, 2).range = CStr(ar1(1))
        iCurTotalRow = iCurTotalRow + 1
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.ParseLine_CN(pIn,pLogoPath)", Array(pIn, pLogoPath)
End Sub
Public Sub Builddocument_TF(pFilename As String)
    On Error GoTo errHandler
Dim strIP As String
Dim strTotalValues As String
Dim i As Integer
Dim ar() As String
Dim strStyle As String
Dim strFullDeviceName As String
    
    iFilenum1 = FreeFile
    Open pFilename For Input As #iFilenum1
    iCurRow = 1
    iCurTotalRow = 1
    
'Read first line containing printer information
    Line Input #iFilenum1, strIP
    ar = Split(strIP, "|")
    
    strFullDeviceName = ar(1)
    If Not CheckPrinterExists(strFullDeviceName) Then
        WritetoErrors "Printer does not exist (printing Transfers): " & strFullDeviceName & ". Using activeprinter: " & wm.WORDApplication.ActivePrinter & vbCrLf & "Filename = " & pFilename
        strFullDeviceName = Printer.DeviceName
    End If
    
'No output device exists, exit
    If strFullDeviceName = "" Then
        Close #iFilenum1
        Exit Sub
    End If
    
'Set style
    strStyle = ar(3)
    Select Case strStyle
    Case "A"
        Set rptTF = New arTF_1
    End Select

'Read each text file line
    Do Until EOF(iFilenum1)
        Line Input #iFilenum1, strIP
        ParseLine_TF strIP
    Loop
    
'Prepare total set
    strTotalValues = ""
    For i = 1 To UBound(LineTotalArray, 1)
        strTotalValues = strTotalValues & IIf(strTotalValues > "", vbCrLf, "") & LineTotalArray(i)
    Next
    rptTF.Component LineArray, strTotalValues, strTotalDescription
    Close #iFilenum1
    
'Open active report
    If mbPreview Then
        rptTF.Show vbModal
    Else
        rptTF.Printer.DeviceName = strFullDeviceName     'ar(1) ') 'ar(1) GetFullDeviceNameFromName(
        rptTF.Printer.Copies = ar(4)
        rptTF.PrintReport False
    End If

    Set rptTF = Nothing
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrCloseFiles iFilenum1
    ErrorIn "ManageDocument.Builddocument_TF(pFilename)", Array(pFilename)
End Sub
Private Sub ParseLine_TF(pIn As String)
    On Error GoTo errHandler

Dim ar() As String
Dim ar1() As String
Dim Tag As String
Dim tmp1 As String

    ar = Split(pIn, ">")
    Tag = ar(0)
    Select Case Tag
    Case "<LINECOUNT"
        lngLineCount = CLng(ar(1))
        ReDim LineArray(lngLineCount)
    Case "<LINETOTALCOUNT"
        lngLineTotalCount = CLng(ar(1))
        ReDim LineTotalArray(lngLineTotalCount)
    Case "<DOCUMENTCODE"
        rptTF.Sections("PAGEHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
    Case "<STATUS"
        rptTF.Sections("PAGEHEADER").Controls("STATUS") = CStr(ar(1))
    Case "<OURADDRESS"
        rptTF.Sections("PAGEHEADER").Controls("OURADDRESS") = UnpackText(ar(1))
    Case "<PHONE"
        rptTF.Sections("PAGEHEADER").Controls("PHONE") = CStr(ar(1))
    Case "<STATUS"
    Case "<BILLTO"
        rptTF.Sections("PAGEHEADER").Controls("BILLTO") = UnpackText(ar(1))
    Case "<DELTO"
        rptTF.Sections("PAGEHEADER").Controls("DELTO") = UnpackText(ar(1))
    Case "<FAX"
        rptTF.Sections("PAGEHEADER").Controls("FAX") = CStr(ar(1))
    Case "<SM"
        rptTF.fSM = CStr(ar(1))
    Case "<ACCNUMBER"
    Case "<DATE"
        rptTF.Sections("PAGEHEADER").Controls("DATE") = CStr(ar(1))
    Case "<CUSTOMER"
        rptTF.Sections("PAGEHEADER").Controls("CUSTOMER") = CStr(ar(1))
    Case "<Line"
        LineArray(iCurRow) = ar(1)
        iCurRow = iCurRow + 1
    Case "<TOTAL"
        LineTotalArray(iCurTotalRow) = ar(1)
        iCurTotalRow = iCurTotalRow + 1
    Case "<TOTAL2"
        LineTotalArray(iCurTotalRow) = ar(1)
        iCurTotalRow = iCurTotalRow + 1
    Case "<TOTAL3"
        LineTotalArray(iCurTotalRow) = ar(1)
        iCurTotalRow = iCurTotalRow + 1
    Case "<TOTAL1"
        LineTotalArray(iCurTotalRow) = ar(1)
        iCurTotalRow = iCurTotalRow + 1
    Case "<TOTALDESCRIPTION"
        strTotalDescription = ar(1)
        strTotalDescription = Replace(strTotalDescription, "|", vbCrLf)
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrCloseFiles iFilenum1
    ErrorIn "ManageDocument.ParseLine_TF(pIn)", pIn
End Sub

Public Sub Builddocument_AR(pFilename As String)
    On Error GoTo errHandler
Dim strIP As String
Dim ar() As String
Dim strStyle As String
Dim strFullDeviceName As String

    iFilenum1 = FreeFile
    Open pFilename For Input As #iFilenum1
    iCurRow = 1
    iCurTotalRow = 1
    
    Line Input #iFilenum1, strIP
    ar = Split(strIP, "|")
    
    strFullDeviceName = ar(1)
    If Not CheckPrinterExists(strFullDeviceName) Then
        WritetoErrors "Printer does not exist (printing appro return): " & strFullDeviceName & ". Using activeprinter: " & wm.WORDApplication.ActivePrinter & vbCrLf & "Filename = " & pFilename
        strFullDeviceName = Printer.DeviceName
    End If
    
    If strFullDeviceName = "" Then
        Close #iFilenum1
        Exit Sub
    End If
    
    strStyle = ar(3)
    
    Select Case strStyle
    Case "A"
        Set rptAR = New arAPPR_1
    End Select
    
    Do Until EOF(iFilenum1)
        Line Input #iFilenum1, strIP
        ParseLine_AR strIP
    Loop
  '  ReplaceMarker "<STATUS>", ""
    rptAR.Component LineArray, LineTotalArray
    Close #iFilenum1
    If mbPreview Then
        rptAR.Show vbModal
    Else
        rptAR.Printer.DeviceName = strFullDeviceName   'Printer.DeviceName
        rptAR.Printer.Copies = ar(4)
        rptAR.PrintReport False
    End If
    Set rptAR = Nothing
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrCloseFiles iFilenum1
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
    Case "<LINECOUNT"
        lngLineCount = CLng(ar(1))
        ReDim LineArray(lngLineCount)
    Case "<LINETOTALCOUNT"
        lngLineTotalCount = CLng(ar(1))
        ReDim LineTotalArray(lngLineTotalCount)
    Case "<DOCUMENTCODE"
        rptAR.Sections("PAGEHEADER").Controls("DOCUMENTCODE") = CStr(ar(1))
    Case "<COMPANYNAME"
        rptAR.Sections("PAGEHEADER").Controls("FFROM") = CStr(ar(1))
    Case "<PHONE"
        rptAR.Sections("PAGEHEADER").Controls("PHONE") = CStr(ar(1))
    Case "<STATUS"
      '  rptPO.Fields("STATUS") = CStr(ar(1))
    Case "<BILLTO"
     '  tmp1 = Replace(ar(1), "%", Chr(13))
        rptAP.Sections("PAGEHEADER").Controls("BILLTO") = UnpackText(ar(1))
    Case "<DELTO"
     '   tmp1 = Replace(ar(1), "%", Chr(13))
        rptAR.Sections("PAGEHEADER").Controls("DELTO") = UnpackText(ar(1))
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

Private Function CheckPrinterExists(pDevice As String) As Boolean
    On Error GoTo errHandler
    CheckPrinterExists = False
    For Each p In Printers
        If p.DeviceName Like pDevice & "*" Then
            CheckPrinterExists = True
           ' pDevice = pDevice & " on " & p.Port
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.CheckPrinterExists(pDevice)", pDevice
End Function

Private Sub LoadFormats()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fc, f, fol

    Set fol = fs.GetFolder(strTemplateFolder)
    Set fc = fol.Files
    For Each f In fc

        Select Case f.Name
        Case "INVOICE.LAY"
            Set rptINVCUSTOM = New arINV_2
            LoadInvoiceFormat f.Path
        End Select
    Next
    For Each f In fc
        If UCase(f.Name) = "LOGO.BMP" Or UCase(f.Name) = "LOGO.JPG" Then
            rptINVCUSTOM.Sections("REPORTHEADER").Controls("fLOGO").Picture = LoadPicture(f.Path)
            rptINVCUSTOM.Sections("REPORTHEADER").Controls("fLOGO").PictureAlignment = ddPATopRight
        End If
    Next
   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.LoadFormats"
End Sub

Private Sub LoadInvoiceFormat(pFilename As String)
    On Error GoTo errHandler
Dim oTF As z_TextFile
Dim ar() As String
Dim ard() As String
Dim arprop() As String
Dim strLine As String
Dim arINV As arINV_2
Dim i As Integer

    Set oTF = New z_TextFile
    
    oTF.OpenTextFileToRead pFilename
    Do While Not oTF.IsEOF
        strLine = oTF.ReadLinefromTextFile
        ar = Split(strLine, "|")
        Select Case ar(0)
        Case "PAGEHEADER"
            Select Case ar(1)
            Case "BILLTO"
                ard = Split(ar(2), ",")
                For i = 0 To UBound(ard)
                    arprop = Split(ard(i), ":")
                    Select Case arprop(0)
                    Case "TOP"
                        rptINVCUSTOM.Sections("PAGEHEADER").Controls("BILLTO").Top = arprop(1)
                        
                    Case "LEFT"
                         rptINVCUSTOM.Sections("PAGEHEADER").Controls("BILLTO").Left = arprop(1)
                    Case "WIDTH"
                         rptINVCUSTOM.Sections("PAGEHEADER").Controls("BILLTO").Width = arprop(1)
                    End Select
                Next i
            End Select
        End Select
    Loop
    oTF.CloseTextFile
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ManageDocument.LoadInvoiceFormat(pFilename)", pFilename
End Sub

Public Sub InitializeManager(pTemplateFolder)
    strTemplateFolder = pTemplateFolder
    LoadFormats

End Sub
