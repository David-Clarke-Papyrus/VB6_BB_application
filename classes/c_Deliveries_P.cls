VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_DEL_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByRef pblnNoRecsReturned As Boolean, Optional plngTPID As Long, Optional pACNO As String, Optional pTRCode As String, _
                Optional pdteDate1 As Date, Optional pdteDate2 As Date, Optional pPID As String, Optional pPCode As String, _
                Optional pPIID As Long, Optional pSUppInvNo As String, Optional pUNissued As Boolean) As String
10        On Error GoTo errHandler
      Dim rs As ADODB.Recordset
      Dim oBatch As z_SQL
      Dim udtProps As dDELProps
      Dim udtData As dDELData
      Dim objPB As PropertyBag
      Dim lngCount As Long
      Dim bIncludesLineDetails As Boolean
      Dim strDateCriteria As String
      Dim OpenResult As Integer
20        errSysHandlerSet

30        strDateCriteria = " TRDate>= {d '" & ReverseDate(pdteDate1) & "'} AND TRDate <= {d '" & ReverseDate(pdteDate2) & "'}"

40        Set rs = New ADODB.Recordset
50        rs.CursorLocation = adUseClient
60        Set oBatch = New z_SQL
      '-------------------------------
70        OpenResult = oPC.OpenDBSHort
      '-------------------------------
80        If plngTPID > 0 And Len(pPID) > 0 Then
90            lngResult = oBatch.RunGetRecordset("Select * FROM DELsPerPID WHERE P_ID = '" & pPID & "' AND TR_TP_ID = " & plngTPID & " ORDER BY TRDate", enText, Array(), "", rs)
100           bIncludesLineDetails = True
110       ElseIf plngTPID > 0 And pPIID > 0 Then
120           lngResult = oBatch.RunGetRecordset("Select * FROM DELsPerPID WHERE iL_PI_ID = " & pPIID & " AND TR_TP_ID = " & plngTPID & " ORDER BY TRDate", enText, Array(), "", rs)
130           bIncludesLineDetails = True
140       ElseIf plngTPID > 0 And Len(pPCode) > 0 Then
150           lngResult = oBatch.RunGetRecordset("Select * FROM DELsPerPID WHERE P_Code = '" & CLARG(pPID) & "' AND TR_TP_ID = " & plngTPID & " ORDER BY TRDate", enText, Array(), "", rs)
160           bIncludesLineDetails = True
170       ElseIf pPCode > "" Then
180           lngResult = oBatch.RunGetRecordset("Select * FROM DELsPerPID WHERE  P_Code = '" & CLARG(pPCode) & "' OR P_EAN = '" & CLARG(pPCode) & "' ORDER BY TRDate", enText, Array(), "", rs)
190           bIncludesLineDetails = True
200       ElseIf Len(pTRCode) > 0 Then
210           lngResult = oBatch.RunGetRecordset("Select * FROM DelsPerTP WHERE TR_Code = '" & CLARG(pTRCode) & "' ORDER BY TRDate", enText, Array(), "", rs)
220           bIncludesLineDetails = False
230       ElseIf plngTPID <> 0 Then
240           lngResult = oBatch.RunGetRecordset("Select * FROM DelsPerTP WHERE " & strDateCriteria & " AND TR_TP_ID = " & plngTPID & " ORDER BY TRDate", enText, Array(), "", rs)
250           bIncludesLineDetails = False
260       ElseIf pACNO > " " Then
270           lngResult = oBatch.RunGetRecordset("Select * FROM DelsPerTP WHERE " & strDateCriteria & " AND ACCNUM = '" & CLARG(pACNO) & "' ORDER BY TRDate", enText, Array(), "", rs)
280           bIncludesLineDetails = False
290       ElseIf pSUppInvNo > "" Then
300           lngResult = oBatch.RunGetRecordset("Select * FROM DelsPerTP WHERE DEL_SupplierInvoiceRef = '" & CLARG(pSUppInvNo) & "' ORDER BY TRDate", enText, Array(), "", rs)
310           bIncludesLineDetails = False
320       ElseIf pdteDate1 >= #1/1/1995# Then
330           lngResult = oBatch.RunGetRecordset("Select * FROM DelsPerTP WHERE " & strDateCriteria & "  ORDER BY TRDate", enText, Array(), "", rs)
340           bIncludesLineDetails = False
350       ElseIf pUNissued Then
360           lngResult = oBatch.RunGetRecordset("Select * FROM DelsPerTP WHERE TR_STATUS = 2 ORDER BY TR_CaptureDate DESC", enText, Array(), "", rs)
370           bIncludesLineDetails = False
380       Else
390           lngResult = oBatch.RunGetRecordset("Select * FROM DelsPerTP ORDER BY TRDate", enText, Array(), "", rs)
400           bIncludesLineDetails = False
410       End If
420       Set objPB = New PropertyBag

          
          Dim i As Long
430       i = 0
440       pblnNoRecsReturned = True
450       Do While Not rs.eof

460               pblnNoRecsReturned = False
470           With udtProps
480               udtProps.TRID = rs!TransactionID
490               udtProps.TPNAME = Left(FNS(rs!TPNAME), 100)
500               udtProps.DOCCode = Left(FNS(rs!TR_CODE), 14)
510               udtProps.DOCDate = FND(rs!TRDATE)
520               udtProps.TPAccNo = Left(FNS(rs!Accnum), 10)
530               udtProps.CaptureDate = FND(rs!TR_CaptureDate)
540               udtProps.Status = FNS(rs.Fields("TR_Status"))
550               udtProps.StaffID = FNN(rs.Fields("TR_STAFFID"))
560               udtProps.InvoiceRef = FNS(rs.Fields("DEL_SupplierInvoiceRef"))
570               udtProps.InvoiceDate = FND(rs.Fields("DEL_SupplierInvoiceDate"))
580               udtProps.InvoiceValue = FNS(rs.Fields("BatchValueTotalF"))
590               udtProps.InvoiceQty = CStr(FNN(rs.Fields("BatchQtyTotal")))
600               udtProps.InvoiceShort = CStr(FNN(rs.Fields("BatchQtyShort")))

610               LSet udtData = udtProps
620               lngCount = lngCount + 1
630               objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
640               rs.MoveNext
650           End With
660           i = i + 1
670       Loop
680       With objPB
690             .WriteProperty "Count", lngCount
700             Fetch = .Contents
710       End With
720       Set objPB = Nothing
EXIT_Handler:
730       rs.Close
740       Set rs = Nothing
      '---------------------------------------------------
750       If OpenResult = 0 Then oPC.DisconnectDBShort
      '---------------------------------------------------
760       Set oBatch = Nothing
770       Exit Function

errHandler:
780       ErrPreserve
790       If Err.Number = -2147217407 Then   'Access violation
800     errRepeat = errRepeat + 1
810     LogSaveToFile "Access violation in c_DEL_P.Fetch, err repeat = " & CStr(errRepeat) & ", line:" & CStr(Erl())
820     If errRepeat < 5 Then
830         Resume Next
840     Else
850         LogSaveToFile "Access violation in c_DEL_P.Fetch after 5 re-attempts"
860         MsgBox "Memory error in z_StockManager:ActionODCOL. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't run search."
870         Err.Clear
880         Exit Function
890     End If
900       End If
910       If ErrMustStop Then Debug.Assert False: Resume
920       ErrorIn "c_DEL_P.Fetch(pblnNoRecsReturned,plngTPID,pACNO,pTRCode,pdteDate1,pdteDate2,pPID,pPCode," & _
              "pPIID,pSUppInvNo)", Array(pblnNoRecsReturned, plngTPID, pACNO, pTRCode, pdteDate1, pdteDate2, pPID, _
               pPCode, pPIID, pSUppInvNo), , , "line number", Array(Erl())
End Function


