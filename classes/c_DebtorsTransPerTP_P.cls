VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_DebtorsTransPerTP_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(TPID As Long, pSince As Date, Optional pACNO As String) As String
Dim lngCount As Long
Dim udtProps As DebtorsProps
Dim udtData As DebtorsData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim objDisplay As d_DebtorsPerTP
Dim lngID As Long
Dim objPB As PropertyBag
Dim OpenResult As Integer

    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If TPID = 0 And pACNO > "" Then
        rs.open "SELECT * FROM zDEBTORS_BalBF WHERE TP_ACNO = '" & pACNO & "' AND TR_DATE > '" & ReverseDate(pSince) & "' AND ISNULL(Bal3,0) <> 0  ORDER BY TR_DATE,TR_CAPTUREDATE DESC", oPC.COShort, adOpenForwardOnly
    Else
        rs.open "SELECT * FROM zDEBTORS_BalBF WHERE TPID = " & TPID & " AND TR_DATE > '" & ReverseDate(pSince) & "' AND ISNULL(Bal3,0) <> 0 ORDER BY TR_DATE,TR_CAPTUREDATE DESC", oPC.COShort, adOpenForwardOnly
    End If
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    On Err GoTo H
    Do While Not rs.eof
        With udtProps
            udtProps.TRCODE = FNS(rs!TR_CODE)
            udtProps.TRDATE = FND(rs!TR_DATE)
            udtProps.TRCaptureDATE = FND(rs!TR_CaptureDate)
            udtProps.TRID = FNN(rs!TR_ID)
            udtProps.DocType = FNS(rs!DocType)
            udtProps.Memo = FNS(rs!TR_NOTE)
            If FNS(udtProps.DocType) = "SE" Or FNS(udtProps.DocType) = "CN" Or FNS(udtProps.DocType) = "PA" Or (FNS(udtProps.DocType) = "JNL" And FNDBL(rs!Bal3) < 0) Then
                udtProps.Credit = FNDBL(rs!Bal3) * -1
                udtProps.Debit = 0
            Else
                udtProps.Debit = FNDBL(rs!Bal3)
                udtProps.Credit = 0
            End If
            udtProps.PayableAmount = FNDBL(rs!Bal3)
            udtProps.VATAmount = FNDBL(rs!Bal2)
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
        End With
    Loop
    With objPB
          .WriteProperty "Count", lngCount
          Fetch = .Contents
    End With
    Set objPB = Nothing
    
EXIT_Handler:
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
    Resume

End Function


Public Function FetchOpenItem(TPID As Long, dteFrom As Date, dteTo As Date, Optional pACNO As String) As String
Dim lngCount As Long
Dim udtProps As DebtorsProps
Dim udtData As DebtorsData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim objDisplay As d_DebtorsPerTP
Dim lngID As Long
Dim objPB As PropertyBag
Dim OpenResult As Integer

    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If TPID = 0 And pACNO > "" Then
        rs.open "SELECT * FROM vOpenItemAll WHERE TP_Acno = '" & pACNO & "' AND TR_DATE BETWEEN '" & ReverseDate(dteFrom) & "' AND '" & ReverseDate(dteTo) & "'", oPC.COShort, adOpenForwardOnly
    Else
    rs.open "SELECT * FROM vOpenItemAll WHERE TPID = " & TPID & " AND TR_DATE BETWEEN '" & ReverseDate(dteFrom) & "' AND '" & ReverseDate(dteTo) & "'", oPC.COShort, adOpenForwardOnly
    End If
    
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        FetchOpenItem = objPB.Contents
        Set objPB = Nothing
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    On Err GoTo H
    Do While Not rs.eof
        With udtProps
            .dbDoc = FNS(rs!dbDoc)
            .dbDate = FNS(rs!dbDocDate)
            .dbDocType = FNS(rs!dbDocType)
            .dbAmt = FNDBL(rs!dbAmt)
            .crDoc = FNS(rs!crDoc)
            .crDate = FNS(rs!crDocDate)
            .crDocType = FNS(rs!crDocType)
            .crAmt = FNS(rs!crAmt)
            .crTotal = FNS(rs!crTotal)
            .Balance_OI = FNS(rs!Balance)
            .TRID = FNN(rs!InvoiceID)
'            If lngMonthsBack = 0 Then
  '              .PayableAmount = FNDBL(rs!Owing)
'            End If
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
        End With
    Loop
    With objPB
          .WriteProperty "Count", lngCount
          FetchOpenItem = .Contents
    End With
    Set objPB = Nothing
    
EXIT_Handler:
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
    Resume

End Function

