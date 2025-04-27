VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_QU_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByRef pblnNoRecsReturned As Boolean, Optional plngTPID As Long, Optional pACNO As String, Optional pTRCode As String, _
                Optional pdteDate1 As Date, Optional pdteDate2 As Date, Optional pPID As String, Optional pPCode As String, Optional pUNissued As Boolean) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL
Dim udtProps As dQUProps
Dim udtData As dQUData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim strDateLimits As String
Dim bIncludesLineDetails As Boolean
Dim OpenResult As Integer

    pblnNoRecsReturned = False
    Set rs = New ADODB.Recordset
    Set oBatch = New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strDateLimits = "TR_CaptureDate >= {d '" & ReverseDate(pdteDate1) & "'} AND TRDATE <= {d '" & ReverseDate(pdteDate2) & "'}"
    If plngTPID > 0 And Len(pPID) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM vQuotationPerPID WHERE " & "P_ID = '" & pPID & "' AND TR_TP_ID = " & plngTPID & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf plngTPID > 0 And Len(pPCode) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM vQuotationPerPID WHERE " & "P_Code = '" & CLARG(pPID) & "' AND TR_TP_ID = " & plngTPID & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf Len(pPCode) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM vQuotationPerPID WHERE " & "P_Code = '" & CLARG(pPCode) & "' OR P_EAN = '" & CLARG(pPCode) & "' ORDER BY TRDATE DESC", enText, Array(), "", rs)
        bIncludesLineDetails = True
    ElseIf Len(pTRCode) > 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM vQuotationPerTP WHERE " & "TR_Code = '" & CLARG(pTRCode) & "' ORDER BY TRDATE DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf plngTPID <> 0 Then
        lngResult = oBatch.RunGetRecordset("Select * FROM vQuotationPerTP WHERE " & strDateLimits & " AND TR_TP_ID = " & plngTPID & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pACNO > " " Then
        lngResult = oBatch.RunGetRecordset("Select * FROM vQuotationPerTP WHERE " & strDateLimits & " AND ACCNUM = '" & CLARG(pACNO) & "' ORDER BY TRDATE DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pdteDate1 >= #1/1/1995# Then
        lngResult = oBatch.RunGetRecordset("Select * FROM vQuotationPerTP WHERE " & strDateLimits & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    ElseIf pUNissued Then
        lngResult = oBatch.RunGetRecordset("Select * FROM vQuotationPerTP WHERE TR_STATUS = 2 ORDER BY TR_CaptureDate DESC", enText, Array(), "", rs)
        bIncludesLineDetails = False
    End If
    Set objPB = New PropertyBag
    
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    On Err GoTo H
    Do While Not rs.eof
        With udtProps
            udtProps.TRID = rs!TransactionID
            udtProps.TPNAME = FNS(rs!TPNAME) & IIf(rs!TP_Title > "", ", " & rs!TP_Title, "") & IIf(rs!TP_Initials > "", " " & rs!TP_Initials, "")
            udtProps.CustomerDisplay = FNS(rs!CustomerDisplay) ' & IIf(rs!TP_Title > "", ", " & rs!TP_Title, "") & IIf(rs!TP_Initials > "", " " & rs!TP_Initials, "")
            udtProps.DOCCode = FNS(rs!TR_CODE)
            udtProps.DOCDate = FND(rs!TRDATE)
            udtProps.CaptureDate = FND(rs!TR_CaptureDate)
            udtProps.TPAccNo = FNS(rs!Accnum)
            udtProps.Status = FNN(rs!TR_Status)
            udtProps.StaffID = FNN(rs!TR_StaffID)
            udtProps.SM = FNS(rs!SM_SHortName)
            udtProps.Log = Left(FNS(rs!TR_Log), 450)
            If bIncludesLineDetails Then
                udtProps.Price = FNN(rs!QUL_Price)
                udtProps.Qty = FNN(rs!QUL_QTY)
                udtProps.Discount = FNDBL(rs!QUL_DiscountPercent)
                udtProps.VATRate = FNDBL(rs!QUL_VATRATE)
                udtProps.QULID = FNN(rs!QUL_ID)
            End If
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
            If lngCount > MAXBROWSE Then Exit Do
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
    Set oBatch = Nothing
    Exit Function
H:
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_QU_P.Fetch(pblnNoRecsReturned,plngTPID,pACNO,pTRCode,pdteDate1,pdteDate2,pPID,pPCode," & _
        "pUNissued)", Array(pblnNoRecsReturned, plngTPID, pACNO, pTRCode, pdteDate1, pdteDate2, pPID, pPCode, _
         pUNissued), EA_NORERAISE
End Function



