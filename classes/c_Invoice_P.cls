VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Invoices_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByRef pblnNoRecsReturned As Boolean, Optional plngTPID As Long, Optional pACNO As String, Optional pLineRef As String, Optional pDocRef As String, _
                Optional pdteDate1 As Date, Optional pdteDate2 As Date, Optional pPID As String, Optional pPCode As String, Optional pPIID As Long, _
                Optional pPF As Boolean, Optional UseIssuedDate As Boolean = False, Optional pUNissued As Boolean) As String
Dim rs As ADODB.Recordset
Dim oBatch As z_SQL
Dim udtProps As dInvoiceProps
Dim udtData As dInvoiceData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim strDateLimits As String
Dim bIncludesLineDetails As Boolean
Dim OpenResult As Integer
Dim strProformaRestriction As String
Dim sLimitSQL As String

    pblnNoRecsReturned = False
    Set rs = New ADODB.Recordset
    Set oBatch = New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If pPF = True Then
        strProformaRestriction = "I_PROFORMA = 1 AND "
    ElseIf pPF = False Then
        strProformaRestriction = "(I_PROFORMA = 0) AND "
    Else
        strProformaRestriction = ""
    End If
'    If AllGDNs = False Then
'        sLimitSQL = " AND  DN_I_ID > 0"
'    Else
        sLimitSQL = ""
'    End If
    If UseIssuedDate = False Then
       ' strDateLimits = "TR_CaptureDate >= {d '" & ReverseDate(pdteDate1) & "'} AND TRDATE <= {d '" & ReverseDate(pdteDate2) & "'}"
        strDateLimits = "TR_CaptureDate >= dbo.StartOfDay('" & ReverseDate(pdteDate1) & "') AND TR_CaptureDate <= dbo.EndOfDay('" & ReverseDate(pdteDate2) & "')"
        If plngTPID > 0 And Len(pPID) > 0 Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerPID WHERE " & strProformaRestriction & "P_ID = '" & pPID & "' AND TR_TP_ID = " & plngTPID & sLimitSQL & "  ORDER BY TRDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = True
        ElseIf plngTPID > 0 And pPIID > 0 Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerPID WHERE " & strProformaRestriction & "iL_PI_ID = " & pPIID & " AND TR_TP_ID = " & plngTPID & sLimitSQL & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = True
        ElseIf plngTPID > 0 And Len(pPCode) > 0 Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerPID WHERE " & strProformaRestriction & "P_Code = '" & CLARG(pPID) & "' AND TR_TP_ID = " & plngTPID & sLimitSQL & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = True
        ElseIf Len(pPCode) > 0 Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerPID WHERE " & strProformaRestriction & "P_Code = '" & CLARG(pPCode) & "' OR P_EAN = '" & CLARG(pPCode) & "'" & sLimitSQL & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = True
        ElseIf Len(pLineRef) > 0 Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerPID WHERE " & strProformaRestriction & "IL_REF = '" & CLARG(pLineRef) & "'" & sLimitSQL & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = True
        ElseIf Len(pDocRef) > 0 Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerTP WHERE " & strProformaRestriction & "TR_Code = '" & CLARG(pDocRef) & "'" & sLimitSQL & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = False
        ElseIf plngTPID <> 0 Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerTP WHERE " & strProformaRestriction & strDateLimits & " AND TR_TP_ID = " & plngTPID & sLimitSQL & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = False
        ElseIf pACNO > " " Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerTP WHERE " & strProformaRestriction & strDateLimits & " AND ACCNUM = '" & CLARG(pACNO) & "'" & sLimitSQL & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = False
        ElseIf pdteDate1 >= #1/1/1995# Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerTP WHERE " & strProformaRestriction & strDateLimits & sLimitSQL & " ORDER BY TRDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = False
        ElseIf pUNissued Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerTP WHERE " & strProformaRestriction & " TR_STATUS = 2  " & sLimitSQL & " ORDER BY TR_CaptureDate DESC", enText, Array(), "", rs)
            bIncludesLineDetails = False
        Else
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerTP WHERE  I_PROFORMA = 1 " & sLimitSQL & " ORDER BY TRDATE", enText, Array(), "", rs)
            bIncludesLineDetails = False
        End If
    Else
        strDateLimits = "TR_PROCESSINGDATE >= dbo.StartOfDay('" & ReverseDate(pdteDate1) & "') AND TR_PROCESSINGDATE <= dbo.EndOfDay('" & ReverseDate(pdteDate2) & "')"
        If plngTPID > 0 And Len(pPID) > 0 Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerPID WHERE " & strProformaRestriction & "P_ID = '" & pPID & "' AND TR_TP_ID = " & plngTPID & sLimitSQL & " ORDER BY TR_PROCESSINGDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = True
        ElseIf plngTPID > 0 And pPIID > 0 Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerPID WHERE " & strProformaRestriction & "iL_PI_ID = " & pPIID & " AND TR_TP_ID = " & plngTPID & sLimitSQL & " ORDER BY TR_PROCESSINGDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = True
        ElseIf plngTPID > 0 And Len(pPCode) > 0 Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerPID WHERE " & strProformaRestriction & "P_Code = '" & CLARG(pPID) & "' AND TR_TP_ID = " & plngTPID & sLimitSQL & " ORDER BY TR_PROCESSINGDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = True
        ElseIf Len(pPCode) > 0 Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerPID WHERE " & strProformaRestriction & "P_Code = '" & CLARG(pPCode) & "' OR P_EAN = '" & CLARG(pPCode) & "'" & sLimitSQL & " ORDER BY TR_PROCESSINGDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = True
        ElseIf Len(pLineRef) > 0 Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerPID WHERE " & strProformaRestriction & "IL_REF = '" & CLARG(pLineRef) & "'" & sLimitSQL & " ORDER BY TR_PROCESSINGDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = True
        ElseIf Len(pDocRef) > 0 Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerTP WHERE " & strProformaRestriction & "TR_Code = '" & CLARG(pDocRef) & "'" & sLimitSQL & " ORDER BY TR_PROCESSINGDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = False
        ElseIf plngTPID <> 0 Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerTP WHERE " & strProformaRestriction & strDateLimits & " AND TR_TP_ID = " & plngTPID & sLimitSQL & " ORDER BY TR_PROCESSINGDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = False
        ElseIf pACNO > " " Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerTP WHERE " & strProformaRestriction & strDateLimits & " AND ACCNUM = '" & CLARG(pACNO) & "'" & sLimitSQL & " ORDER BY TR_PROCESSINGDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = False
        ElseIf pdteDate1 >= #1/1/1995# Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerTP WHERE " & strProformaRestriction & strDateLimits & sLimitSQL & " ORDER BY TR_PROCESSINGDATE DESC", enText, Array(), "", rs)
            bIncludesLineDetails = False
        ElseIf pUNissued Then
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerTP WHERE " & strProformaRestriction & " TR_STATUS = 2 " & sLimitSQL & " ORDER BY TR_CaptureDate DESC", enText, Array(), "", rs)
            bIncludesLineDetails = False
        Else
            lngResult = oBatch.RunGetRecordset("Select * FROM InvoicesPerTP WHERE  I_PROFORMA = 1 " & sLimitSQL & " ORDER BY TR_PROCESSINGDATE", enText, Array(), "", rs)
            bIncludesLineDetails = False
        End If
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
            udtProps.DOCCode = FNS(rs!TR_CODE)
            udtProps.DOCDate = FND(rs!TRDATE)
            udtProps.CustomerDisplay = FNS(rs!CustomerDisplay)
            udtProps.CaptureDate = FND(rs!TR_CaptureDate)
            udtProps.TPAccNo = FNS(rs!Accnum)
            udtProps.Status = FNN(rs!TR_Status)
            udtProps.StaffID = FNN(rs!TR_StaffID)
            udtProps.SM = FNS(rs!SM_SHortName)
            udtProps.CURRID = FNN(rs!I_Curr_ID)
            udtProps.CurrRate = FNDBL(rs!I_CUrrRate)
            udtProps.Proforma = FNB(rs!I_Proforma)
            udtProps.Log = Left(FNS(rs!TR_Log), 450)
            udtProps.InvoiceQty = FNN(rs!Qty)
            udtProps.InvoiceValue = FNDBL(rs!val)
            udtProps.IsPreDelivery = FNB(rs!I_IsPredelivery)
            udtProps.Exchange = FNS(rs!Exchange)
            If bIncludesLineDetails Then
                udtProps.PIID = FNN(rs!IL_PI_ID)
                udtProps.Price = FNN(rs!IL_Price)
                udtProps.Qty = FNN(rs!IL_QTY)
                udtProps.Discount = FNDBL(rs!IL_DiscountRate)
                udtProps.VATRate = FNDBL(rs!IL_VATRATE)
                udtProps.ILID = FNN(rs!IL_ID)
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
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Function
