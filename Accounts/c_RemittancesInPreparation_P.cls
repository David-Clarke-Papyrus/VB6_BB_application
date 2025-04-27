VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_RemittancesInPreparation_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(DueDays As Long, bSelected As Boolean, Optional lngTPID As Long) As String
Dim lngCount As Long
Dim udtProps As DebtorsProps
Dim udtData As DebtorsData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim dCreditorsTransaction As d_CreditorsTransaction
Dim lngID As Long
Dim objPB As PropertyBag
Dim OpenResult As Integer

    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If bSelected Then
        rs.open "SELECT * FROM vRemittancePreparation Where Dateadd(d," & CStr(DueDays) & ",GetDate()) > DueDate " _
                & " and TPID = " & CStr(lngTPID) & " ORDER BY TRDATE,CAPTUREDATE DESC,TRID,DocType DESC", _
                oPC.COShort, adOpenForwardOnly
    Else
        rs.open "SELECT * FROM vRemittancePreparation Where Dateadd(d," & CStr(DueDays) & ",GetDate()) > DueDate " _
                & " ORDER BY TRDATE,CAPTUREDATE DESC,TRID,DocType DESC", _
                oPC.COShort, adOpenForwardOnly
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
            udtProps.TRCODE = FNS(rs!TRCODE)
            udtProps.TRDATE = FND(rs!TRDATE)
            udtProps.TRCaptureDATE = FND(rs!CaptureDate)
            udtProps.TRID = FNN(rs!TRID)
            udtProps.TPID = FNN(rs!TPID)
            udtProps.REMID = FNN(rs!REMID)
            udtProps.DocType = FNS(rs!DocType)
            udtProps.CreditorDocDate = FND(rs!SupplierDocDate)
            udtProps.TRProcessingDate = FND(rs!ProcessingDate)
            udtProps.DueDate = FND(rs!DueDate)
            udtProps.PayableAmount = FNDBL(rs!InvoicePayable)
            udtProps.PayableAmountF = FNS(rs!InvoicePayableF)
            udtProps.PayableAfterSettDisc = FNDBL(rs!InvoicePayableAfterSettDisc)
            udtProps.PayableAfterSettDiscF = FNS(rs!InvoicePayableAfterSettDiscF)
            udtProps.EffectivePayableAfterSettDisc = FNDBL(rs!EffectivePayableAfterSettDisc)
            udtProps.EffectivePayableAfterSettDiscF = FNS(rs!EffectivePayableAfterSettDiscF)
            udtProps.SettlementDueDate = FND(rs!SettlementDueDate)
            udtProps.ClaimValue = FND(rs!ClaimValue)
            udtProps.ClaimValueF = FNS(rs!ClaimValueF)
            udtProps.Owing = FND(rs!Owing)
            udtProps.OwingF = FNS(rs!OwingF)
            udtProps.TempRemittance = FND(rs!TempRemittance)
            udtProps.TempRemittanceF = FNS(rs!TempRemittanceF)
'            udtProps.TempRemittance = FND(rs!TempRemittance)
'            udtProps.TempRemittanceF = FNS(rs!TempRemittanceF)
            udtProps.Memo = FNS(rs!TempRemittanceNote)
            udtProps.Status = FNS(rs!TempRemittanceStatus)
            udtProps.SupplierInvoiceCode = FNS(rs!SupplierInvoiceCode)
            
'            udtProps.Memo = FNS(rs!TRNOTE)
'            If FNS(udtProps.DocType) = "CN" Or FNS(udtProps.DocType) = "PA" Or (FNS(udtProps.DocType) = "JNL" And FNDBL(rs!InvoicePayable) < 0) Then
'                udtProps.Credit = FNDBL(rs!Bal3) * -1
'                udtProps.Debit = 0
'            Else
'                udtProps.Debit = FNDBL(rs!InvoicePayable)
'                udtProps.Credit = 0
'            End If
'            udtProps.PayableAmount = FNDBL(rs!InvoicePayable)
'            udtProps.VATAmount = FNDBL(rs!LocalVATPayable)
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

