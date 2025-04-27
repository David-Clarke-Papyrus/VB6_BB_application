VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Cashups_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function Fetch(pBranchCode As String, pDiscrepancy As String, pTIllpoint As String, dteFrom As Date, dteTo As Date, pStatus As String) As String
    On Error GoTo ErrHandler
Dim udtProps As CashupProps
Dim udtData As CashupData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim lngID As Long
Dim oPB As PropertyBag
Dim oPBSet As PropertyBag
Dim strSQL As String
Dim bNoRS As Boolean
Dim OpenResult As Integer
Dim oSQL As New z_SQL
Dim lngCount As Long

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set oPBSet = New PropertyBag
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    oSQL.LoadBlindCashups rs, pBranchCode, pDiscrepancy, pTIllpoint, dteFrom, dteTo, pStatus

    If rs.EOF Then
        pblnNoRecsReturned = True
        oPBSet.WriteProperty "Count", 0
        Fetch = oPBSet.Contents
        Set oPBSet = Nothing
        Exit Function
    End If

    lngCount = 0
    Set oPB = New PropertyBag
    Do While Not rs.EOF
            udtProps.XID = FNS(rs!CU_XID)
            udtProps.BranchCode = FNS(rs!CU_BRANCHCODE)
            udtProps.Tillpoint = FNS(rs!CU_Tillpoint)
            udtProps.OpenSessionTime = FND(rs!CU_OpenSessionTime)
            udtProps.CloseSessionTime = FND(rs!CU_CloseSessionTime)
            udtProps.CapturedBy = FNS(rs!CU_CapturedBy)
            udtProps.CapturedDate = FND(rs!CU_CaptureDate)
            udtProps.IssuedBy = FNS(rs!CU_IssuedBy)
            udtProps.IssuedDate = FND(rs!CU_IssueDate)
            udtProps.ExplainedBy = FNS(rs!CU_ExplainedBy)
            udtProps.ExplainedDate = FND(rs!CU_ExplainDate)
            udtProps.OpeningFloat = FNDBL(rs!CU_OpeningFloat)
            udtProps.ClosingFloat = FNDBL(rs!CU_ClosingFloat)
            udtProps.Cash = FNDBL(rs!CU_Cash)
            udtProps.Cheques = FNDBL(rs!CU_Cheques)
            udtProps.CreditCards = FNDBL(rs!CU_CreditCards)
            udtProps.DebitCards = FNDBL(rs!CU_DebitCards)
            udtProps.DirectDeposits = FNDBL(rs!CU_DirectDeposits)
            udtProps.VouchersRedeemed = FNDBL(rs!CU_Vouchers)
            udtProps.FloatBreakdownAtEnd = FNS(rs!CU_FloatBreakdownAtEnd)
            udtProps.Explanation = FNS(rs!CU_Explanation)
            udtProps.DiscrepancyCash = FNDBL(rs!CU_DiscrepancyCash)
            udtProps.DiscrepancyCheques = FNDBL(rs!CU_DiscrepancyCheques)
            udtProps.DiscrepancyCards = FNDBL(rs!CU_DiscrepancyCards)
            udtProps.DiscrepancyVouchers = FNDBL(rs!CU_DiscrepancyVouchersRedeemed)
            udtProps.DiscrepancyDeposits = FNDBL(rs!CU_DiscrepancyDeposits)
            udtProps.DiscrepancyFloat = FNDBL(rs!CU_DiscrepancyFloat)
            udtProps.DiscrepancyTotal = FNDBL(rs!TotalDiscrepancy)
            udtProps.STATUS = FNS(rs!STATUS)
            udtProps.StatusDate = FND(rs!StatusDate)
            udtProps.StatusSignature = FNS(rs!StatusSignature)
            udtProps.Wages = FNDBL(rs!CU_WAGES)
            udtProps.LeavePay = FNDBL(rs!CU_LeavePay)
            udtProps.SickLeave = FNDBL(rs!CU_SickLeave)
            udtProps.TotalSales = FNDBL(rs!CU_TotalSales)
            udtProps.COGS = FNDBL(rs!CU_COGS)
            udtProps.Retained = FNDBL(rs!CU_Retained)
            udtProps.Returned = FNDBL(rs!CU_Returned)
            udtProps.GiftVouchersSold = FNDBL(rs!CU_GiftVouchersSold)
            udtProps.OtherVouchersSold = FNDBL(rs!CU_OtherVouchersSold)
            udtProps.BankedAfterAdjustments = FNDBL(rs!CU_BankedAfterAdjustments)
            LSet udtData = udtProps
            lngCount = lngCount + 1
            oPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
    Loop
    With oPB
          .WriteProperty "Count", lngCount
          Fetch = .Contents
    End With
    Set oPB = Nothing

EXIT_Handler:
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Cashups_P.Fetch(dteFrom,dteTo,pBranchCode,pStatus,pDiscrepancy)", Array(dteFrom, dteTo, _
         pBranchCode, pStatus, pDiscrepancy)
End Function

Public Function FetchRecordset(pBranchCode As String, pDiscrepancy As String, pTIllpoint As String, dteFrom As Date, dteTo As Date, pStatus As String, prs As ADODB.Recordset)
    On Error GoTo ErrHandler
Dim OpenResult As Integer
Dim oSQL As New z_SQL
Dim lngCount As Long
Dim rs As ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    oSQL.LoadBlindCashups rs, pBranchCode, pDiscrepancy, pTIllpoint, dteFrom, dteTo, pStatus
    rs.ActiveConnection = Nothing
    Set prs = rs

'
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open cmd, , adOpenStatic
'    Set rs.ActiveConnection = Nothing
'    Set cmd = Nothing
'    Set GetTrackingActionsDetailsbyPOL = rs
'


EXIT_Handler:
   ' rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Cashups_P.Fetch(dteFrom,dteTo,pBranchCode,pStatus,pDiscrepancy)", Array(dteFrom, dteTo, _
         pBranchCode, pStatus, pDiscrepancy)
End Function


