VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Exchanges_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Public Function Fetch(pOPID As String, pZID As String) As String
'On Error GoTo errHandler
'Dim lngCount As Long
'Dim udtProps As ExchangeProps
'Dim udtData As ExchangeData
'Dim pblnNoRecsReturned As Boolean
'Dim rs As ADODB.Recordset
'Dim objDisplay As d_Exchange
'Dim lngID As Long
'Dim objPB As PropertyBag
'    Set objPB = New PropertyBag
'    Set rs = New ADODB.Recordset
'    If pOPID > "" Then
'        rs.Open "SELECT * FROM tExchange LEFT JOIN tTP ON EXCH_TP_ID = TP_ID  WHERE EXCH_OPSESSIONID ='" & pOPID & "' ORDER BY EXCH_SALEDATE", oPC.COsHORT, adOpenForwardOnly
'    Else
'        rs.Open "SELECT * FROM tExchange LEFT JOIN tTP ON EXCH_TP_ID = TP_ID WHERE EXCH_ZSESSIONID ='" & pZID & "' ORDER BY EXCH_SALEDATE", oPC.COsHORT, adOpenForwardOnly
'    End If
'    If rs.EOF Then
'        pblnNoRecsReturned = True
'        objPB.WriteProperty "Count", 0
'        Fetch = objPB.Contents
'        Set objPB = Nothing
'        Exit Function
'    End If
'
'    Do While Not rs.EOF
'        With udtProps
'            udtProps.ID = FNS(rs!EXCH_ID)
'            udtProps.OPSID = FNS(rs!EXCH_OPSessionID)
'            udtProps.ZID = FNS(rs!EXCH_ZSessionID)
'            udtProps.TotalValueSales = FNN(rs!EXCH_Amount)
'            udtProps.TPID = FNN(rs!EXCH_TP_ID)
'            udtProps.CustomerName = FNS(rs!TP_Name)
'       '     udtProps.TotalValueCredit = FND(rs!Z_EndDate)
'            udtProps.TotalValueDiscount = FNN(rs!EXCH_General_Disc)
'            udtProps.SaleDate = FND(rs!EXCH_SaleDate)
'            udtProps.ChangeGiven = FNN(rs!EXCH_ChangeGiven)
'            udtProps.SalesPersonID = FNN(rs!EXCH_SupervisorID)
'            LSet udtData = udtProps
'            lngCount = lngCount + 1
'            objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
'            rs.MoveNext
'        End With
'    Loop
'    With objPB
'          .WriteProperty "Count", lngCount
'          Fetch = .Contents
'    End With
'    Set objPB = Nothing
'
'EXIT_Handler:
'    rs.Close
'    Set rs = Nothing
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "c_Exchanges_P.Fetch(pOPID As String, pZID As String)", Array(pOPID, pZID)
'End Function

Public Function Fetch(dteFrom As Date, dteTo As Date, pExchangeNumber As Long, pBranchCode As String) As String
    On Error GoTo errHandler
Dim lngCount As Long
Dim lngCount2 As Long
Dim lngCount3 As Long
Dim udtProps As ExchangeProps
Dim udtData As ExchangeData
Dim udtProps2 As CSLProps
Dim udtData2 As CSLData
Dim udtProps3 As PaymentProps
Dim udtData3 As PaymentData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim lngID As Long
Dim oPB As PropertyBag
Dim oPBCSL As PropertyBag
Dim oPBPAY As PropertyBag
Dim oPBCSLSet As PropertyBag
Dim oPBPAYSet As PropertyBag
Dim oPBSet As PropertyBag
Dim strSQL As String
Dim bNoRS As Boolean
Dim OpenResult As Integer


'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set oPBSet = New PropertyBag
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    If pExchangeNumber > 0 Then
        rs.Open "SELECT DISTINCT * FROM tExchange LEFT JOIN tTP ON CUSTOMERACNO = TP_ACNO WHERE EXCHANGEDATE >'" & ReverseDate(dteFrom) & "' AND EXCHANGEDATE < '" & ReverseDate(dteTo) & "' AND EXCHANGENUMBER  = '" & pExchangeNumber & "'  AND ISNULL(CUSTOMERACNO,'') <> '' ORDER BY EXCHANGENUMBER DESC", oPC.COShort, adOpenForwardOnly
    ElseIf pBranchCode > "" Then
        rs.Open "SELECT DISTINCT * FROM tExchange LEFT JOIN tTP ON CUSTOMERACNO = TP_ACNO WHERE EXCHANGEDATE >'" & ReverseDate(dteFrom) & "' AND EXCHANGEDATE < '" & ReverseDate(dteTo) & "' AND BRANCHCODE  = '" & pBranchCode & "' AND ISNULL(CUSTOMERACNO,'') <> '' ORDER BY EXCHANGENUMBER DESC", oPC.COShort, adOpenForwardOnly
    Else
        rs.Open "SELECT DISTINCT * FROM tExchange LEFT JOIN tTP ON CUSTOMERACNO = TP_ACNO WHERE EXCHANGEDATE >'" & ReverseDate(dteFrom) & "' AND EXCHANGEDATE < '" & ReverseDate(dteTo) & "'  AND ISNULL(CUSTOMERACNO,'') <> '' ORDER BY EXCHANGENUMBER DESC", oPC.COShort, adOpenForwardOnly
    End If
    If rs.EOF Then
        pblnNoRecsReturned = True
        oPBSet.WriteProperty "Count", 0
        Fetch = oPBSet.Contents
        Set oPBSet = Nothing
        Exit Function
    End If

    lngCount = 0
    Do While Not rs.EOF
            Set oPB = New PropertyBag
            udtProps.ID = FNS(rs!EXCHANGEID)
            udtProps.Note = FNS(rs!TP_Name)
            udtProps.ZID = FNS(rs!BRANCHCODE)
            udtProps.OPSID = FNS(rs!TILL)
            udtProps.Type = FNS(rs!ExchangeType)
            udtProps.ExchangeNumber = FNN(rs!ExchangeNumber)
            udtProps.TotalPayable = FNN(rs!SalesValue)
            udtProps.TotalDiscount = FNN(rs!DISCOUNTVALUE)
            udtProps.TotalVAT = FNN(rs!VATValue)
            udtProps.ChangeGiven = FNN(rs!ChangeGiven)
            udtProps.ExchangeDate = FND(rs!ExchangeDate)
            udtProps.SalesPersonName = FNS(rs!Operator)

            Set rs2 = New ADODB.Recordset
            strSQL = "SELECT * FROM vCSL_Sale WHERE ExchangeID = '" & FNS(udtProps.ID) & "'"
            rs2.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            lngCount2 = 0
            Set oPBCSL = Nothing
            Set oPBCSL = New PropertyBag
            Set oPBCSLSet = Nothing
            Set oPBCSLSet = New PropertyBag
            If rs2.EOF Then
                oPBCSL.WriteProperty "Count", 0
                Fetch = oPBCSL.Contents
                oPB.WriteProperty "CSLS", oPBCSL.Contents
                Set oPBCSL = Nothing
            Else
              Do While Not rs2.EOF
                  udtProps2.CodeF = FNS(rs2!CodeF)
                  udtProps2.Title = FNS(rs2!P_Title)
                  udtProps2.Author = FNS(rs2!P_MainAuthor)
                  udtProps2.Qty = FNN(rs2!Qty)
                  udtProps2.Price = FNN(rs2!Price)
                  udtProps2.VATRate = FNDBL(rs2!VATRate)
                  udtProps2.DiscountRate = FNDBL(rs2!Discount)
                  udtProps2.EXCHANGE_GUID = FNS(rs2!EXCHANGEID)
                  LSet udtData2 = udtProps2
                  lngCount2 = lngCount2 + 1
                  oPBCSL.WriteProperty "Item" & CStr(lngCount2), udtData2.buffer
                  rs2.MoveNext
              Loop
              rs2.Close
              Set rs2 = Nothing
              oPBCSL.WriteProperty "COUNT", lngCount2
              oPB.WriteProperty "CSLS", oPBCSL.Contents
              Set oPBCSL = Nothing
              Set oPBCSLSet = Nothing
            End If

            Set rs3 = New ADODB.Recordset
            rs3.Open "SELECT * FROM vPayment WHERE ExchangeID = '" & udtProps.ID & "'", oPC.COShort, adOpenKeyset, adLockOptimistic
            lngCount3 = 0
            Set oPBPAY = Nothing
            Set oPBPAY = New PropertyBag
            Set oPBPAYSet = Nothing
            Set oPBPAYSet = New PropertyBag
            If rs3.EOF Then
                oPBPAY.WriteProperty "Count", 0
                Fetch = oPBPAY.Contents
            End If
            Do While Not rs3.EOF
             '   udtProps3.PaymentID = FNS(rs3!PAY_ID)
                udtProps3.EXCHANGE_GUID = FNS(rs3!EXCHANGEID)
                udtProps3.Amt = FNN(rs3!TENDER)
             '   udtProps3.AmtTendered = FNN(rs3!PAY_Amt_Tendered)
                udtProps3.Type = FNS(rs3!PAYTYPE)
             '   udtProps3.Note = FNS(rs3!PAY_Note)
                LSet udtData3 = udtProps3
                lngCount3 = lngCount3 + 1
                oPBPAY.WriteProperty "Item" & CStr(lngCount3), udtData3.buffer
                rs3.MoveNext
            Loop
            rs3.Close
            Set rs3 = Nothing
            oPBPAY.WriteProperty "COUNT", lngCount3
            oPB.WriteProperty "PAYS", oPBPAY.Contents
            Set oPBPAY = Nothing
            Set oPBPAYSet = New PropertyBag
            
            rs.MoveNext
            LSet udtData = udtProps
            lngCount = lngCount + 1
            oPB.WriteProperty "State", udtData.buffer
            oPBSet.WriteProperty "Item" & CStr(lngCount), oPB.Contents
            Set oPB = Nothing
    Loop
    With oPBSet
          .WriteProperty "Count", lngCount
          Fetch = .Contents
    End With
    Set oPBSet = Nothing

EXIT_HANDLER:
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges_P.Fetch(dteFrom,dteTo)", Array(dteFrom, dteTo), EA_NORERAISE
End Function


Public Function FetchCSL(EXCHID As String) As String
    On Error GoTo errHandler
Dim lngCount As Long
Dim lngCount2 As Long
Dim lngCount3 As Long
Dim udtProps As CSLProps
Dim udtData As CSLData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim lngID As Long
Dim oPB As PropertyBag

Dim strSQL As String
Dim bNoRS As Boolean
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set oPB = New PropertyBag
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    strSQL = "SELECT * FROM vCSL_Sale WHERE ExchangeID = '" & EXCHID & "'"
    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
        pblnNoRecsReturned = True
        oPB.WriteProperty "Count", 0
        FetchCSL = oPB.Contents
        Set oPB = Nothing
        Exit Function
    End If

    lngCount = 0
    Do While Not rs.EOF
            udtProps.CodeF = FNS(rs!CodeF)
            udtProps.Title = FNS(rs!P_Title)
            udtProps.Author = FNS(rs!P_MainAuthor)
            udtProps.Qty = FNN(rs!Qty)
            udtProps.Price = FNN(rs!Price)
            udtProps.VATRate = FNDBL(rs!VATRate)
            udtProps.DiscountRate = FNDBL(rs!Discount)
            udtProps.EXCHANGE_GUID = FNS(rs!EXCHANGEID)
            LSet udtData = udtProps
            lngCount = lngCount + 1
            oPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
    Loop
    With oPB
          .WriteProperty "Count", lngCount
          FetchCSL = .Contents
    End With
    Set oPB = Nothing
    

EXIT_HANDLER:
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges_P.FetchCSL(EXCHID)", EXCHID
End Function
Public Function FetchPay(EXCHID As String) As String
    On Error GoTo errHandler
Dim lngCount As Long
Dim lngCount2 As Long
Dim lngCount3 As Long
Dim udtProps As PaymentProps
Dim udtData As PaymentData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim lngID As Long
Dim oPB As PropertyBag
Dim strSQL As String
Dim bNoRS As Boolean
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set oPB = New PropertyBag
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    strSQL = "SELECT * FROM tPayment WHERE EXCHANGEID = '" & EXCHID & "'"
    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
        pblnNoRecsReturned = True
        oPB.WriteProperty "Count", 0
        FetchPay = oPB.Contents
        Set oPB = Nothing
        Exit Function
    End If

    lngCount = 0
    Do While Not rs.EOF
            udtProps.Amt = FNS(rs!TENDER)
            udtProps.Type = FNS(rs!PAYTYPE)
            udtProps.EXCHANGE_GUID = FNS(rs!EXCHANGEID)
            LSet udtData = udtProps
            lngCount = lngCount + 1
            oPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
    Loop
    With oPB
          .WriteProperty "Count", lngCount
          FetchPay = .Contents
    End With
    Set oPB = Nothing
    

EXIT_HANDLER:
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges_P.FetchPAY(EXCHID)", EXCHID
End Function

Public Function FetchCustomerRequests(pOSorALL As String) As String
    On Error GoTo errHandler
Dim lngCount As Long
Dim lngCount2 As Long
Dim lngCount3 As Long
Dim udtProps As ExchangeProps
Dim udtData As ExchangeData
Dim udtProps2 As CSLProps
Dim udtData2 As CSLData
Dim udtProps3 As PaymentProps
Dim udtData3 As PaymentData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim lngID As Long
Dim oPB As PropertyBag
Dim oPBCSL As PropertyBag
Dim oPBPAY As PropertyBag
Dim oPBCSLSet As PropertyBag
Dim oPBPAYSet As PropertyBag
Dim oPBSet As PropertyBag
Dim strSQL As String
Dim bNoRS As Boolean
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set oPBSet = New PropertyBag
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    bNoRS = False
    If pOSorALL = "OS" Then
        rs.Open "SELECT * FROM tExchange Left JOIN tStaffMember ON EXCH_SUPERVISORID = SM_ID WHERE EXCH_Type ='OR' AND ISNULL(EXCH_STATUS,0) = 0 AND DATEDIFF(d,EXCH_SALEDATE,GETDATE()) < 45 ORDER BY EXCH_SALEDATE DESC", oPC.COShort, adOpenForwardOnly
    Else
        rs.Open "SELECT  * FROM tExchange Left JOIN tStaffMember ON EXCH_SUPERVISORID = SM_ID WHERE EXCH_Type ='OR' AND DATEDIFF(d,EXCH_SALEDATE,GETDATE()) < 45 ORDER BY EXCH_SALEDATE DESC", oPC.COShort, adOpenForwardOnly
    End If
    If rs.EOF Then
        pblnNoRecsReturned = True
        oPBSet.WriteProperty "Count", 0
        FetchCustomerRequests = oPBSet.Contents
        Set oPBSet = Nothing
        Exit Function
    End If

    lngCount = 0
    Do While Not rs.EOF
            Set oPB = New PropertyBag
            udtProps.ID = FNS(rs!EXCH_ID)
            udtProps.OPSID = FNS(rs!EXCH_OPSessionID)
            udtProps.ZID = FNS(rs!EXCH_ZSessionID)
            udtProps.TotalPayable = FNN(rs!EXCH_SALESVALUE)
            udtProps.TPID = FNN(rs!EXCH_TP_ID)
            udtProps.TotalDiscount = FNN(rs!EXCH_DISCOUNTVALUE)
            udtProps.TotalVAT = FNN(rs!EXCH_VATVALUE)
            udtProps.Type = FNS(rs!EXCH_TYPE)
            udtProps.ExchangeNumber = FNN(rs!EXCH_NUMBER)
            udtProps.VOIDED = IIf(FNN(rs!EXCH_VOIDED), True, False)
            udtProps.ExchangeDate = FND(rs!EXCH_SaleDate)
            udtProps.ChangeGiven = FNN(rs!EXCH_ChangeGiven)
            udtProps.SupervisorID = FNN(rs!EXCH_SupervisorID)
            udtProps.SalesPersonName = FNS(rs!SM_SHortName)
            udtProps.Note = FNS(rs!EXCH_CUSTOMERNAME)
            udtProps.Status = FNS(rs!EXCH_Status)
            Set rs2 = New ADODB.Recordset
            rs2.Open "SELECT * FROM vCSL_Sale WHERE Exchange_GUID = '" & udtProps.ID & "'", oPC.COShort, adOpenKeyset, adLockOptimistic
            lngCount2 = 0
            Set oPBCSL = Nothing
            Set oPBCSL = New PropertyBag
            Set oPBCSLSet = Nothing
            Set oPBCSLSet = New PropertyBag
            If rs2.EOF Then
                oPBCSL.WriteProperty "Count", 0
                FetchCustomerRequests = oPBCSL.Contents
                oPB.WriteProperty "CSLS", oPBCSL.Contents
                Set oPBCSL = Nothing
            Else
              Do While Not rs2.EOF
                  udtProps2.CodeF = FNS(rs2!CodeF)
                  udtProps2.Title = FNS(rs2!P_Title)
                  udtProps2.Author = FNS(rs2!P_MainAuthor)
                  udtProps2.Qty = FNN(rs2!CSL_Qty)
                  udtProps2.Price = FNN(rs2!CSL_Price)
                  udtProps2.VATRate = FNDBL(rs2!CSL_VATRate)
                  udtProps2.DiscountRate = FNDBL(rs2!CSL_DiscountRate)
                  udtProps2.EXCHANGE_GUID = FNS(rs2!CSL_EXCHANGE_GUID)
                  LSet udtData2 = udtProps2
                  lngCount2 = lngCount2 + 1
                  oPBCSL.WriteProperty "Item" & CStr(lngCount2), udtData2.buffer
                  rs2.MoveNext
              Loop
              rs2.Close
              Set rs2 = Nothing
              oPBCSL.WriteProperty "COUNT", lngCount2
              oPB.WriteProperty "CSLS", oPBCSL.Contents
              Set oPBCSL = Nothing
              Set oPBCSLSet = Nothing
            End If
            
            Set rs3 = New ADODB.Recordset
            rs3.Open "SELECT * FROM vPayment WHERE Pay_Exchange_GUID = '" & udtProps.ID & "'", oPC.COShort, adOpenKeyset, adLockOptimistic
            lngCount3 = 0
            Set oPBPAY = Nothing
            Set oPBPAY = New PropertyBag
            Set oPBPAYSet = Nothing
            Set oPBPAYSet = New PropertyBag
            If rs3.EOF Then
                oPBPAY.WriteProperty "Count", 0
                FetchCustomerRequests = oPBPAY.Contents
            End If
            Do While Not rs3.EOF
                udtProps3.PaymentID = FNS(rs3!PAY_ID)
                udtProps3.EXCHANGE_GUID = FNS(rs3!PAY_Exchange_GUID)
                udtProps3.Amt = FNN(rs3!PAY_Amt)
                udtProps3.AmtTendered = FNN(rs3!PAY_Amt_Tendered)
                udtProps3.Type = FNS(rs3!PAY_PaymentType)
                udtProps3.Note = FNS(rs3!PAY_Note)
                LSet udtData3 = udtProps3
                lngCount3 = lngCount3 + 1
                oPBPAY.WriteProperty "Item" & CStr(lngCount3), udtData3.buffer
                rs3.MoveNext
            Loop
            rs3.Close
            Set rs3 = Nothing
            oPBPAY.WriteProperty "COUNT", lngCount3
            oPB.WriteProperty "PAYS", oPBPAY.Contents
            Set oPBPAY = Nothing
            Set oPBPAYSet = New PropertyBag
            
            rs.MoveNext
            LSet udtData = udtProps
            lngCount = lngCount + 1
            oPB.WriteProperty "State", udtData.buffer
            oPBSet.WriteProperty "Item" & CStr(lngCount), oPB.Contents
            Set oPB = Nothing
    Loop
    With oPBSet
          .WriteProperty "Count", lngCount
          FetchCustomerRequests = .Contents
    End With
    Set oPBSet = Nothing

EXIT_HANDLER:
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges_P.FetchCustomerRequests(pOSorALL)", Array(pOSorALL)
End Function

