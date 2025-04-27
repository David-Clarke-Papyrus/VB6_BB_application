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
Attribute VB_Exposed = False
Option Explicit

'Public Function Fetch(pOPID As String, pZID As String) As String
'On Error GoTo ErrHandler
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
'        rs.Open "SELECT * FROM tExchange LEFT JOIN tTP ON EXCH_TP_ID = TP_ID  WHERE EXCH_OPSESSIONID ='" & pOPID & "' ORDER BY EXCH_SALEDATE", oPC.COShort, adOpenForwardOnly
'    Else
'        rs.Open "SELECT * FROM tExchange LEFT JOIN tTP ON EXCH_TP_ID = TP_ID WHERE EXCH_ZSESSIONID ='" & pZID & "' ORDER BY EXCH_SALEDATE", oPC.COShort, adOpenForwardOnly
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
'            udtProps.TotalValueDiscount = FNN(rs!EXCHERR_GENERAL_Disc)
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
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "c_Exchanges_P.Fetch(pOPID As String, pZID As String)", Array(pOPID, pZID)
'End Function

Public Function Fetch(pOPID As String, pZID As String, pEID As String) As String
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

    Set oPBSet = New PropertyBag
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    bNoRS = False
    If pOPID > "" Then
        rs.Open "SELECT DISTINCT * FROM tExchange LEFT JOIN tTP ON EXCH_TP_ID = TP_ID  Left JOIN tStaffMember ON EXCH_SUPERVISORID = SM_ID WHERE EXCH_OPSESSIONID ='" & pOPID & "' ORDER BY EXCH_SALEDATE", oPC.COShort, adOpenForwardOnly
    ElseIf pZID > "" Then
        rs.Open "SELECT DISTINCT * FROM tExchange LEFT JOIN tTP ON EXCH_TP_ID = TP_ID   Left JOIN tStaffMember ON EXCH_SUPERVISORID = SM_ID WHERE EXCH_ZSESSIONID ='" & pZID & "' ORDER BY EXCH_SALEDATE", oPC.COShort, adOpenForwardOnly
    ElseIf pEID > "" Then
        rs.Open "SELECT DISTINCT * FROM tExchange LEFT JOIN tTP ON EXCH_TP_ID = TP_ID   Left JOIN tStaffMember ON EXCH_SUPERVISORID = SM_ID WHERE EXCH_ID ='" & pEID & "' ORDER BY EXCH_SALEDATE", oPC.COShort, adOpenForwardOnly
    Else
        bNoRS = True
    End If
    If bNoRS Then
        pblnNoRecsReturned = True
        oPBSet.WriteProperty "Count", 0
        Fetch = oPBSet.Contents
        Set oPBSet = Nothing
        Exit Function
    ElseIf rs.eof Then
        pblnNoRecsReturned = True
        oPBSet.WriteProperty "Count", 0
        Fetch = oPBSet.Contents
        Set oPBSet = Nothing
        Exit Function
    End If

    lngCount = 0
    Do While Not rs.eof
            Set oPB = New PropertyBag
            udtProps.ID = FNS(rs!EXCH_ID)
            udtProps.OPSID = FNS(rs!EXCH_OPSessionID)
            udtProps.ZID = FNS(rs!EXCH_ZSessionID)
            udtProps.TotalPayable = FNN(rs!EXCH_SALESVALUE)
            udtProps.TPID = FNN(rs!EXCH_TP_ID)
            udtProps.Note = FNS(rs!TP_Name) & " " & FNS(rs!TP_Initials)
            If FNN(rs!EXCH_TP_ID) = oPC.Configuration.CSCustomerID Then
              '  If FNS(rs!EXCH_CUSTOMERNAME) > "" Then
                '    udtProps = FNS(rs!EXCH_CUSTOMERNAME)
              '  End If
            End If
            udtProps.TotalDiscount = FNN(rs!EXCH_DISCOUNTVALUE)
            udtProps.TotalVAT = FNN(rs!EXCH_VATVALUE)
            udtProps.Type = FNS(rs!EXCH_TYPE)
         '   If FNN(rs!EXCH_NUMBER) = 2072 Then MsgBox "Here"
            udtProps.ExchangeNumber = FNN(rs!EXCH_NUMBER)
            udtProps.VOIDED = IIf(FNN(rs!EXCH_VOIDED), True, False)
            udtProps.ToVoid = FNN(rs!EXCH_VOIDS)
            udtProps.ExchangeDate = FND(rs!EXCH_SaleDate)
            udtProps.ChangeGiven = FNN(rs!EXCH_ChangeGiven)
            udtProps.SupervisorID = FNN(rs!EXCH_SupervisorID)
            udtProps.SalesPersonName = FNS(rs!SM_SHortName)

            Set rs2 = New ADODB.Recordset
            If Trim(udtProps.Type) = "APP" Then
                strSQL = "SELECT * FROM vCSL_APP WHERE Exchange_GUID = '" & udtProps.ID & "'"
            ElseIf Trim(udtProps.Type) = "AR" Then
                strSQL = "SELECT * FROM vCSL_APPR  WHERE Exchange_GUID = '" & FNS(udtProps.ID) _
                & "' UNION ALL SELECT * FROM vCSL_Sale WHERE Exchange_GUID = '" & FNS(udtProps.ID) _
                & "' UNION ALL SELECT * FROM vCSL_INV WHERE Exchange_GUID = '" & FNS(udtProps.ID) & "'"
            ElseIf Trim(udtProps.Type) = "A" Then
                strSQL = "SELECT * FROM vCSL_INV WHERE Exchange_GUID = '" & Trim(udtProps.ID) & "'"
            Else
                strSQL = "SELECT * FROM vCSL_Sale WHERE Exchange_GUID = '" & Trim(udtProps.ID) & "'"
            End If
            rs2.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            lngCount2 = 0
            Set oPBCSL = Nothing
            Set oPBCSL = New PropertyBag
            Set oPBCSLSet = Nothing
            Set oPBCSLSet = New PropertyBag
            If rs2.eof Then
                oPBCSL.WriteProperty "Count", 0
                Fetch = oPBCSL.Contents
                oPB.WriteProperty "CSLS", oPBCSL.Contents
                Set oPBCSL = Nothing
            Else
              Do While Not rs2.eof
                  udtProps2.CodeF = FNS(rs2!CodeF)
                  udtProps2.Title = FNS(rs2!P_Title)
                  udtProps2.Author = FNS(rs2!P_MainAuthor)
                  udtProps2.Qty = FNN(rs2!Qty)
                  udtProps2.QtyF = FNS(rs2!QtyF)
                  udtProps2.Price = FNN(rs2!Price)
                  udtProps2.VATRate = FNDBL(rs2!VATRate)
                  udtProps2.DiscountRate = FNDBL(rs2!DiscountRate)
                  udtProps2.EXCHANGE_GUID = FNS(rs2!EXCHANGE_GUID)
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
            If rs3.eof Then
                oPBPAY.WriteProperty "Count", 0
                Fetch = oPBPAY.Contents
            End If
            Do While Not rs3.eof
                udtProps3.PaymentID = FNS(rs3!PAY_ID)
                udtProps3.EXCHANGE_GUID = FNS(rs3!PAY_Exchange_GUID)
                udtProps3.Amt = FNN(rs3!PAY_Amt)
                udtProps3.AmtTendered = FNN(rs3!PAY_Amt_Tendered)
                udtProps3.Type = FNS(rs3!PAY_PaymentType)
                udtProps3.Note = FNS(rs3!PAY_Note)
                udtProps3.Reference = FNS(rs3!PAY_Reference)
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

EXIT_Handler:
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges_P.Fetch(pOPID,pZID)", Array(pOPID, pZID)
End Function


Public Function FetchCustomequests(pOSorALL As String, dSince As Date) As String
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

    Set oPBSet = New PropertyBag
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    bNoRS = False
    If pOSorALL = "OS" Then
        rs.Open "SELECT * FROM vOrderRequests WHERE EXCH_SALEDATE > '" & ReverseDate(dSince) & "' AND ActionedStatus IS NULL ORDER BY EXCH_SALEDATE DESC", oPC.COShort, adOpenForwardOnly
    Else
        rs.Open "SELECT * FROM vOrderRequests WHERE EXCH_SALEDATE > '" & ReverseDate(dSince) & "' ORDER BY EXCH_SALEDATE DESC", oPC.COShort, adOpenForwardOnly
    End If
    If rs.eof Then
        pblnNoRecsReturned = True
        oPBSet.WriteProperty "Count", 0
        FetchCustomequests = oPBSet.Contents
        Set oPBSet = Nothing
        Exit Function
    End If

    lngCount = 0
    Do While Not rs.eof
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
            udtProps.OR_ActionedDate = FND(rs!ActionedStatus)
            Set rs2 = New ADODB.Recordset
            rs2.Open "SELECT * FROM vCSL_Sale WHERE Exchange_GUID = '" & udtProps.ID & "'", oPC.COShort, adOpenKeyset, adLockOptimistic
            lngCount2 = 0
            Set oPBCSL = Nothing
            Set oPBCSL = New PropertyBag
            Set oPBCSLSet = Nothing
            Set oPBCSLSet = New PropertyBag
            If rs2.eof Then
                oPBCSL.WriteProperty "Count", 0
                FetchCustomequests = oPBCSL.Contents
                oPB.WriteProperty "CSLS", oPBCSL.Contents
                Set oPBCSL = Nothing
            Else
              Do While Not rs2.eof
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
            If rs3.eof Then
                oPBPAY.WriteProperty "Count", 0
                FetchCustomequests = oPBPAY.Contents
            End If
            Do While Not rs3.eof
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
          FetchCustomequests = .Contents
    End With
    Set oPBSet = Nothing

EXIT_Handler:
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_Exchanges_P.FetchCustomequests(pOSorALL)", Array(pOSorALL)
End Function

