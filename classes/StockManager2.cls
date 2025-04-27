VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_StockManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Event eMax(Max As Long)
Event eProgress(Progress As Long)


Public Function IssueInvoice(pInvoice As a_Invoice, pWasProforma As Boolean, pToProForma As Boolean)
Dim strSQL As String
Dim rsIL As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsPI As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsSales As ADODB.Recordset
Dim rsWants As ADODB.Recordset
Dim rsCOL As ADODB.Recordset
On Error GoTo ERRH
'If invoice is proforma
    'If invoice line is for a product instance
        'On_Reserve is set to true
    'Else
        'reserve levels are increased but no stock levels are altered
     'end if
'else
    'If invoice line is for a product instance
        'if ir was on reserve then
            'On-Reserve is set off
        'end if
        'Date of sale is set
    'Else
        'if the invoice WAS Proforma
            'Reserve values are reduced
        'end if
        'stock levels are reduced
    'end if
'end if
    
    Set rsIL = New ADODB.Recordset
    strSQL = "SELECT IL_Price,IL_P_ID,IL_COL_ID,IL_PI_ID,IL_QTY,IL_VATValue,IL_SalesValue FROM tILine WHERE IL_TR_ID = " & pInvoice.InvoiceID
    rsIL.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    
    Do While Not rsIL.EOF
    
        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_LastDateSold,P_LastQtySold,P_LastPriceSold,P_QtyOnHand,P_QtyCopiesOnHand,p_QtyReserved,P_QtyExpectedBack,P_QTYONBACKORDER FROM tProduct Where P_ID = '" & rsIL.Fields("IL_P_ID") & "'"
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        Set rsPI = New ADODB.Recordset
        strSQL = "SELECT PI_OnReserve,PI_ExpectedBack,PI_SoldDate,PI_SoldTo_ID FROM tPI Where PI_ID = " & rsIL.Fields("IL_PI_ID")
        rsPI.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_P_ID,STP_ST_ID,STP_QtyOnBackorder,STP_QTYONORDER,STP_QTYLASTSTOCKTAKE,STP_QTYRESERVED,STP_QTYONHAND,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED FROM tStoreP Where STP_P_ID = '" & rsIL.Fields("IL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        If rsSP.EOF Then  'This might be a new product
            rsSP.AddNew
            rsSP!STP_P_ID = rsIL!IL_P_ID
            rsSP!STP_ST_ID = oPC.Configuration.DefaultStoreID
            rsSP!STP_QtyOnBackorder = 0
            rsSP!STP_QtyOnOrder = 0
            rsSP!STP_QtyLastStockTake = 0
            rsSP!STP_QTYRESERVED = 0
            rsSP!STP_QTYCOPIESONHAND = 0
            rsSP!STP_QTYONHAND = 0
            rsSP!STP_QTYCOPIESRESERVED = 0
            rsSP.Update
            rsSP.Bookmark = rsSP.Bookmark
        End If
        strSQL = "SELECT tCOL.* FROM tCOL JOIN tCO ON COL_TR_ID = CO_ID WHERE CO_ORDERTYPE = 2 AND COL_P_ID = '" & rsIL!IL_P_ID & "' AND " & rsIL!IL_PI_ID & " > 0"
        Set rsWants = New ADODB.Recordset
        rsWants.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        If Not rsWants.EOF Then
            rsWants!COL_QtyDispatched = 1
            rsWants!COL_Fulfilled = "FUL"
            rsWants!COL_ActionTaken = 9
            rsWants.Update
        End If
        strSQL = "SELECT tCOL.* FROM tCOL JOIN tCO ON COL_TR_ID = CO_ID WHERE CO_ORDERTYPE = 1 AND COL_ID = " & FNN(rsIL!IL_COL_ID)
        Set rsCOL = New ADODB.Recordset
        rsCOL.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        If Not rsCOL.EOF Then
            rsCOL!COL_QtyDispatched = FNN(rsIL!IL_QTY)
            If FNN(rsCOL!COL_QtyDispatched) >= FNN(rsCOL!COL_Qty) Then
                rsCOL!COL_Fulfilled = "FUL"
                rsCOL!COL_ActionTaken = 9
            End If
            rsCOL.Update
            rsSP.Fields("STP_QTYONBACKORDER") = NonNegative_Long(FNN(rsSP.Fields("STP_QTYONBACKORDER")) - FNN(rsIL.Fields("IL_QTY")))
            rsP.Fields("P_QTYONBACKORDER") = NonNegative_Long(FNN(rsP.Fields("P_QTYONBACKORDER")) - FNN(rsIL.Fields("IL_QTY")))
        End If
        If pToProForma Then
            If Not pWasProforma Then  ' mark it as proforma
                If rsIL.Fields("IL_PI_ID") > 0 Then
                    If Not IsNull(rsPI.Fields("PI_ONRESERVE")) Then
                        rsPI.Fields("PI_ONRESERVE") = 1
                    End If
                    rsSP.Fields("STP_QTYCOPIESRESERVED") = FNN(rsSP.Fields("STP_QTYCOPIESRESERVED")) + 1
                Else
                    rsSP.Fields("STP_QTYRESERVED") = FNN(rsSP.Fields("STP_QTYRESERVED")) + FNN(rsIL.Fields("IL_QTY"))
                End If
            End If
        Else
            If FNN(rsIL.Fields("IL_PI_ID")) > 0 Then
                rsPI.Fields("PI_ONRESERVE") = 0
                rsPI.Fields("PI_SOLDDATE") = ReverseDate(pInvoice.DocDate)
                rsPI.Fields("PI_SoldTo_ID") = pInvoice.Customer.ID
                If pWasProforma Then
                    rsSP.Fields("STP_QTYCOPIESRESERVED") = FNN(rsSP.Fields("STP_QTYCOPIESRESERVED")) - 1
                    If rsSP.Fields("STP_QTYCOPIESRESERVED") < 0 Then rsSP.Fields("STP_QTYCOPIESRESERVED") = 0
                End If
                rsSP.Fields("STP_QTYCOPIESONHAND") = FNN(rsSP.Fields("STP_QTYCOPIESONHAND")) - 1
                Set rsSales = New ADODB.Recordset
                rsSales.Open "SELECT * FROM tSALES_CYr WHERE SCY_P_ID = '" & rsIL.Fields("IL_P_ID") & "' AND SCY_Dte = '" & ReverseDate(pInvoice.DocDate) & "' AND SCY_TP_ID = " & pInvoice.Customer.ID, oPC.CO, adOpenDynamic, adLockOptimistic
                If Not rsSales.EOF Then
                    rsSales.Fields("SCY_QTYCopies") = FNN(rsSales.Fields("SCY_QTYCopies")) + FNN(rsIL.Fields("IL_QTY"))
                    rsSales.Fields("SCY_QTY") = FNN(rsSales.Fields("SCY_QTY")) + FNN(rsIL.Fields("IL_QTY"))
                    rsSales.Fields("SCY_Value") = FNN(rsSales.Fields("SCY_Value")) + FNN(rsIL.Fields("IL_SalesValue"))
                    rsSales.Fields("SCY_VAT") = FNN(rsSales.Fields("SCY_VAT")) + FNN(rsIL.Fields("IL_VATVALUE"))
                    rsSales.Update
                Else  'There is no sale for this customer/product/day
                    rsSales.AddNew
                    rsSales.Fields("SCY_P_ID") = rsIL.Fields("IL_P_ID")
                    rsSales.Fields("SCY_Dte") = pInvoice.DocDate
                    rsSales.Fields("SCY_TP_ID") = pInvoice.Customer.ID
                    rsSales.Fields("SCY_QTYCopies") = FNN(rsIL.Fields("IL_QTY"))
                    rsSales.Fields("SCY_QTY") = FNN(rsIL.Fields("IL_QTY"))
                    rsSales.Fields("SCY_Value") = FNN(rsIL.Fields("IL_SalesValue"))
                    rsSales.Fields("SCY_VAT") = FNN(rsIL.Fields("IL_VATVALUE"))
                    rsSales.Update
                    
                End If
            Else
                If pWasProforma Then
                    rsSP.Fields("STP_QTYRESERVED") = FNN(rsSP.Fields("STP_QTYRESERVED")) - FNN(rsIL.Fields("IL_QTY"))
                End If
                '''
                rsSP.Fields("STP_QTYONHAND") = FNN(rsSP.Fields("STP_QTYONHAND")) - FNN(rsIL.Fields("IL_QTY"))
                rsP.Fields("P_QtyONHAND") = FNN(rsP.Fields("P_QTYONHAND")) - FNN(rsIL.Fields("IL_QTY"))
                rsP.Fields("P_LastDateSold") = pInvoice.DocDate
                rsP.Fields("P_LASTQTYSOLD") = FNN(rsIL.Fields("IL_QTY"))
                rsP.Fields("P_LASTPRICESOLD") = FNN(rsIL.Fields("IL_PRICE"))
                Set rsSales = New ADODB.Recordset
                rsSales.Open "SELECT * FROM tSALES_CYr WHERE SCY_P_ID = '" & rsIL.Fields("IL_P_ID") & "' AND SCY_Dte = '" & ReverseDate(pInvoice.DocDate) & "' AND SCY_TP_ID = " & pInvoice.Customer.ID, oPC.CO, adOpenDynamic, adLockOptimistic
                If Not rsSales.EOF Then
                    rsSales.Fields("SCY_QTY") = FNN(rsSales.Fields("SCY_QTY")) + FNN(rsIL.Fields("IL_QTY"))
                    rsSales.Fields("SCY_Value") = FNN(rsSales.Fields("SCY_Value")) + FNN(rsIL.Fields("IL_SalesValue"))
                    rsSales.Fields("SCY_VAT") = FNN(rsSales.Fields("SCY_VAT")) + FNN(rsIL.Fields("IL_VATVALUE"))
                    rsSales.Update
                Else  'There is no sale for this customer/product/day
                    rsSales.AddNew
                    rsSales.Fields("SCY_P_ID") = rsIL.Fields("IL_P_ID")
                    rsSales.Fields("SCY_Dte") = pInvoice.DocDate
                    rsSales.Fields("SCY_TP_ID") = pInvoice.Customer.ID
                    rsSales.Fields("SCY_QTY") = FNN(rsIL.Fields("IL_QTY"))
                    rsSales.Fields("SCY_Value") = FNN(rsIL.Fields("IL_SalesValue"))
                    rsSales.Fields("SCY_VAT") = FNN(rsIL.Fields("IL_VATVALUE"))
                    rsSales.Update
                End If
            End If
        End If
        If Not rsSP.EOF Then rsSP.Update
        If Not rsP.EOF Then rsP.Update
        If Not rsPI.EOF Then rsPI.Update
        rsIL.MoveNext
    Loop
    rsIL.Close
    Set rsIL = Nothing
    rsSP.Close
    rsP.Close
    rsPI.Close
    Set rsSP = Nothing
    Set rsP = Nothing
    Set rsPI = Nothing
    Exit Function
ERRH:
    MsgBox "z_Stockmanager:IssueInvoice: " & Error
    rsIL.Close
    oError.SetError Err, "z_Stockmanager:IssueInvoice: " & Error, Now(), "", "", ""
    On Error Resume Next
    Set rsIL = Nothing
    rsSP.Close
    rsP.Close
    rsPI.Close
    Set rsSP = Nothing
    Set rsP = Nothing
    Set rsPI = Nothing
    Exit Function
    Resume
End Function
Public Function CancelInvoice(pInvoice As a_Invoice) As Boolean
Dim strSQL As String
Dim rsIL As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsPI As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsSales As ADODB.Recordset

'If invoice is proforma
    'If invoice line is for a product instance
        'On_Reserve is set to true
    'Else
        'reserve levels are increased but no stock levels are altered
     'end if
'else
    'If invoice line is for a product instance
        'if ir was on reserve then
            'On-Reserve is set off
        'end if
        'Date of sale is set
    'Else
        'if the invoice WAS Proforma
            'Reserve values are reduced
        'end if
        'stock levels are reduced
    'end if
'end if
'    strSQL = " SELECT IL_QTY,PI_ONRESERVE,PI_ExpectedBack,STP_QTYRESERVED,I_PROFORMA,PI_SOLDDATE,STP_QTYONHAND,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED,IL_COPY_ID " _
             & " FROM tILine JOIN TInvoice On I_ID = IL_TR_ID JOIN tProduct on P_ID = IL_P_ID  Join tStoreP on STP_P_ID = P_ID left outer Join tPI On IL_COPY_ID = PI_ID " _
             & " WHERE IL_TR_ID = " & pInvoice.InvoiceID & " and STP_ST_ID = " & oPC.Configuration.DefaultStoreID
    Set rsIL = New ADODB.Recordset
    strSQL = "SELECT IL_P_ID,IL_SalesValue,IL_VATValue,IL_PI_ID,IL_QTY FROM tILine WHERE IL_TR_ID = " & pInvoice.InvoiceID
    rsIL.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    
    Do While Not rsIL.EOF
    
        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QtyOnHand,P_QtyCopiesOnHand,p_QtyReserved,P_QtyExpectedBack FROM tProduct Where P_ID = '" & rsIL.Fields("IL_P_ID") & "'"
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        Set rsPI = New ADODB.Recordset
        strSQL = "SELECT PI_OnReserve,PI_ExpectedBack,PI_SoldDate FROM tPI Where PI_ID = " & FNN(rsIL.Fields("IL_PI_ID"))
        rsPI.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QTYRESERVED,STP_QTYONHAND,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED FROM tStoreP Where STP_P_ID = '" & rsIL.Fields("IL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        CancelInvoice = False
        If pInvoice.StatusF = "PROFORMA" Then
            If rsIL.Fields("IL_PI_ID") > 0 Then
                If Not IsNull(rsPI.Fields("PI_ONRESERVE")) Then
                    rsPI.Fields("PI_ONRESERVE") = 0
                End If
                rsSP.Fields("STP_QTYCOPIESRESERVED") = FNN(rsSP.Fields("STP_QTYCOPIESRESERVED")) - 1
            Else
                rsSP.Fields("STP_QTYRESERVED") = NonNegative_Lng(FNN(rsSP.Fields("STP_QTYRESERVED")) - FNN(rsIL.Fields("IL_QTY")))
            End If
            CancelInvoice = True
'        End If
        Else
            If FNN(rsIL.Fields("IL_PI_ID")) > 0 Then
                rsPI.Fields("PI_ONRESERVE") = 0
                rsPI.Fields("PI_SOLDDATE") = CDate(0)
                Set rsSales = New ADODB.Recordset
                rsSales.Open "SELECT * FROM tSALES_CYr WHERE SCY_P_ID = '" & rsIL.Fields("IL_P_ID") & "' AND SCY_Dte = '" & ReverseDate(pInvoice.DocDate) & "' AND SCY_TP_ID = " & pInvoice.Customer.ID, oPC.CO, adOpenDynamic, adLockOptimistic
                If Not rsSales.EOF Then
                    rsSales.Fields("SCY_QTYCopies") = FNN(rsSales.Fields("SCY_QTYCopies")) - FNN(rsIL.Fields("IL_QTY"))
                    rsSales.Fields("SCY_QTY") = FNN(rsSales.Fields("SCY_QTY")) - FNN(rsIL.Fields("IL_QTY"))
                    rsSales.Fields("SCY_Value") = FNN(rsSales.Fields("SCY_Value")) - FNN(rsIL.Fields("IL_SalesValue"))
                    rsSales.Fields("SCY_VAT") = FNN(rsSales.Fields("SCY_VAT")) - FNN(rsIL.Fields("IL_VATVALUE"))
                    rsSales.Update
                End If
            Else
                rsSP.Fields("STP_QTYONHAND") = FNN(rsSP.Fields("STP_QTYONHAND")) + FNN(rsIL.Fields("IL_QTY"))
                Set rsSales = New ADODB.Recordset
                rsSales.Open "SELECT * FROM tSALES_CYr WHERE SCY_P_ID = '" & rsIL.Fields("IL_P_ID") & "' AND SCY_Dte = '" & ReverseDate(pInvoice.DocDate) & "' AND SCY_TP_ID = " & pInvoice.Customer.ID, oPC.CO, adOpenDynamic, adLockOptimistic
                If Not rsSales.EOF Then
                    rsSales.Fields("SCY_QTY") = FNN(rsSales.Fields("SCY_QTY")) - FNN(rsIL.Fields("IL_QTY"))
                    rsSales.Fields("SCY_Value") = FNN(rsSales.Fields("SCY_Value")) - FNN(rsIL.Fields("IL_SalesValue"))
                    rsSales.Fields("SCY_VAT") = FNN(rsSales.Fields("SCY_VAT")) - FNN(rsIL.Fields("IL_VATVALUE"))
                    rsSales.Update
                End If
            End If
        End If
        If Not rsSP.EOF Then rsSP.Update
        If Not rsP.EOF Then rsP.Update
        If Not rsPI.EOF Then rsPI.Update
        rsIL.MoveNext
    Loop
    rsIL.Close
    Set rsIL = Nothing
    rsSP.Close
    rsP.Close
    rsPI.Close
    Set rsSP = Nothing
    Set rsP = Nothing
    Set rsPI = Nothing
    oPC.CO.execute "UPDATE tTR SET TR_STATUS = 6 WHERE TR_ID = " & pInvoice.InvoiceID
    Exit Function
ERRH:
    MsgBox "z_Stockmanager:CancelInvoice: " & Error
    rsIL.Close
    oError.SetError Err, "z_Stockmanager:CancelInvoice: " & Error, Now(), "", "", ""
    On Error Resume Next
    Set rsIL = Nothing
    rsSP.Close
    rsP.Close
    rsPI.Close
    Set rsSP = Nothing
    Set rsP = Nothing
    Set rsPI = Nothing

End Function
Public Function IssueCreditNote(pCN As a_CN, pWasProforma As Boolean, pToProForma As Boolean)
On Error GoTo ERRH
Dim strSQL As String
Dim rsCNL As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsPI As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsSales As ADODB.Recordset

'If CN is proforma
    'If CN line is for a product instance
        'On_Reserve is set to true
    'Else
        'reserve levels are increased but no stock levels are altered
     'end if
'else
    'If CN line is for a product instance
        'if ir was on reserve then
            'On-Reserve is set off
        'end if
        'Date of sale is set
    'Else
        'if the CN WAS Proforma
            'Reserve values are reduced
        'end if
        'stock levels are reduced
    'end if
'end if
    Set rsCNL = New ADODB.Recordset
    strSQL = "SELECT CNL_P_ID,CNL_PI_ID,CNL_QTY,CNL_SalesValue,CNL_VATVALUE FROM tCNL WHERE CNL_TR_ID = " & pCN.TRID
    rsCNL.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    
    Do While Not rsCNL.EOF
    
        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QtyOnHand,P_QtyCopiesOnHand,p_QtyReserved,P_QtyExpectedBack,P_QtyTotalSold FROM tProduct Where P_ID = '" & rsCNL.Fields("CNL_P_ID") & "'"
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        Set rsPI = New ADODB.Recordset
        strSQL = "SELECT PI_OnReserve,PI_FirstSoldDate,PI_ExpectedBack,PI_SoldDate,PI_SoldTo_ID FROM tPI Where PI_ID = " & rsCNL.Fields("CNL_PI_ID")
        rsPI.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QTYRESERVED,STP_QTYONHAND,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED FROM tStoreP Where STP_P_ID = '" & rsCNL.Fields("CNL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        If pToProForma Then
            If Not pWasProforma Then  ' mark it as proforma
                If rsCNL.Fields("CNL_PI_ID") > 0 Then
                    If Not IsNull(rsPI.Fields("PI_EXPECTEDBACK,")) Then
                        rsPI.Fields("PI_EXPECTEDBACK") = 1
                    End If
                    If rsSP.Fields("STP_QTYCOPIESRESERVED") > 0 Then
                        rsSP.Fields("STP_QTYCOPIESRESERVED") = rsSP.Fields("STP_QTYCOPIESRESERVED") + 1
                    End If
                Else
                    rsSP.Fields("STP_QTYRESERVED") = rsSP.Fields("STP_QTYRESERVED") + rsCNL.Fields("CNL_QTY")
                    If rsSP.Fields("STP_QTYRESERVED") < 0 Then rsSP.Fields("STP_QTYRESERVED") = 0
                End If
            End If
        Else
            If FNN(rsCNL.Fields("CNL_PI_ID")) > 0 Then
                rsPI.Fields("PI_ONRESERVE") = 0
                rsPI.Fields("PI_FirstSoldDate") = FND(rsPI.Fields("PI_SOLDDATE"))
                rsPI.Fields("PI_SOLDDATE") = CDate(0)
                rsPI.Fields("PI_SoldTO_ID") = 0
                rsP.Fields("P_QtyOnHand") = FNN(rsSP.Fields("STP_QTYONHAND")) + 1
                rsSP.Fields("STP_QTYCOPIESONHAND") = FNN(rsSP.Fields("STP_QTYCOPIESONHAND")) + 1
            Else
                rsSP.Fields("STP_QTYONHAND") = FNN(rsSP.Fields("STP_QTYONHAND")) + rsCNL.Fields("CNL_QTY")
                rsP.Fields("P_QtyOnHand") = FNN(rsSP.Fields("STP_QTYONHAND")) + FNN(rsCNL.Fields("CNL_QTY"))
                rsP.Fields("P_QtyTotalSold") = NonNegative_Lng(FNN(rsSP.Fields("P_QtyTotalSold")) - FNN(rsCNL.Fields("CNL_QTY")))
            End If
            Set rsSales = New ADODB.Recordset
            rsSales.Open "SELECT * FROM tSALES_CYr WHERE SCY_P_ID = '" & rsCNL.Fields("CNL_P_ID") & "' AND SCY_Dte = '" & ReverseDate(pCN.DocDate) & "' AND SCY_TP_ID = " & pCN.Customer.ID, oPC.CO, adOpenDynamic, adLockOptimistic
            If Not rsSales.EOF Then
                rsSales.Fields("SCY_QTYCopies") = FNN(rsSales.Fields("SCY_QTYCopies")) - FNN(rsCNL.Fields("CNL_QTY"))
                rsSales.Fields("SCY_QTY") = FNN(rsSales.Fields("SCY_QTY")) - FNN(rsCNL.Fields("CNL_QTY"))
                rsSales.Fields("SCY_Value") = FNN(rsSales.Fields("SCY_Value")) - FNN(rsCNL.Fields("CNL_SalesValue"))
                rsSales.Fields("SCY_VAT") = FNN(rsSales.Fields("SCY_VAT")) - FNN(rsCNL.Fields("CNL_VATVALUE"))
                rsSales.Update
            End If
        End If
        If Not rsSP.EOF Then rsSP.Update
        If Not rsP.EOF Then rsP.Update
        If Not rsPI.EOF Then rsPI.Update
        rsCNL.MoveNext
    Loop
    rsCNL.Close
    Set rsCNL = Nothing
    rsSP.Close
    rsP.Close
    rsPI.Close
    Set rsSP = Nothing
    Set rsP = Nothing
    Set rsPI = Nothing
    Exit Function
ERRH:
    MsgBox "z_Stockmanager:IssueCreditnote: " & Error
    oError.SetError Err, "z_Stockmanager:IssueCreditnote: " & Error, Now(), "", "", ""
    On Error Resume Next
    rsCNL.Close
    Set rsCNL = Nothing
    rsSP.Close
    rsP.Close
    rsPI.Close
    Set rsSP = Nothing
    Set rsP = Nothing
    Set rsPI = Nothing
    Exit Function
    Resume
End Function

Public Function CancelCreditNote(pCN As a_CN)
On Error GoTo ERRH
Dim strSQL As String
Dim rsCNL As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsPI As ADODB.Recordset
Dim rsSP As ADODB.Recordset

    Set rsCNL = New ADODB.Recordset
    strSQL = "SELECT CNL_P_ID,CNL_PI_ID,CNL_QTY FROM tCNL WHERE CNL_TR_ID = " & pCN.TRID
    rsCNL.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    
    Do While Not rsCNL.EOF
    
        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QtyOnHand,P_QtyCopiesOnHand,p_QtyReserved,P_QtyExpectedBack FROM tProduct Where P_ID = '" & rsCNL.Fields("CNL_P_ID") & "'"
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        Set rsPI = New ADODB.Recordset
        strSQL = "SELECT PI_OnReserve,PI_FirstSoldDate,PI_ExpectedBack,PI_SoldDate FROM tPI Where PI_ID = " & rsCNL.Fields("CNL_PI_ID")
        rsPI.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QTYRESERVED,STP_QTYONHAND,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED FROM tStoreP Where STP_P_ID = '" & rsCNL.Fields("CNL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        If pCN.StatusF = "PROFORMA" Then 'stock has not been issued
            If rsCNL.Fields("CNL_PI_ID") > 0 Then
                If Not IsNull(rsPI.Fields("PI_EXPECTEDBACK,")) Then
                    rsPI.Fields("PI_EXPECTEDBACK") = 0
                End If
            Else
                rsSP.Fields("STP_QTYRESERVED") = rsSP.Fields("STP_QTYRESERVED") - rsCNL.Fields("COL_QTY")
                If rsSP.Fields("STP_QTYRESERVED") < 0 Then rsSP.Fields("STP_QTYRESERVED") = 0
            End If
        Else
            If FNN(rsCNL.Fields("CNL_PI_ID")) > 0 Then
                rsPI.Fields("PI_ONRESERVE") = 0
                rsPI.Fields("PI_SOLDDATE") = FND(rsPI.Fields("PI_FirstSOLDDATE"))
                rsSP.Fields("STP_QTYCOPIESONHAND") = rsSP.Fields("STP_QTYCOPIESONHAND") - 1
            Else
                rsSP.Fields("STP_QTYONHAND") = rsSP.Fields("STP_QTYONHAND") - rsCNL.Fields("CNL_QTY")
                rsP.Fields("P_QtyOnHand") = FNN(rsSP.Fields("STP_QTYONHAND")) - FNN(rsCNL.Fields("CNL_QTY"))
                rsP.Fields("P_QtySold") = FNN(rsSP.Fields("P_QtySold")) - FNN(rsCNL.Fields("CNL_QTY"))
            End If
        End If
        If Not rsSP.EOF Then rsSP.Update
        If Not rsP.EOF Then rsP.Update
        If Not rsPI.EOF Then rsPI.Update
        rsCNL.MoveNext
    Loop
    rsCNL.Close
    Set rsCNL = Nothing
    rsSP.Close
    rsP.Close
    rsPI.Close
    Set rsSP = Nothing
    Set rsP = Nothing
    Set rsPI = Nothing
    Exit Function
ERRH:
    MsgBox "z_Stockmanager:CancelCreditNote: " & Error
    oError.SetError Err, "z_Stockmanager:CancelCreditNote: " & Error, Now(), "", "", ""
    On Error Resume Next
    rsCNL.Close
    Set rsCNL = Nothing
    rsSP.Close
    rsP.Close
    rsPI.Close
    Set rsSP = Nothing
    Set rsP = Nothing
    Set rsPI = Nothing

End Function

Public Function IssueCO(pCO As a_CO)
On Error GoTo ERRH
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsP As ADODB.Recordset
'UPdate     STP_QtyOnBackorder

    Set rs = New ADODB.Recordset
    Set rsP = New ADODB.Recordset
    Set rsSP = New ADODB.Recordset
    strSQL = "SELECT COL_P_ID,COL_QTY FROM tCOL WHERE COL_TR_ID = " & pCO.TRID
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    
    Do While Not rs.EOF
    
        strSQL = "SELECT STP_QTYRESERVED,STP_QTYONHAND,STP_QTYONBACKORDER FROM tStoreP Where STP_P_ID = '" & rs.Fields("COL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        If Not rsSP.EOF Then
            rsSP.Fields("STP_QTYONBACKORDER") = FNN(rsSP.Fields("STP_QTYONBACKORDER")) + FNN(rs.Fields("COL_QTY"))
            rsSP.Update
        End If
        rsSP.Close
        strSQL = "SELECT P_QTYONBACKORDER FROM tPRODUCT WHERE P_ID = '" & rs.Fields("COL_P_ID") & "'"
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        rsP!P_QTYONBACKORDER = FNN(rsP!P_QTYONBACKORDER) + FNN(rs.Fields("COL_QTY"))
        If Not rsP.EOF Then rsP.Update
        rsP.Close
        rs.MoveNext
    Loop
    rs.Close
    Set rsP = Nothing
    Set rs = Nothing
    Set rsSP = Nothing
    Exit Function
ERRH:
    oError.SetError Err, "z_Stockmanager:IssueCO: " & Error, Now(), "", "", ""
    MsgBox "z_Stockmanager:IssueCO: " & Error
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    rsSP.Close
    Set rsSP = Nothing
    Exit Function
    Resume
End Function
Public Function CancelCO(pCO As a_CO)
On Error GoTo ERRH
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsSP As ADODB.Recordset
    If Not pCO.CanCancel Then Exit Function
    Set rs = New ADODB.Recordset
    strSQL = "SELECT COL_P_ID,COL_QTY FROM tCOL WHERE COL_TR_ID = " & pCO.TRID
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    
    oPC.CO.BeginTrans
    Do While Not rs.EOF
        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnBackorder,STP_QTYRESERVED,STP_QTYONHAND,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED FROM tStoreP Where STP_P_ID = '" & rs.Fields("COL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsSP.Fields("STP_QtyOnBackorder") = NonNegative_Lng(FNN(rsSP.Fields("STP_QtyOnBackorder")) - FNN(rs.Fields("COL_QTY")))
        If Not rsSP.EOF Then rsSP.Update
        
        strSQL = "SELECT P_QTYONBACKORDER FROM tPRODUCT WHERE P_ID = '" & rs.Fields("COL_P_ID") & "'"
        Set rsP = New ADODB.Recordset
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        rsP!P_QTYONBACKORDER = NonNegative_Lng(FNN(rsP!P_QTYONBACKORDER) - FNN(rs.Fields("COL_QTY")))
        If Not rsP.EOF Then rsP.Update
        rsP.Close
        Set rsP = Nothing
        
        
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    rsSP.Close
    Set rsSP = Nothing
    oPC.CO.execute "UPDATE tTR SET TR_STATUS = 6 WHERE TR_ID = " & pCO.TRID
    oPC.CO.CommitTrans
    Exit Function
ERRH:
    oPC.CO.RollbackTrans
    MsgBox Error
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    rsSP.Close
    Set rsSP = Nothing
    Exit Function
    Resume
End Function
Public Function MarkCOLsActionedForProduct(pPID As String)
On Error GoTo ERRH

    oPC.CO.execute "UPDATE tCOL SET COL_ACTIONTAKEN = 3 FROM tCOL JOIN tTR ON COL_TR_ID = TR_ID  WHERE COL_P_ID = '" & pPID & "' AND (COL_ACTIONTAKEN = 0 Or COL_ACTIONTAKEN = 2) AND TR_STATUS = 3"
Exit Function
ERRH:
    MsgBox Error
    
End Function
Public Function CancelCOL(pCOLID As Long) ', pPID As String, pQty As Long)
Dim strSQL As String
Dim rsP As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim lngQty As Long
Dim strPID As String

On Error GoTo ERRH
    oPC.CO.BeginTrans
    Set rs = New ADODB.Recordset
    rs.Open "Select COL_QTY,COL_P_ID FROM tCOL WHERE COL_ID = " & pCOLID, oPC.CO, adOpenKeyset
    strPID = rs.Fields("COL_P_ID")
    lngQty = rs.Fields("COL_Qty")
    rs.Close
    Set rs = Nothing
    
    oPC.CO.execute "UPDATE tCOL SET COL_Fulfilled = 'CAN' WHERE COL_ID = " & pCOLID
    
    Set rsP = New ADODB.Recordset
    strSQL = "SELECT P_QTYONBACKORDER FROM tPRODUCT WHERE P_ID = '" & strPID & "'"
    rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    rsP!P_QTYONBACKORDER = FNN(rsP!P_QTYONBACKORDER) - lngQty
    If Not rsP.EOF Then rsP.Update
    rsP.Close
    Set rsP = Nothing
    
    Set rsSTP = New ADODB.Recordset
    strSQL = "SELECT STP_QTYONBACKORDER FROM tStoreP WHERE STP_P_ID = '" & strPID & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
    rsSTP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    rsSTP!STP_QtyOnBackorder = FNN(rsSTP!STP_QtyOnBackorder) - lngQty
    If Not rsSTP.EOF Then rsSTP.Update
    rsSTP.Close
    Set rsSTP = Nothing
    oPC.CO.CommitTrans
    Exit Function
    
ERRH:
    oPC.CO.RollbackTrans
    oError.SetError Err, "Z-StockManager:CancelCOL: " & Error, Now(), "", "", ""
    Exit Function

End Function
Public Sub MarkProductObsolete(pPID As String)
    oPC.CO.execute "UPDATE tPRODUCT SET P_Obsolete = 1 WHERE P_ID = '" & pPID & "'"
End Sub

Public Function MarkCOL(pCOLID As Long) ', pPID As String, pQty As Long)
Dim strSQL As String
Dim rsP As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim lngQty As Long
Dim strPID As String

On Error GoTo ERRH
    
    oPC.CO.execute "UPDATE tCOL SET COL_ActionTaken = 3 WHERE COL_ID = " & pCOLID
    
    
ERRH:
    oError.SetError Err, "Z-StockManager:CancelCOL: " & Error, Now(), "", "", ""
    Exit Function

End Function

Public Function IssuePO(pPO As a_PO)
On Error GoTo ERRH
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsCOL As ADODB.Recordset
Dim rsAGG As New ADODB.Recordset
'UPdate     STP_QtyOnBackorder
'           P_LastQtyFirmOrdered,P_LastQtySSOrdered,P_LastDateOrdered,P_LastPriceOrdered,
'           COL_ActionTaken
    rsAGG.Fields.Append "PID", adGUID
    rsAGG.Fields.Append "QTYFirm", adInteger
    rsAGG.Fields.Append "QTYSS", adInteger
    rsAGG.Open
    If Not pPO.CanCancel Then Exit Function
    Set rs = New ADODB.Recordset
    strSQL = "SELECT POL_Section,POL_ProductTypeID,POL_ID,POL_Section,POL_P_ID,POL_DL_ID,POL_QTYSS,POL_QTYFirm,POL_Price,POL_COLID FROM tPOL WHERE POL_TR_ID = " & pPO.TRID
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    
    Do While Not rs.EOF
        rsAGG.Find "PID = '" & rs!POL_P_ID & "'"
        If rsAGG.EOF Then
            rsAGG.AddNew
        End If
        rsAGG.Fields("PID") = (rs!POL_P_ID)
        rsAGG.Fields("QTYFIRM") = (rsAGG.Fields("QTYFIRM")) + (rs!POL_QtyFirm)
        rsAGG.Fields("QTYSS") = rsAGG.Fields("QTYSS") + (rs!POL_QtySS)
        rsAGG.Update
        Set rsSP = New ADODB.Recordset
   '     strSQL = "SELECT STP_QTYOnOrder FROM tStoreP Where STP_P_ID = '" & rs.Fields("POL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        strSQL = "SELECT STP_P_ID,STP_ST_ID,STP_QtyOnBackorder,STP_QTYONORDER,STP_QTYLASTSTOCKTAKE,STP_QTYRESERVED,STP_QTYONHAND,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED FROM tStoreP Where STP_P_ID = '" & rs.Fields("POL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        If rsSP.EOF Then  'This might be a new product
            rsSP.AddNew
            rsSP!STP_P_ID = rs!POL_P_ID
            rsSP!STP_ST_ID = oPC.Configuration.DefaultStoreID
            rsSP!STP_QtyOnBackorder = 0
            rsSP!STP_QtyOnOrder = FNN(rsSP!STP_QtyOnOrder) + FNN(rs.Fields("POL_QTYSS")) + FNN(rs.Fields("POL_QTYFIRM"))
            rsSP!STP_QtyLastStockTake = 0
            rsSP!STP_QTYRESERVED = 0
            rsSP!STP_QTYCOPIESONHAND = 0
            rsSP!STP_QTYONHAND = 0
            rsSP!STP_QTYCOPIESRESERVED = 0
            rsSP.Update
        Else
            rsSP.Fields("STP_QTYOnOrder") = FNN(rsSP.Fields("STP_QTYOnOrder")) + FNN(rs.Fields("POL_QTYSS")) + FNN(rs.Fields("POL_QTYFIRM"))
            rsSP.Update
        End If
        
        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_RRP,P_Section,P_ProductType_ID,P_DealID,P_SupplierID,P_QtyOnOrder,P_LastQtyFirmOrdered,P_LastQtySSOrdered,P_LastDateOrdered,P_LastPriceOrdered FROM tProduct Where P_ID = '" & rs.Fields("POL_P_ID") & "'"
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsP!P_LastQtyFirmOrdered = rs!POL_QtyFirm
            rsP!P_LastQtySSOrdered = rs!POL_QtySS
            rsP!P_QtyOnOrder = FNN(rsP!P_QtyOnOrder) + FNN(rs.Fields("POL_QTYSS")) + FNN(rs.Fields("POL_QTYFIRM"))
            rsP!P_LastDateOrdered = pPO.DocDate
            rsP!P_LastPriceOrdered = rs!POL_Price
            If Not pPO.ISForeignCurrency Then
                rsP!P_RRP = rs!POL_Price
            End If
            rsP!P_SupplierID = pPO.Supplier.ID
            rsP!P_DealID = rs!POL_DL_ID
            If rs!POL_ProductTypeID > 0 Then
                rsP!P_ProductType_ID = rs!POL_ProductTypeID
            End If
            rsP!P_Section = rs!POL_Section
        If Not rsP.EOF Then rsP.Update
        
        Set rsCOL = New ADODB.Recordset
      '  strSQL = "SELECT COL_ACtionTaken,COL_SOL_ID FROM tCOL Where COL_ID = " & rs.Fields("POL_COLID")
        strSQL = "SELECT COL_ACtionTaken,COL_SOL_ID FROM tCOL JOIN tPOL ON COL_P_ID = POL_P_ID JOIN tTR ON COL_TR_ID = TR_ID AND (COL_ACTIONTAKEN = 0 or COL_ACTIONTAKEN = 2) AND TR_Status = 3 AND COL_P_ID = '" & rs!POL_P_ID & "'"
        rsCOL.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        If Not rsCOL.EOF Then
            rsCOL.Fields("COL_ActionTaken") = 1
            rsCOL.Fields("COL_SOL_ID") = rs.Fields("POL_ID")
            rsCOL.Update
        End If
        
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    rsSP.Close
    rsCOL.Close
    Set rsCOL = Nothing
    Set rsSP = Nothing
    rsP.Close
    Set rsP = Nothing
    Exit Function
ERRH:
    MsgBox Error
    oError.SetError Err, "z_Stockmanager:IssuePO: " & Error, Now(), "", "", ""
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    rsSP.Close
    rsP.Close
    Set rsSP = Nothing
    Set rsP = Nothing
    rsP.Close
    Set rsP = Nothing
    rsCOL.Close
    Set rsCOL = Nothing
    Exit Function
    Resume
End Function

Public Function IssueAPP(pAPP As a_APP)
On Error GoTo ERRH
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
Dim rsP As ADODB.Recordset

'UPdate     STP_QtyOnAppro,STP_QtyOnHand
'           P_QtyOnHand,P QtyOnAppro
    Set rs = New ADODB.Recordset
    strSQL = "SELECT APPL_P_ID,APPL_QTY FROM tAPPL WHERE APPL_TR_ID = " & pAPP.TRID
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic

    Do While Not rs.EOF

        Set rsSTP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnAppro,STP_QtyOnHand FROM tStoreP Where STP_P_ID = '" & rs.Fields("APPL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSTP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsSTP.Fields("STP_QtyOnAppro") = FNN(rsSTP.Fields("STP_QtyOnAppro")) + FNN(rs.Fields("APPL_QTY"))
            rsSTP.Fields("STP_QtyOnHand") = FNN(rsSTP.Fields("STP_QtyOnHand")) - FNN(rs.Fields("APPL_QTY"))
        If Not rsSTP.EOF Then rsSTP.Update

        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_LastApproToTPID,P_QtyOnHand,P_QtyOnAppro FROM tProduct Where P_ID = '" & rs.Fields("APPL_P_ID") & "'"
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) - FNN(rs.Fields("APPL_QTY"))
            rsP!P_QtyOnAppro = FNN(rsP!P_QtyOnAppro) + FNN(rs.Fields("APPL_QTY"))
            rsP!P_LastApproToTPID = pAPP.Customer.ID
        If Not rsP.EOF Then rsP.Update



        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    rsSTP.Close
    Set rsSTP = Nothing
    rsP.Close
    Set rsP = Nothing
    Exit Function
ERRH:
    MsgBox Error
    oError.SetError Err, "z_Stockmanager:IssuePO: " & Error, Now(), "", "", ""
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    rsSTP.Close
    rsP.Close
    Set rsSTP = Nothing
    Set rsP = Nothing
    rsP.Close
    Set rsP = Nothing
    Exit Function
    Resume
End Function
Public Function CancelAPP(pAPP As a_APP)
On Error GoTo ERRH
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
    If Not pAPP.CanCancel Then Exit Function
'UPdate     STP_QtyOnAppro,STP_QtyOnHand
'           P_QtyOnHand,P QtyOnAppro
    Set rs = New ADODB.Recordset
    strSQL = "SELECT APPL_P_ID,APPL_QTY FROM tAPPL WHERE APPL_TR_ID = " & pAPP.TRID
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic

    Do While Not rs.EOF

        Set rsSTP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnAppro,STP_QtyOnHand FROM tStoreP Where STP_P_ID = '" & rs.Fields("APPL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSTP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsSTP.Fields("STP_QtyOnAppro") = NonNegative_Lng(FNN(rsSTP.Fields("STP_QtyOnAppro")) - FNN(rs.Fields("APPL_QTY")))
            rsSTP.Fields("STP_QtyOnHand") = FNN(rsSTP.Fields("STP_QtyOnHand")) + FNN(rs.Fields("APPL_QTY"))
        If Not rsSTP.EOF Then rsSTP.Update

        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QtyOnHand,P_QtyOnAppro FROM tProduct Where P_ID = '" & rs.Fields("APPL_P_ID") & "'"
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) + FNN(rs.Fields("APPL_QTY"))
            rsP!P_QtyOnAppro = NonNegative_Lng(FNN(rsP!P_QtyOnAppro) - FNN(rs.Fields("APPL_QTY")))
        If Not rsP.EOF Then rsP.Update
        rs.MoveNext
    Loop
    oPC.CO.execute "UPDATE tTR SET TR_Status = 6 WHERE TR_ID = " & pAPP.TRID
    rs.Close
    Set rs = Nothing
    rsSTP.Close
    Set rsSTP = Nothing
    rsP.Close
    Set rsP = Nothing
    Exit Function
ERRH:
    oPC.CO.RollbackTrans
    MsgBox Error
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    rsSTP.Close
    Set rsSTP = Nothing
End Function


Public Function IssueDEL(pDEL As a_Delivery)
On Error GoTo ERRH
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsPOL As ADODB.Recordset
Dim rsCOL As ADODB.Recordset
Dim rsPO As ADODB.Recordset
Dim bFulfilled As Boolean
'Must update    P_LastPriceDelivered,P_LastQtyDelivered,P_QtyOnHand,P_LastDateDelivered
'               POL_QtyReceivedSofar,POL_FulfilledStatus
'               STP_QtyOnHand,STP_QTYONORDER,STP_LastReceived,STP_FirstReceived
'Must generate COLAllocation rows
    Set rs = New ADODB.Recordset
    strSQL = "SELECT DELL_ID,DELL_Discount,DELL_P_ID,DELL_QTYFIRM,DELL_QTYSS,DELL_PRICE,DELL_PRICESELL,DELL_POL_ID FROM tDELL WHERE DELL_TR_ID = " & pDEL.TRID
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    
    Do While Not rs.EOF
    
        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_RRP,P_SP,P_Cost,P_QtyOnOrder,P_DefaultDeliveryDays,P_LastPriceDelivered,P_LastQtyDelivered,P_QtyOnHand,P_LastDateDelivered FROM tProduct Where P_ID = '" & rs.Fields("DELL_P_ID") & "'"
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsP!P_cost = FNN(rs!DELL_Price) * ((100 - FNN(rs!DELL_Discount)) / 100)
            rsP!P_RRP = FNN(rs!DELL_Price) ' * ((100 - FNN(rs!DELL_Discount)) / 100)
            rsP!P_SP = FNN(rs!DELL_PriceSELL) ' * ((100 - FNN(rs!DELL_Discount)) / 100)
            rsP!P_LastPriceDelivered = FNN(rs!DELL_Price)
            rsP!P_LastQtyDelivered = FNN(rs!DELL_QTYFIRM) + FNN(rs!DELL_QTYSS)
            rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) + FNN(rs!DELL_QTYFIRM) + FNN(rs!DELL_QTYSS)
            rsP!P_QtyOnOrder = NonNegative_Long(FNN(rsP!P_QtyOnOrder) - FNN(rs!DELL_QTYFIRM) - FNN(rs!DELL_QTYSS))
            rsP!P_LastDateDelivered = pDEL.DocDate

'Note the update follows below

        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QTYRESERVED,STP_FirstReceived,STP_QTYONHAND,STP_LastReceived,STP_QTYONORDER,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED FROM tStoreP Where STP_P_ID = '" & rs.Fields("DELL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsSP.Fields("STP_QtyOnHand") = FNN(rsSP.Fields("STP_QtyOnHand")) + (FNN(rs!DELL_QTYFIRM) + FNN(rs!DELL_QTYSS))
            If rs.Fields("DELL_POL_ID") > 0 Then
                rsSP!STP_QtyOnOrder = NonNegative_Long(FNN(rsSP!STP_QtyOnOrder) - FNN(rs!DELL_QTYFIRM) - FNN(rs!DELL_QTYSS))
            End If
            rsSP.Fields("STP_LastReceived") = pDEL.DocDate
            If IsNull(rsSP.Fields("STP_FirstReceived")) Then
                rsSP.Fields("STP_FirstReceived") = pDEL.DocDate
            End If
        If Not rsSP.EOF Then rsSP.Update
        
        Set rsPOL = New ADODB.Recordset
        strSQL = "SELECT TR_DATE,POL_QtyReceivedSofar,POL_Fulfilled,POL_TR_ID,POL_QTYSS,POL_QTYFirm FROM tPOL JOIN tTR ON POL_TR_ID = TR_ID Where POL_ID = " & rs.Fields("DELL_POL_ID")
        rsPOL.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            If Not rsPOL.EOF Then
                rsPOL.Fields("POL_QtyReceivedSofar") = FNN(rsPOL.Fields("POL_QtyReceivedSofar")) + (FNN(rs!DELL_QTYFIRM) + FNN(rs!DELL_QTYSS))
                If (rsPOL.Fields("POL_QTYSS") + FNN(rsPOL.Fields("POL_QTYFirm"))) > FNN(rsPOL.Fields("POL_QtyReceivedSofar")) Then
                    rsPOL.Fields("POL_FulFilled") = "OS"
                Else
                    rsPOL.Fields("POL_FulFilled") = "FUL"
                End If
            End If
        If Not rsPOL.EOF Then
            rsPOL.Update
            rsSP.Fields("STP_QTYONORDER") = NonNegative_Lng(FNN(rsSP.Fields("STP_QTYONORDER")) - (FNN(rs!DELL_QTYFIRM) + FNN(rs!DELL_QTYSS)))
            rsP!P_DefaultDeliveryDays = DateDiff("d", rsPOL!TR_Date, pDEL.DocDate)
        End If
        If Not rsP.EOF Then rsP.Update
        
        
        
        If Not rsPOL.EOF Then
            strSQL = "Select POL_Fulfilled from tPOL WHERE POL_TR_ID = " & rsPOL!POL_TR_ID
            Set rsPO = New ADODB.Recordset
            rsPO.Open strSQL, oPC.CO, adOpenKeyset, adLockReadOnly
            bFulfilled = True
            Do While Not rsPO.EOF
                If Trim(rsPO!POL_Fulfilled) <> "FUL" Then
                    bFulfilled = False
                    Exit Do
                End If
                rsPO.MoveNext
            Loop
            If bFulfilled Then
                oPC.CO.execute "UPDATE tTR SET TR_STATUS = 4 WHERE TR_ID = " & rsPOL!POL_TR_ID
            End If
        End If
        
        'Generate COLAllocation rows
'        strSQL = "SELECT * FROM vCOL_ISSUED  WHERE COL_P_ID = '" & FNS(rs.Fields("DELL_P_ID")) & "'"
'        Set rsCOL = New ADODB.Recordset
'        rsCOL.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
'        Do While Not rsCOL.EOF
'            strSQL = "INSERT INTO tCOLAllocation (COLALLOC_COL_ID,COLALLOC_DELL_ID,COLALLOC_Qty,COLALLOC_Status) VALUES (" & FNS(rs.Fields("DELL_ID")) & "," & FNS(rsCOL.Fields("COL_ID")) & ",0,1)"
'            oPC.CO.execute strSQL, oPC.CO
'            rsCOL.MoveNext
'        Loop
'
        rs.MoveNext
    Loop
    On Error Resume Next
    rs.Close
    rsSP.Close
    rsP.Close
    rsPOL.Close
    rsCOL.Close
    Set rsPOL = Nothing
    Set rsCOL = Nothing
    Set rsSP = Nothing
    Set rsP = Nothing
    Set rs = Nothing
    Exit Function
ERRH:
'Resume
    MsgBox "z_Stockmanager:IssueDEL: " & Error
    oError.SetError Err, "z_Stockmanager:IssueDEL: " & Error, Now(), "", "", ""
    On Error Resume Next
    rs.Close
    rsSP.Close
    rsP.Close
    rsPOL.Close
    rsCOL.Close
    Set rsPOL = Nothing
    Set rsCOL = Nothing
    Set rsSP = Nothing
    Set rsP = Nothing
    Set rs = Nothing
    Exit Function
    Resume
End Function

Public Function IssueAPPR(pAPPR As a_APPR)
On Error GoTo ERRH
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsAPPL As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsAPP As ADODB.Recordset
Dim bFulfilled As Boolean

'UPdate     STP_QtyOnBackorder
'           P_LastQtyFirmOrdered,P_LastQtySSOrdered,P_LastDateOrdered,P_LastPriceOrdered,
    Set rs = New ADODB.Recordset
    strSQL = "SELECT APPRL_APPL_ID,APPRL_Qty FROM tAPPRL WHERE APPRL_TR_ID = " & pAPPR.TRID
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic

    Do While Not rs.EOF
        
        Set rsAPPL = New ADODB.Recordset
        strSQL = "SELECT APPL_TR_ID,APPL_P_ID,APPL_QTY,APPL_QTYReturned,APPL_Fulfilled FROM tAPPL WHERE APPL_ID = " & rs!APPRL_APPL_ID
        rsAPPL.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsAPPL!APPL_QtyReturned = FNN(rsAPPL!APPL_QtyReturned) + FNN(rs!APPRL_Qty)
            If FNN(rsAPPL!APPL_Qty) <= FNN(rsAPPL!APPL_QtyReturned) Then
                rsAPPL!APPL_Fulfilled = "FUL"
            End If
        If Not rsAPPL.EOF Then rsAPPL.Update
        
        
        strSQL = "Select APPL_QTY,APPL_QTYReturned from tAPPL WHERE APPL_TR_ID = " & rsAPPL!APPL_TR_ID
        Set rsAPP = New ADODB.Recordset
        rsAPP.Open strSQL, oPC.CO, adOpenForwardOnly, adLockReadOnly
        bFulfilled = True
        Do While Not rsAPP.EOF
            If FNN(rsAPP!APPL_Qty) > FNN(rsAPP!APPL_QtyReturned) Then
                bFulfilled = False
                Exit Do
            End If
            rsAPP.MoveNext
        Loop
        If bFulfilled Then
            oPC.CO.execute "UPDATE tTR SET TR_STATUS = 4 WHERE TR_ID = " & rsAPPL!APPL_TR_ID
        End If
        rsAPP.Close
        Set rsAPP = Nothing
        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnHand,STP_QtyOnAppro FROM tStoreP Where STP_P_ID = '" & rsAPPL.Fields("APPL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.CO, adOpenKeyset, adLockOptimistic
            rsSP.Fields("STP_QtyOnHand") = FNN(rsSP.Fields("STP_QtyOnHand")) + FNN(rs.Fields("APPRL_Qty"))
            rsSP.Fields("STP_QtyOnAppro") = NonNegative_Lng(FNN(rsSP.Fields("STP_QtyOnAppro")) - FNN(rs.Fields("APPRL_Qty")))
        If Not rsSP.EOF Then rsSP.Update
        rsSP.Close
        Set rsSP = Nothing

        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QtyOnHand,P_QtyOnAppro FROM tProduct Where P_ID = '" & rsAPPL.Fields("APPL_P_ID") & "'"
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) + FNN(rs!APPRL_Qty)
            rsP!P_QtyOnAppro = FNN(rsP!P_QtyOnAppro) - FNN(rs!APPRL_Qty)
        If Not rsP.EOF Then rsP.Update
        rsP.Close
        Set rsP = Nothing


        rs.MoveNext
        rsAPPL.Close
        Set rsAPPL = Nothing
        
    Loop
    rs.Close
    On Error Resume Next
    Set rs = Nothing
    rsSP.Close
    Set rsSP = Nothing
    rsP.Close
    Set rsP = Nothing
    Exit Function
ERRH:
    MsgBox Error
    oError.SetError Err, "z_Stockmanager:IssuePO: " & Error, Now(), "", "", ""
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    rsSP.Close
    rsP.Close
    Set rsSP = Nothing
    Set rsP = Nothing
    rsP.Close
    Set rsP = Nothing
    Exit Function
    Resume
End Function
Public Function CancelAPPR(pAPPR As a_APPR)
On Error GoTo ERRH
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsAPPL As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsAPP As ADODB.Recordset
Dim bFulfilled As Boolean
    Set rs = New ADODB.Recordset
    strSQL = "SELECT APPRL_APPL_ID,APPRL_Qty FROM tAPPRL WHERE APPRL_TR_ID = " & pAPPR.TRID
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic

    Do While Not rs.EOF
        Set rsAPPL = New ADODB.Recordset
        strSQL = "SELECT APPL_TR_ID,APPL_P_ID,APPL_QTY,APPL_QTYReturned,APPL_Fulfilled FROM tAPPL WHERE APPL_ID = " & rs!APPRL_APPL_ID
        rsAPPL.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsAPPL!APPL_QtyReturned = NonNegative_Lng(FNN(rsAPPL!APPL_QtyReturned) - FNN(rs!APPRL_Qty))
            rsAPPL!APPL_Fulfilled = "OS"
        If Not rsAPPL.EOF Then rsAPPL.Update
        oPC.CO.execute "UPDATE tTR SET TR_STATUS = 3 WHERE TR_ID = " & rsAPPL!APPL_TR_ID 'It cant be COMPLETE but has to be ISSUED only
        
        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnHand,STP_QtyOnAppro FROM tStoreP Where STP_P_ID = '" & rsAPPL.Fields("APPL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsSP.Fields("STP_QtyOnHand") = FNN(rsSP.Fields("STP_QtyOnHand")) - FNN(rs.Fields("APPRL_Qty"))
            rsSP.Fields("STP_QtyOnAppro") = FNN(rsSP.Fields("STP_QtyOnAppro")) + FNN(rs.Fields("APPRL_Qty"))
        If Not rsSP.EOF Then rsSP.Update

        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QtyOnHand,P_QtyOnAppro FROM tProduct Where P_ID = '" & rsAPPL.Fields("APPL_P_ID") & "'"
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) - FNN(rs!APPRL_Qty)
            rsP!P_QtyOnAppro = FNN(rsP!P_QtyOnAppro) + FNN(rs!APPRL_Qty)
        If Not rsP.EOF Then rsP.Update



        rs.MoveNext
    Loop
    oPC.CO.execute "UPDATE tTR SET TR_STATUS = 6 WHERE TR_ID = " & pAPPR.TRID 'It cant be COMPLETE but has to be ISSUED only
    rs.Close
    Set rs = Nothing
    rsSP.Close
    Set rsSP = Nothing
    rsP.Close
    Set rsP = Nothing
    Exit Function
ERRH:
    MsgBox Error
    oError.SetError Err, "z_Stockmanager:IssuePO: " & Error, Now(), "", "", ""
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    rsSP.Close
    rsP.Close
    Set rsSP = Nothing
    Set rsP = Nothing
    rsP.Close
    Set rsP = Nothing
    Exit Function
    Resume
End Function

Public Function CancelPO(pTRID As Long)
On Error GoTo ERRH
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsCOL As ADODB.Recordset
Dim rsPOL As ADODB.Recordset
Dim rsMAXPOL As ADODB.Recordset
Dim Cmd As ADODB.Command
Dim par As ADODB.Parameter

'UPdate     STP_QtyOnBackorder
'           P_LastQtyFirmOrdered,P_LastQtySSOrdered,P_LastDateOrdered,P_LastPriceOrdered,
'           COL_ActionTaken
    Set rs = New ADODB.Recordset
    strSQL = "SELECT POL_ID,POL_P_ID,POL_QTYSS,POL_QTYFirm,POL_Price,POL_COLID FROM tPOL WHERE POL_TR_ID = " & pTRID & " AND POL_Fulfilled <> 'CAN'"
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    oPC.CO.BeginTrans
    Do While Not rs.EOF
    
        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QTYONORDER FROM tStoreP Where STP_P_ID = '" & rs.Fields("POL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsSP.Fields("STP_QTYONORDER") = NonNegative_Lng(FNN(rsSP.Fields("STP_QTYONORDER")) - FNN(rs.Fields("POL_QTYSS")) - FNN(rs.Fields("POL_QTYFIRM")))
        If Not rsSP.EOF Then rsSP.Update
        rsSP.Close
        Set rsSP = Nothing
        
        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QTYONORDER FROM tPRODUCT WHERE P_ID = '" & rs.Fields("POL_P_ID") & "'"
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsP!P_QtyOnOrder = NonNegative_Lng(FNN(rsP!P_QtyOnOrder) - FNN(rs.Fields("POL_QTYSS")) - FNN(rs.Fields("POL_QTYFIRM")))
        If Not rsP.EOF Then rsP.Update
        rsP.Close
        Set rsP = Nothing

        
        Set Cmd = New ADODB.Command
        Cmd.CommandText = "sp_ResetLastPO"
        Cmd.CommandType = adCmdStoredProc
        Set par = Cmd.CreateParameter("PID", adGUID, adParamInput)
        Cmd.Parameters.Append par
        par.Value = rs.Fields("POL_P_ID")
        Cmd.ActiveConnection = oPC.CO
        Cmd.execute
        Set Cmd = Nothing
        
        Set rsCOL = New ADODB.Recordset
        strSQL = "SELECT COL_ACtionTaken FROM tCOL Where COL_ID = " & FNN(rs.Fields("POL_COLID"))
        rsCOL.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        Set rsPOL = New ADODB.Recordset
        rsPOL.Open "SELECT * FROM tPOL WHERE POL_COLID = " & FNN(rs.Fields("POL_COLID")), oPC.CO, adOpenForwardOnly, adLockOptimistic
        If Not rsCOL.EOF Then
            If rsPOL.EOF Then 'There is no other POL for this COL
                rsCOL.Fields("COL_ActionTaken") = 0
            End If
            rsCOL.Update
        End If
        On Error Resume Next
        rsPOL.Close
        rsCOL.Close
        Set rsPOL = Nothing
        Set rsCOL = Nothing
        Set rsCOL = Nothing
        rs.MoveNext
    Loop
    oPC.CO.execute "UPDATE tTR SET TR_STATUS = 6 WHERE TR_ID = " & pTRID
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    rsSP.Close
    rsCOL.Close
    Set rsCOL = Nothing
    Set rsSP = Nothing
    rsP.Close
    Set rsP = Nothing
    On Error GoTo ERRH
    oPC.CO.CommitTrans
    Exit Function
ERRH:
    oPC.CO.RollbackTrans
    MsgBox Error
    oError.SetError Err, "z_Stockmanager:IssuePO: " & Error, Now(), "", "", ""
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    rsSP.Close
    rsP.Close
    Set rsSP = Nothing
    Set rsP = Nothing
    rsP.Close
    Set rsP = Nothing
    rsCOL.Close
    Set rsCOL = Nothing
    Exit Function
    Resume
'Resume
End Function
Public Function CancelPOL(pPOLID As Long, pPID As String, pQty As Long)
Dim Cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSQL As String
Dim rsP As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
On Error GoTo ERRH
    oPC.CO.BeginTrans
    oPC.CO.execute "UPDATE tPOL SET POL_Fulfilled = 'CAN' WHERE POL_ID = " & pPOLID
    
    Set rsP = New ADODB.Recordset
    strSQL = "SELECT P_QTYONORDER FROM tPRODUCT WHERE P_ID = '" & pPID & "'"
    rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        rsP!P_QtyOnOrder = NonNegative_Lng(FNN(rsP!P_QtyOnOrder) - pQty)
    If Not rsP.EOF Then rsP.Update
    rsP.Close
    Set rsP = Nothing
    
        Set Cmd = New ADODB.Command
        Cmd.CommandText = "sp_ResetLastPO"
        Cmd.CommandType = adCmdStoredProc
        Set par = Cmd.CreateParameter("PID", adGUID, adParamInput)
        Cmd.Parameters.Append par
        par.Value = pPID
        Cmd.ActiveConnection = oPC.CO
        Cmd.execute
        Set Cmd = Nothing
    
    Set rsSTP = New ADODB.Recordset
    strSQL = "SELECT STP_QTYONORDER FROM tStoreP WHERE STP_P_ID = '" & pPID & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
    rsSTP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        rsSTP!STP_QtyOnOrder = NonNegative_Lng(FNN(rsSTP!STP_QtyOnOrder) - pQty)
    If Not rsSTP.EOF Then rsSTP.Update
    rsSTP.Close
    Set rsSTP = Nothing
    oPC.CO.CommitTrans
    Exit Function
    
ERRH:
    oPC.CO.RollbackTrans
    oError.SetError Err, "Z-StockManager:CancelPOL: " & Error, Now(), "", "", ""
    Exit Function
End Function
Public Function MarkPOLFulfilled(pPOLID As Long, pPID As String, pQty As Long)
Dim Cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSQL As String
Dim rsP As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
On Error GoTo ERRH
    oPC.CO.BeginTrans
    oPC.CO.execute "UPDATE tPOL SET POL_Fulfilled = 'FUL' WHERE POL_ID = " & pPOLID
    
    Set rsP = New ADODB.Recordset
    strSQL = "SELECT P_QTYONORDER FROM tPRODUCT WHERE P_ID = '" & pPID & "'"
    rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    rsP!P_QtyOnOrder = NonNegative_Lng(FNN(rsP!P_QtyOnOrder) - pQty)
    If Not rsP.EOF Then rsP.Update
    rsP.Close
    Set rsP = Nothing
    
    
    Set rsSTP = New ADODB.Recordset
    strSQL = "SELECT STP_QTYONORDER FROM tStoreP WHERE STP_P_ID = '" & pPID & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
    rsSTP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    rsSTP!STP_QtyOnOrder = NonNegative_Lng(FNN(rsSTP!STP_QtyOnOrder) - pQty)
    If Not rsSTP.EOF Then rsSTP.Update
    rsSTP.Close
    Set rsSTP = Nothing
    oPC.CO.CommitTrans
    Exit Function
    
ERRH:
    oPC.CO.RollbackTrans
    oError.SetError Err, "Z-StockManager:CancelPOL: " & Error, Now(), "", "", ""
    Exit Function
End Function

Public Function AdjustStock(pPID As String, pCount As Long, pStaffID As Long) As Long
Dim rsP As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
Dim strSQL As String
Dim lngCalc As Long
Dim lngAdj As Long

On Error GoTo ERRH
    AdjustStock = 0
    
 '   oPC.CO.BeginTrans
    
    Set rsP = New ADODB.Recordset
    strSQL = "SELECT P_QtyOnHand,P_DateLastStockTake,P_QtyLastStockTake FROM tProduct WHERE P_ID = '" & pPID & "'"
    rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    lngCalc = FNN(rsP!P_QtyOnHand)
    If lngCalc <> pCount Then
        rsP!P_QtyOnHand = pCount
        rsP.Update
        
        lngAdj = pCount - lngCalc
        oPC.CO.execute "INSERT INTO tADJ (ADJ_PID,ADJ_Adjustment,ADJ_Counted,ADJ_Date,ADJ_StaffID) VALUES ('" & pPID & "'," & lngAdj & "," & pCount & ", {d '" & ReverseDate(Date) & "'}," & pStaffID & ")"
        
        Set rsSTP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnHand,STP_DateLastStockTake,STP_QtyLastStockTake FROM tStoreP WHERE STP_P_ID = '" & pPID & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSTP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        rsSTP!STP_QTYONHAND = pCount
        rsSTP.Update
        rsSTP.Close
    End If
   
    rsP.Close
    Set rsP = Nothing
    Set rsSTP = Nothing
    Exit Function
ERRH:
    oPC.CO.RollbackTrans
    AdjustStock = 99
    Exit Function
    Resume
End Function
Public Function CancelDEL(pDEL As a_Delivery)
On Error GoTo ERRH
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsDELL As ADODB.Recordset
Dim rsPOL As ADODB.Recordset
Dim rsCOL As ADODB.Recordset
Dim rsPO As ADODB.Recordset
Dim rsMAXDELL As ADODB.Recordset
Dim Cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim bFulfilled As Boolean
'Must update    P_LastPriceDelivered,P_LastQtyDelivered,P_QtyOnHand,P_LastDateDelivered
'               POL_QtyReceivedSofar,POL_FulfilledStatus
'               STP_QtyOnHand,STP_QTYONORDER,STP_LastReceived,STP_FirstReceived
MsgBox "dont forget STP_LastReceived,STP_FirstReceived"
    Set rs = New ADODB.Recordset
    strSQL = "SELECT DELL_ID,DELL_P_ID,DELL_QTYFIRM,DELL_QTYSS,DELL_PRICE,DELL_POL_ID FROM tDELL WHERE DELL_TR_ID = " & pDEL.TRID
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    oPC.CO.BeginTrans
    Do While Not rs.EOF


        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QTYRESERVED,STP_FirstReceived,STP_QTYONHAND,STP_LastReceived,STP_QTYONORDER,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED FROM tStoreP Where STP_P_ID = '" & rs.Fields("DELL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            rsSP.Fields("STP_QtyOnHand") = FNN(rsSP.Fields("STP_QtyOnHand")) - (rs!DELL_QTYFIRM + rs!DELL_QTYSS)
            If FNN(rs.Fields("DELL_POL_ID")) > 0 Then
                rsSP.Fields("STP_QTYONORDER") = FNN(rsSP.Fields("STP_QTYONORDER")) + (rs!DELL_QTYFIRM + rs!DELL_QTYSS)
            End If
            rsSP.Fields("STP_LastReceived") = CDate(0)
            If rsSP.Fields("STP_FirstReceived") = pDEL.DocDate Then
                rsSP.Fields("STP_FirstReceived") = CDate(0)
            End If
        If Not rsSP.EOF Then rsSP.Update
        rsSP.Close
        
        Set Cmd = New ADODB.Command
        Cmd.CommandText = "sp_ResetLastDelivery"
        Cmd.CommandType = adCmdStoredProc
        Set par = Cmd.CreateParameter("PID", adGUID, adParamInput)
        Cmd.Parameters.Append par
        par.Value = rs.Fields("DELL_P_ID")
        Cmd.ActiveConnection = oPC.CO
        Cmd.execute
        Set Cmd = Nothing
'''''''''''''''''''''''''''''''''''
'        Set rsMAXDELL = New ADODB.Recordset
'        strSQL = "SELECT MAX(DELL_ID) as MAXID FROM ActiveDELLs WHERE DELL_P_ID = '" & rs!DELL_P_ID & "'"
'        rsMAXDELL.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
'        If Not rsMAXDELL.EOF Then
'            Set rsDELL = New ADODB.Recordset
'            rsDELL.Open "SELECT TR_Date,DELL_PRICE,DELL_QtyFirm,DELL_QtySS,TR_Date FROM tDELL JOIN tTR ON DELL_TR_ID = TR_ID WHERE DELL_ID = " & rsMAXDELL.Fields("MAXID"), oPC.CO, adOpenKeyset, adLockPessimistic
'        End If
        Set rsP = New ADODB.Recordset
        rsP.Open "SELECT P_QtyOnHand,P_LastQtyDelivered,P_LastDateDelivered,P_LastPriceDelivered FROM tProduct Where P_ID = '" & rs.Fields("DELL_P_ID") & "'", oPC.CO, adOpenKeyset, adLockPessimistic
'            If rsMAXDELL.EOF Then
'                rsP!P_LastPriceDelivered = 0           'CANT SET THESE correctly - DAYEND WILL
'                rsP!P_LastQtyDelivered = 0
'                rsP!P_LastDateDelivered = CDate(0)
'            Else
'                rsP!P_LastPriceDelivered = rsDELL!DELL_Price
'                rsP!P_LastQtyDelivered = FNN(rsDELL!DELL_QtyFirm) + FNN(rsDELL!DELL_QtySS)
'                rsP!P_LastDateDelivered = rsDELL!TR_Date
'            End If
            rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) - FNN(rs!DELL_QTYFIRM) - FNN(rs!DELL_QTYSS)
        If Not rsP.EOF Then rsP.Update
        rsP.Close
        Set rsP = Nothing
'        Set rsMAXDELL = Nothing
        
''''''''''''''''''''''''''''''
        
        Set rsPOL = New ADODB.Recordset
        strSQL = "SELECT POL_QtyReceivedSofar,POL_Fulfilled,POL_TR_ID,POL_QTYSS,POL_QTYFirm FROM tPOL Where POL_ID = " & rs.Fields("DELL_POL_ID")
        rsPOL.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        If Not rsPOL.EOF Then
            rsPOL.Fields("POL_QtyReceivedSofar") = NonNegative_Lng(FNN(rsPOL.Fields("POL_QtyReceivedSofar")) - (FNN(rs!DELL_QTYFIRM) + FNN(rs!DELL_QTYSS)))
            rsPOL.Fields("POL_FulFilled") = "OS"
            rsPOL.Update
        End If

        oPC.CO.execute "UPDATE tTR SET TR_STATUS = 6 WHERE TR_ID = " & pDEL.TRID
        rs.MoveNext
    Loop
    Set rsPOL = Nothing
    Set rsSP = Nothing
    Set rsP = Nothing
    Set rs = Nothing
    oPC.CO.CommitTrans
    Exit Function
ERRH:
'Resume
    oPC.CO.RollbackTrans
    MsgBox "z_Stockmanager:IssueDEL: " & Error
    oError.SetError Err, "z_Stockmanager:IssueDEL: " & Error, Now(), "", "", ""
    On Error Resume Next
    rs.Close
    rsSP.Close
    rsP.Close
    rsPOL.Close
    rsCOL.Close
    Set rsPOL = Nothing
    Set rsCOL = Nothing
    Set rsSP = Nothing
    Set rsP = Nothing
    Set rs = Nothing
    Exit Function
    Resume
End Function

Public Sub ActionODPO(pRS As ADODB.Recordset)
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset

Dim strSQL As String

    Set rs = pRS
    On Error Resume Next
    oPC.CO.execute "DROP TABLE #TMP"
    oPC.CO.execute "DROP TABLE tREMINDER"
    
    On Error GoTo ERRH
    oPC.CO.execute "CREATE TABLE #TMP (POLID INT PRIMARY KEY,ACTN1 CHAR(1),ACTN2 CHAR(1),ACTN3 CHAR(2))"
    oPC.CO.execute "CREATE TABLE tREMINDER (POLID INT PRIMARY KEY)"
    rs.MoveFirst
    Do While Not rs.EOF
        oPC.CO.execute "INSERT INTO #TMP (POLID,ACTN1,ACTN2,ACTN3) VALUES(" & rs.Fields(0) & ",'" & FNS(rs.Fields(1)) & "','" & FNS(rs.Fields(2)) & "','" & FNS(rs.Fields(3)) & "')"
        rs.MoveNext
    Loop
    Set rs = Nothing
    strSQL = "SELECT #TMP.*, tPOL.* FROM #TMP JOIN tPOL ON POLID = POL_ID"
    Set rs2 = New ADODB.Recordset
    rs2.Open strSQL, oPC.CO, adOpenKeyset, adLockOptimistic
    Do While Not rs2.EOF
        If FNS(rs2!POL_LastAction) > "" Then
            oPC.CO.execute "INSERT INTO tPOLA (POLA_POL_ID,POLA_DATE,POLA_CODE) VALUES (" & rs2!POL_ID & ",'" & ReverseDate(rs2!POL_LastActionDate) & "','" & Trim(rs2!POL_LastAction) & "')"
        End If
        Select Case Right(rs2!ACTN3, 1)
        Case "W"
            rs2!POL_ETA = DateAdd("ww", Left(rs2!ACTN3, 1), Date)
        Case "M"
            rs2!POL_ETA = DateAdd("m", Left(rs2!ACTN3, 1), Date)
        End Select
        Select Case rs2!ACTN2
        Case "C"
            rs2!POL_Fulfilled = "CAN"
            strSQL = "UPDATE tPRODUCT SET P_QTYOnOrder =  dbo.NonNegative(P_QTYOnOrder - " & FNN(rs2!POL_QtySS) + FNN(rs2!POL_QtyFirm) - FNN(rs2!POL_QTYReceivedSoFar) & ") WHERE P_ID = '" & rs2!POL_P_ID & "'"
            oPC.CO.execute strSQL
            strSQL = "UPDATE tStoreP SET STP_QTYOnOrder =  dbo.NonNegative(STP_QTYOnOrder - " & FNN(rs2!POL_QtySS) + FNN(rs2!POL_QtyFirm) - FNN(rs2!POL_QTYReceivedSoFar) & ") WHERE STP_P_ID = '" & rs2!POL_P_ID & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
            oPC.CO.execute strSQL
        Case "R"
            oPC.CO.execute "INSERT INTO tReminder (POLID) VALUES(" & rs2!POL_ID & ")"
        End Select
        Select Case rs2!ACTN1
        Case "O"
            strSQL = "UPDATE tPRODUCT SET P_Status =  'O' WHERE P_ID = '" & rs2!POL_P_ID & "'"
            oPC.CO.execute strSQL
        Case "R"
            strSQL = "UPDATE tPRODUCT SET P_Status =  'R' WHERE P_ID = '" & rs2!POL_P_ID & "'"
            oPC.CO.execute strSQL
        End Select
        rs2!POL_LastAction = FNS(rs2!ACTN1) & FNS(rs2!ACTN2) & FNS(rs2!ACTN3)
        rs2!POL_LastActionDate = Date
        rs2.Update
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing
    Exit Sub
ERRH:
    MsgBox Error
    Exit Sub
    Resume
End Sub
Public Sub ActionODCO(pRS As ADODB.Recordset)
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset

Dim strSQL As String
    If pRS.EOF Then Exit Sub
    Set rs = pRS
    On Error Resume Next
    oPC.CO.execute "DROP TABLE #TMP"
    On Error GoTo ERRH
    oPC.CO.execute "CREATE TABLE #TMP (COLID INT PRIMARY KEY,ACTN1 CHAR(2),ACTN2 VARCHAR(150))"
    rs.MoveFirst
    Do While Not rs.EOF
        oPC.CO.execute "INSERT INTO #TMP (COLID,ACTN1,ACTN2) VALUES(" & rs.Fields(0) & ",'" & FNS(rs.Fields(1)) & "','" & FNS(rs.Fields(2)) & "')"
        rs.MoveNext
    Loop
    Set rs = Nothing
    strSQL = "SELECT #TMP.*, tCOL.* FROM #TMP JOIN tCOL ON COLID = COL_ID  "
    Set rs2 = New ADODB.Recordset
    rs2.Open strSQL, oPC.CO, adOpenKeyset, adLockOptimistic
    Do While Not rs2.EOF
        If FNS(rs2!ACTN2) > "" And FNS(rs2!COL_LastAction) > "" Then
            oPC.CO.execute "INSERT INTO tCOLA (COLA_COL_ID,COLA_DATE,COLA_REPORT) VALUES (" & rs2!COL_ID & ",'" & ReverseDate(rs2!COL_LastActionDate) & "','" & rs2!COL_LastAction & "')"
        End If
        Select Case Right(rs2!ACTN1, 1)
        Case "W"
            rs2!COL_ETA = DateAdd("ww", Left(rs2!ACTN1, 1), Date)
        Case "M"
            rs2!COL_ETA = DateAdd("m", Left(rs2!ACTN1, 1), Date)
        End Select
        rs2!COL_LastAction = FNS(rs2!ACTN2)  ' & FNS(rs2!ACTN2) & FNS(rs2!ACTN3)
        rs2!COL_LastActionDate = Date
        rs2.Update
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing
    Exit Sub
ERRH:
    MsgBox Error
    Exit Sub
    Resume
End Sub

Public Function IssueTF(pTF As a_TF)
On Error GoTo ERRH
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
Dim rsP As ADODB.Recordset

'UPdate     STP_QtyOnTFro,STP_QtyOnHand
'           P_QtyOnHand,P QtyOnTFro
    Set rs = New ADODB.Recordset
    strSQL = "SELECT TFRL_P_ID,TFRL_QTY FROM tTFRL WHERE TFRL_TR_ID = " & pTF.TRID
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic

    Do While Not rs.EOF

        Set rsSTP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnAppro,STP_QtyOnHand FROM tStoreP Where STP_P_ID = '" & rs.Fields("TFRL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSTP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            If pTF.INOUT = "OUT" Then
                rsSTP.Fields("STP_QtyOnHand") = FNN(rsSTP.Fields("STP_QtyOnHand")) - FNN(rs.Fields("TFRL_QTY"))
            Else
                rsSTP.Fields("STP_QtyOnHand") = FNN(rsSTP.Fields("STP_QtyOnHand")) + FNN(rs.Fields("TFRL_QTY"))
            End If
        If Not rsSTP.EOF Then rsSTP.Update

        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QtyOnHand,P_QtyOnAppro FROM tProduct Where P_ID = '" & rs.Fields("TFRL_P_ID") & "'"
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            If pTF.INOUT = "OUT" Then
                rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) - FNN(rs.Fields("TFRL_QTY"))
            Else
                rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) + FNN(rs.Fields("TFRL_QTY"))
            End If
        If Not rsP.EOF Then rsP.Update



        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    rsSTP.Close
    Set rsSTP = Nothing
    rsP.Close
    Set rsP = Nothing
    Exit Function
ERRH:
    MsgBox Error
    oError.SetError Err, "z_Stockmanager:IssueTF: " & Error, Now(), "", "", ""
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    rsSTP.Close
    rsP.Close
    Set rsSTP = Nothing
    Set rsP = Nothing
    rsP.Close
    Set rsP = Nothing
    Exit Function
    Resume
End Function
Public Function CancelTF(pTF As a_TF)
On Error GoTo ERRH
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsSTP As ADODB.Recordset


    Set rs = New ADODB.Recordset
    strSQL = "SELECT TFRL_P_ID,TFRL_QTY FROM tTFRL WHERE TFRL_TR_ID = " & pTF.TRID
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    
    
    Do While Not rs.EOF

        Set rsSTP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnAppro,STP_QtyOnHand FROM tStoreP Where STP_P_ID = '" & rs.Fields("TFRL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSTP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            If pTF.INOUT = "OUT" Then
                rsSTP.Fields("STP_QtyOnHand") = FNN(rsSTP.Fields("STP_QtyOnHand")) + FNN(rs.Fields("TFRL_QTY"))
            Else
                rsSTP.Fields("STP_QtyOnHand") = FNN(rsSTP.Fields("STP_QtyOnHand")) - FNN(rs.Fields("TFRL_QTY"))
            End If
        If Not rsSTP.EOF Then rsSTP.Update

        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QtyOnHand,P_QtyOnAppro FROM tProduct Where P_ID = '" & rs.Fields("TFRL_P_ID") & "'"
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
            If pTF.INOUT = "OUT" Then
                rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) + FNN(rs.Fields("TFRL_QTY"))
            Else
                rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) - FNN(rs.Fields("TFRL_QTY"))
            End If
        If Not rsP.EOF Then rsP.Update
        rs.MoveNext
    Loop
    oPC.CO.execute "UPDATE tTR SET TR_Status = 6 WHERE TR_ID = " & pTF.TRID
    rs.Close
    Set rs = Nothing
    rsSTP.Close
    Set rsSTP = Nothing
    rsP.Close
    Set rsP = Nothing
    
    
    Exit Function
ERRH:
    oPC.CO.RollbackTrans
    MsgBox Error
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    rsSTP.Close
    Set rsSTP = Nothing
End Function

Public Function GenerateReturnsPerPub(pTPID As Long, pMinDate As Date, pMaxDate As Date, pStaffID As Long) As Integer
'NOTE Min date is the most recent date and maxdate is the least recent date
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsR As ADODB.Recordset
Dim rsRL As ADODB.Recordset
Dim rsTR As ADODB.Recordset
Dim oCode As New z_Code
Dim lngTRID  As Long
Dim i As Long
Dim lngCurrentRLID As Long
Dim strRef As String
Dim lngPrice As Long
Dim dblDiscount As Double
Dim iCount As Integer
Dim strDefaultCo As String
On Error GoTo ERRH


    'Get titles that have been received SeeSafe on supplier invoices prior to or on the pDATE parameter.
    'The Titles should have QtyOnHand > 0
'    strSQL = "SELECT SUM(dbo.tDELL.DELL_QtySS) AS QTYSS,P_ID,P_QTYONHAND,P_RRP,Min(DELL_Discount) as Discount " _
'            & " FROM dbo.tTR as tTR_1 INNER JOIN " _
'            & " dbo.tDELL ON tTR_1.TR_ID = dbo.tDELL.DELL_TR_ID INNER JOIN " _
'            & " dbo.tDEL ON dbo.tDELL.DELL_TR_ID = dbo.tDEL.DEL_ID INNER JOIN " _
'            & " dbo.tPRODUCT ON dbo.tPRODUCT.P_ID = dbo.tDELL.DELL_P_ID " _
'            & " WHERE (tTR_1.TR_Status = 3 OR tTR_1.TR_Status = 4) AND tTR_1.TR_TP_ID = " & pTPID & " AND dbo.tProduct.P_QTYONHAND > 0 " _
'            & " AND DEL_SupplierInvoiceDate >= {d '" & ReverseDate(pMaxDate) & "'} AND DEL_SupplierInvoiceDate  <= {d '" & ReverseDate(pMinDate) & "'} AND DELL_RetQtyAppr is NULL  " _
'            & " GROUP BY P_ID,P_QTYONHAND,P_RRP " _
'            & " HAVING Sum(dbo.tDELL.DELL_QtySS) > 0 "

    strSQL = "SELECT dbo.tDELL.DELL_QtySS AS QTYSS, dbo.tProduct.P_ID, dbo.tProduct.P_QtyOnHand, dbo.tProduct.P_RRP, dbo.tDELL.DELL_Discount AS Discount, " _
             & " dbo.tDELL.DELL_ID, dbo.tDEL.DEL_SupplierInvoiceRef, dbo.tDEL.DEL_SupplierInvoiceDate, tTR_1.TR_ID " _
             & " FROM dbo.tTR tTR_1 INNER JOIN " _
             & " dbo.tDELL ON tTR_1.TR_ID = dbo.tDELL.DELL_TR_ID INNER JOIN " _
             & " dbo.tDEL ON dbo.tDELL.DELL_TR_ID = dbo.tDEL.DEL_ID INNER JOIN " _
             & " dbo.tProduct ON dbo.tProduct.P_ID = dbo.tDELL.DELL_P_ID " _
             & " WHERE (tTR_1.TR_Status = 3 OR tTR_1.TR_Status = 4) AND (tTR_1.TR_TP_ID = 409) AND (dbo.tProduct.P_QtyOnHand > 0) AND " _
             & " (dbo.tDEL.DEL_SupplierInvoiceDate >= {d '" & ReverseDate(pMaxDate) & "'} AND (dbo.tDEL.DEL_SupplierInvoiceDate <= {d '" & ReverseDate(pMinDate) & "'} " _
             & " AND (dbo.tDELL.DELL_RetQtyAppr IS NULL) AND (dbo.tDELL.DELL_QtySS > 0)"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, oPC.CO, adOpenForwardOnly, adLockOptimistic
    
    If rs.EOF Then
        GenerateReturnsPerPub = 0
        Exit Function
    Else
        GenerateReturnsPerPub = 1
    End If
    RaiseEvent eMax(rs.RecordCount)
    
    oPC.CO.BeginTrans
    
    Set rsTR = New ADODB.Recordset
  '  rsTR.CursorLocation = adUseClient
    rsTR.Open "Select * FROM tTR WHERE TR_ID = 0", oPC.CO, adOpenKeyset, adLockOptimistic
    rsTR.AddNew
        rsTR!TR_Date = ReverseDate(Date)
        rsTR!TR_TP_ID = pTPID
        rsTR!TR_CaptureDate = ReverseDateTime(Now)
        strDefaultCo = oPC.Configuration.DefaultCompany.CompanyCode
        rsTR!TR_Code = strDefaultCo & "R" & oCode.GetNextCode(enReturn, "GETANDWRITE")
        rsTR!TR_Type = 11
        rsTR!TR_Status = 2
        rsTR!TR_StaffID = pStaffID
        rsTR!TR_COMP_ID = oPC.Configuration.DefaultCOMPID
    rsTR.Update
    rsTR.Bookmark = rsTR.Bookmark
    lngTRID = rsTR("TR_ID")
    
    Set rsR = New ADODB.Recordset
    rsR.Open "Select * FROM tR WHERE R_ID = 0", oPC.CO, adOpenKeyset, adLockOptimistic
    rsR.AddNew
        rsR!R_ID = lngTRID
        rsR!R_Status = 0
    rsR.Update
    
    Set rsRL = New ADODB.Recordset
    rsRL.Open "Select * FROM tRL WHERE RL_ID = 0", oPC.CO, adOpenKeyset, adLockOptimistic
    i = 0
    Do While Not rs.EOF
        rsRL.AddNew
            rsRL!RL_P_ID = rs!P_ID
            rsRL!RL_TR_ID = rs!TR_ID
            rsRL!RL_DELL_ID = rs!DELL_ID
            rsRL!RL_Price = rs!P_RRP
            rsRL!RL_QtySystem = rs!P_QtyOnHand
            rsRL!RL_QtyRequested = GetMin(CLng(rs!P_QtyOnHand), CLng(rs!QtySS))
            rsRL!RL_SIRef = FNS(rs!DEL_SUpplierInvoiceRef) & ":" & Format(rs!DEL_SUpplierInvoiceDate, "dd/mm/yyyy")
        rsRL.Update
        rs.MoveNext
        i = i + 1
        RaiseEvent eProgress(i)
    Loop
    
'    strSQL = "SELECT DELL_QTYSS,RL_ID,DEL_SupplierInvoiceRef,DEL_SupplierInvoiceDate,DELL_Price,DELL_Discount " _
'            & " FROM tDEL " _
'            & " INNER JOIN tDELL on DELL_TR_ID = DEL_ID " _
'            & " INNER JOIN tRL ON RL_P_ID = DELL_P_ID " _
'            & " INNER JOIN tTR ON DELL_TR_ID = TR_ID " _
'            & " WHERE RL_TR_ID = " & lngTRID & " AND (TR_STATUS = 3 or TR_Status = 4) " _
'            & " AND DEL_SupplierInvoiceDate >= {d '" & ReverseDate(pMaxDate) & "'} " _
'            & " AND DEL_SupplierInvoiceDate <= {d '" & ReverseDate(pMinDate) & "'} " _
'            & " ORDER BY RL_ID,DEL_SupplierInvoiceDate"
'    rs.Close
'    rs.Open strSQL, oPC.CO, adOpenForwardOnly, adLockOptimistic
'    Do While Not rs.EOF
'        lngCurrentRLID = rs!RL_ID
'        strRef = ""
'        iCount = 0
'        Do While lngCurrentRLID = rs!RL_ID And iCount < 4
'            If rs!DELL_QTYSS > 0 Then
'                strRef = strRef & IIf(Len(strRef) > 0, ",", "") & FNS(rs!DEL_SUpplierInvoiceRef) & ": " & Format(rs!DEL_SUpplierInvoiceDate, "dd/mm/yyyy") & "(" & rs!DELL_QTYSS & ")"
'                lngPrice = rs!DELL_Price
'                dblDiscount = rs!DELL_Discount
'            End If
'            rs.MoveNext
'            iCount = 1 + 1
'            If rs.EOF Then Exit Do
'        Loop
'        oPC.CO.execute "UPDATE tRL SET RL_SIRef = '" & strRef & "' ,RL_Discount = " & dblDiscount & " , RL_Price = " & lngPrice & " WHERE RL_ID = " & lngCurrentRLID
'    Loop
    rs.Close
    Set rs = Nothing
    rsTR.Close
    Set rsTR = Nothing
    rsR.Close
    Set rsR = Nothing
    rsRL.Close
    Set rsRL = Nothing
    
    oPC.CO.CommitTrans
Exit Function

ERRH:
    MsgBox "Error in Generating returns =" & Error
    oPC.CO.RollbackTrans
    Exit Function
    Resume
End Function
Public Function GenerateReturns_Step1(pRID As Long, pRS As ADODB.Recordset, pStaffID As Long) As Boolean
Dim rs As ADODB.Recordset
Dim strSQL As String
    
    On Error Resume Next
    oPC.CO.execute "DROP TABLE #Returns"
    On Error GoTo ERRH
    GenerateReturns_Step1 = True
    oPC.CO.BeginTrans
    oPC.CO.execute "Create table #Returns (RLID INT,Counted Int,Requested Int)"
    Set rs = pRS
    rs.MoveFirst
    Do While Not rs.EOF
        oPC.CO.execute "INSERT INTO #Returns (RLID,COUNTED,Requested) VALUES (" & rs.Fields(0) & "," & rs.Fields("Counted") & "," & rs.Fields("Requested") & ")"
        If pRS!System <> pRS!COunted Then
            AdjustStock rs!PID, pRS!COunted, pStaffID
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    strSQL = "UPdate tRL SET tRL.RL_QTYReturned = #Returns.Requested,tRL.RL_QTYApproved = #Returns.Requested,tRL.RL_QTYRequested = #Returns.Requested,tRL.RL_QTYCounted = #Returns.Counted FROM tRL, #Returns   WHERE RL_TR_ID = " & pRID & " AND  RL_ID = RLID"
    oPC.CO.execute strSQL
    oPC.CO.execute "Update tTR Set TR_Status = 3,TR_STAFFID = " & pStaffID & "  WHERE TR_ID = " & pRID
    oPC.CO.CommitTrans
    Exit Function
    
ERRH:
    GenerateReturns_Step1 = False
    MsgBox Error
    oPC.CO.RollbackTrans
    Exit Function
    Resume
End Function
Public Function GenerateReturns_Step2(pRID As Long, pRS As ADODB.Recordset, pStaffID As Long) As Boolean
Dim rs As ADODB.Recordset
Dim strSQL As String
    On Error Resume Next
    oPC.CO.execute "DROP TABLE #Returns"
    On Error GoTo ERRH
    GenerateReturns_Step2 = True
    oPC.CO.BeginTrans
    oPC.CO.execute "Create table #Returns (RLID INT,Counted Int,Approved Int,Returned Int)"
    Set rs = pRS
    rs.MoveFirst
    Do While Not rs.EOF
        oPC.CO.execute "INSERT INTO #Returns (RLID,COUNTED,Approved,Returned) VALUES (" & rs.Fields("RLID") & "," & rs.Fields("Counted") & "," & rs.Fields("Approved") & "," & rs.Fields("Returned") & ")"
        If rs!System <> rs!COunted Then
            AdjustStock rs!PID, pRS!COunted, pStaffID
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
'    strSQL = "UPDATE tDELL SET DELL_RetQtyAppr =  " & #Returns.Approved & " WHERE DELL_ID = " & rs!DELLID
    
    strSQL = "UPdate tRL SET tRL.RL_QTYReturned = #Returns.Returned,tRL.RL_QTYApproved = #Returns.Approved,tRL.RL_QTYCounted = #Returns.Counted FROM tRL, #Returns   WHERE RL_TR_ID = " & pRID & " AND  RL_ID = RLID"
    oPC.CO.execute strSQL
    strSQL = "UPdate tProduct SET tProduct.P_QTYOnHand = tProduct.P_QTYOnHand - tRL.RL_QtyReturned FROM tPRODUCT JOIN tRL ON RL_P_ID = P_ID  WHERE RL_TR_ID = " & pRID
    oPC.CO.execute strSQL
    strSQL = "UPdate tSTOREP SET tStoreP.STP_QTYOnHand = tStoreP.STP_QTYOnHand - tRL.RL_QtyReturned FROM tStoreP JOIN tRL ON RL_P_ID = STP_P_ID WHERE RL_TR_ID = " & pRID
    oPC.CO.execute strSQL
    

    oPC.CO.execute "Update tTR Set TR_Status = 4,TR_Date = '" & ReverseDate(Date) & "',TR_CaptureDate = '" & ReverseDateTime(Now) & "' WHERE TR_ID = " & pRID
    oPC.CO.CommitTrans
    GenerateReturns_Step2 = True
    Exit Function
    
ERRH:
    GenerateReturns_Step2 = False
    MsgBox Error
    oPC.CO.RollbackTrans
    Exit Function
    Resume
End Function
Public Function IssueReturn(pTRID As Long)
Dim rs As ADODB.Recordset
Dim strSQL As String
 '   oPC.CO.BeginTrans
 '   Set rs = New ADODB.Recordset
 '   rs.Open "SELECT * FROM tRL WHERE RL_TR_ID = " & pTRID, oPC.CO, adOpenForwardOnly, adLockReadOnly
 '   rs.MoveFirst
 '   Do While Not rs.EOF
        strSQL = "UPdate tProduct SET tProduct.P_QTYOnHand = tProduct.P_QTYOnHand - tRL.RL_QtyReturned FROM tPRODUCT JOIN tRL ON RL_P_ID = P_ID  WHERE RL_TR_ID = " & pTRID
        oPC.CO.execute strSQL
        strSQL = "UPdate tSTOREP SET tStoreP.STP_QTYOnHand = tStoreP.STP_QTYOnHand - tRL.RL_QtyReturned FROM tStoreP JOIN tRL ON RL_P_ID = STP_P_ID WHERE RL_TR_ID = " & pTRID
        oPC.CO.execute strSQL
 '       rs.MoveNext
 '   Loop
    

    oPC.CO.execute "Update tTR Set TR_Status = 4 WHERE TR_ID = " & pTRID
 '   rs.Close
 '   Set rs = Nothing
'    oPC.CO.CommitTrans
    Exit Function
    
ERRH:
    MsgBox Error
'    oPC.CO.RollbackTrans
    Exit Function

End Function
Public Function CountProductsPerPTID(pPTID As Long) As Long
Dim Cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set Cmd = New ADODB.Command
    Cmd.CommandText = "sp_CountProductsPerPT"
    Cmd.CommandType = adCmdStoredProc
    
    Set par = Cmd.CreateParameter("@PTID", adInteger, adParamInput)
    Cmd.Parameters.Append par
    par.Value = pPTID
    Set par = Cmd.CreateParameter("@COUNT", adInteger, adParamOutput)
    Cmd.Parameters.Append par
    
    Cmd.ActiveConnection = oPC.CO
    Cmd.execute
    
    CountProductsPerPTID = Cmd.Parameters(1)
    Set Cmd = Nothing

End Function

Public Sub MergePTs(pPTOld As Long, pPTNew As Long)
    oPC.CO.execute "Update tProduct SET P_PRODUCTTYPE_ID = " & pPTNew & " WHERE P_PRODUCTTYPE_ID = " & pPTOld
End Sub
Public Sub DeleteUnusedPTs()
    oPC.CO.execute "DELETE dbo.tPT FROM dbo.tPT LEFT OUTER JOIN dbo.tProduct ON dbo.tProduct.P_ProductType_ID = dbo.tPT.PT_ID Where (dbo.tProduct.P_ProductType_ID Is Null)"
End Sub
Public Sub DeleteRelatedDocs(pTP As Long)
    oPC.CO.execute "DELETE FROM dbo.tTR Where TR_TP_ID = pTP"
End Sub

Public Function CancelR(pR As a_R)
On Error GoTo ERRH
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsSP As ADODB.Recordset
    If Not pR.CanCancel Then Exit Function
    Set rs = New ADODB.Recordset
    strSQL = "SELECT RL_P_ID,RL_QTYRETURNED FROM tRL WHERE RL_TR_ID = " & pR.TRID
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
    
    oPC.CO.BeginTrans
    Do While Not rs.EOF
        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnBackorder,STP_QTYRESERVED,STP_QTYONHAND,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED FROM tStoreP Where STP_P_ID = '" & rs.Fields("RL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        rsSP.Fields("STP_QTYONHAND") = FNN(rsSP.Fields("STP_QTYONHAND")) + FNN(rs.Fields("RL_QTYRETURNED"))
        If Not rsSP.EOF Then rsSP.Update
        
        strSQL = "SELECT P_QTYONHAND FROM tPRODUCT WHERE P_ID = '" & rs.Fields("RL_P_ID") & "'"
        Set rsP = New ADODB.Recordset
        rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) + FNN(rs.Fields("RL_QTYRETURNED"))
        If Not rsP.EOF Then rsP.Update
        rsP.Close
        Set rsP = Nothing
        
        
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    rsSP.Close
    Set rsSP = Nothing
    oPC.CO.execute "UPDATE tTR SET TR_STATUS = 6 WHERE TR_ID = " & pR.TRID
    oPC.CO.CommitTrans
    Exit Function
ERRH:
    oPC.CO.RollbackTrans
    MsgBox Error
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    rsSP.Close
    Set rsSP = Nothing
    Exit Function
    Resume
End Function

Public Sub ReverseReturnToRequested(pTRID As Long)
Dim rs As ADODB.Recordset
Dim strSQL As String
    
    oPC.CO.BeginTrans
    
        strSQL = "UPdate tProduct SET tProduct.P_QTYOnHand = ISNULL(tProduct.P_QTYOnHand,0) + tRL.RL_QtyReturned FROM tPRODUCT JOIN tRL ON RL_P_ID = P_ID  WHERE RL_TR_ID = " & pTRID
        oPC.CO.execute strSQL
        strSQL = "UPdate tSTOREP SET tStoreP.STP_QTYOnHand = IsNull(tStoreP.STP_QTYOnHand,0) + tRL.RL_QtyReturned FROM tStoreP JOIN tRL ON RL_P_ID = STP_P_ID WHERE RL_TR_ID = " & pTRID
        oPC.CO.execute strSQL

    oPC.CO.execute "Update tTR Set TR_Status = 3 WHERE TR_ID = " & pTRID
    oPC.CO.CommitTrans
    Exit Sub
    
ERRH:
    MsgBox Error
    oPC.CO.RollbackTrans
    Exit Sub
End Sub
Public Function CancelRL(pRLID As Long, pPID As String, pQty As Long)
Dim Cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSQL As String
Dim rsP As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
On Error GoTo ERRH
    oPC.CO.BeginTrans
    oPC.CO.execute "UPDATE tRL SET RL_STatus = 'CAN' WHERE RL_ID = " & pRLID
    
    Set rsP = New ADODB.Recordset
    strSQL = "SELECT P_QTYONHAND FROM tPRODUCT WHERE P_ID = '" & pPID & "'"
    rsP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) + IIf(IsNull(pQty), 0, pQty)
    If Not rsP.EOF Then rsP.Update
    rsP.Close
    Set rsP = Nothing
    
    
    Set rsSTP = New ADODB.Recordset
    strSQL = "SELECT STP_QTYONHAND FROM tStoreP WHERE STP_P_ID = '" & pPID & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
    rsSTP.Open strSQL, oPC.CO, adOpenKeyset, adLockPessimistic
        rsSTP!STP_QTYONHAND = FNN(rsSTP!STP_QTYONHAND) + IIf(IsNull(pQty), 0, pQty)
    If Not rsSTP.EOF Then rsSTP.Update
    rsSTP.Close
    Set rsSTP = Nothing
    oPC.CO.CommitTrans
    Exit Function
    
ERRH:
    oPC.CO.RollbackTrans
    oError.SetError Err, "Z-StockManager:CancelRL: " & Error, Now(), "", "", ""
    Exit Function
End Function

Public Function GetLabelsToPrint(pType As String, PStart As Boolean, Optional PID As Long, Optional pCODE As String) As XArrayDB
Dim rs As New ADODB.Recordset
Static XA As New XArrayDB


    rs.CursorLocation = adUseClient
    If pType = "D" Then
        rs.Open "SELECT P_CODE,P_EAN,P_TITLE,P_SP,SUM(DELL_QTYFIRM) as QTYFIRM,SUM(DELL_QTYSS) as QTYSS FROM tDELL JOIN tProduct ON DELL_P_ID = P_ID WHERE DELL_TR_ID = " & PID & " GROUP BY P_CODE,P_TITLE,P_EAN,P_SP", oPC.CO, adOpenStatic, adLockReadOnly
    ElseIf pType = "T" Then
        rs.Open "SELECT P_CODE,P_EAN,P_TITLE,P_SP,SUM(TFRL_QTY) as QTYFIRM,0 as QTYSS FROM tTFRL JOIN tProduct ON TFRL_P_ID = P_ID WHERE TFRL_TR_ID = " & PID & " GROUP BY P_CODE,P_TITLE,P_EAN,P_SP", oPC.CO, adOpenStatic, adLockReadOnly
    ElseIf pType = "P" Then
        rs.Open "SELECT P_CODE,P_EAN,P_TITLE,P_SP,1 as QTYFIRM,0 as QTYSS FROM tProduct WHERE P_Code = '" & pCODE & "' GROUP BY P_CODE,P_TITLE,P_EAN,P_SP", oPC.CO, adOpenStatic, adLockReadOnly
    End If
  '  XA.ReDim 1, rs.RecordCount + Marker, 1, 7
    If PStart Then XA.ReDim 1, 0, 1, 7
 '   tmpMarker = Marker
    Do While Not rs.EOF
        XA.InsertRows XA.UpperBound(1) + 1, 1
        XA(XA.UpperBound(1), 1) = FormatCode(FNS(rs!P_Code), True)
        XA(XA.UpperBound(1), 2) = FNS(rs!P_Title)
        XA(XA.UpperBound(1), 3) = Format(FNN(rs!P_SP) / oPC.Configuration.DefaultCurrency.Divisor, "####0.00")
        XA(XA.UpperBound(1), 4) = CStr(FNN(rs!QtySS) + FNN(rs!QtyFirm))
        XA(XA.UpperBound(1), 5) = False
        XA(XA.UpperBound(1), 6) = FNS(rs!P_EAN)
        XA(XA.UpperBound(1), 7) = FNN(rs!P_SP)
  '      XA(rs.AbsolutePosition + tmpMarker, 6) = FNN(rs!P_SP)
        rs.MoveNext
  '      Marker = Marker + 1
    Loop
    rs.Close
    Set GetLabelsToPrint = XA
End Function

Public Function GetApproSlips(pAPPID As Long) As ADODB.Recordset
Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM vAPPROSLIPS WHERE TR_ID = " & pAPPID & " ORDER BY AUTHOR", oPC.CO, adOpenForwardOnly, adLockReadOnly
    Set GetApproSlips = rs
    Set rs = Nothing
End Function

Sub FixMissingInvoiceRefsOnReturn(pRID As Long, pMinDate As Date, pMaxDate As Date)
Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim strRef As String
Dim iCount As Integer
Dim lngPrice As Long
Dim dblDiscount As Double
Dim lngCurrentRLID As Long

    strSQL = "SELECT DELL_QTYSS,RL_ID,DEL_SupplierInvoiceRef,DEL_SupplierInvoiceDate,DELL_Price,DELL_Discount " _
            & " FROM tDEL " _
            & " INNER JOIN tDELL on DELL_TR_ID = DEL_ID " _
            & " INNER JOIN tRL ON RL_P_ID = DELL_P_ID " _
            & " INNER JOIN tTR ON DELL_TR_ID = TR_ID " _
            & " WHERE RL_TR_ID = " & pRID & " AND (TR_STATUS = 3 or TR_Status = 4) " _
            & " AND DEL_SupplierInvoiceDate >= {d '" & ReverseDate(pMaxDate) & "'} " _
            & " AND DEL_SupplierInvoiceDate <= {d '" & ReverseDate(pMinDate) & "'} " _
            & " ORDER BY RL_ID,DEL_SupplierInvoiceDate"
    rs.Open strSQL, oPC.CO, adOpenForwardOnly, adLockOptimistic
    Do While Not rs.EOF
        lngCurrentRLID = rs!RL_ID
        strRef = ""
        iCount = 0
        Do While lngCurrentRLID = rs!RL_ID And iCount < 4
            If rs!DELL_QTYSS > 0 Then
                strRef = strRef & IIf(Len(strRef) > 0, ",", "") & FNS(rs!DEL_SUpplierInvoiceRef) & ": " & Format(rs!DEL_SUpplierInvoiceDate, "dd/mm/yyyy") & "(" & rs!DELL_QTYSS & ")"
                lngPrice = rs!DELL_Price
                dblDiscount = rs!DELL_Discount
            End If
            rs.MoveNext
            iCount = 1 + 1
            If rs.EOF Then Exit Do
        Loop
        oPC.CO.execute "UPDATE tRL SET RL_SIRef = '" & strRef & "' ,RL_Discount = " & dblDiscount & " , RL_Price = " & lngPrice & " WHERE RL_ID = " & lngCurrentRLID
    Loop
    rs.Close
    Set rs = Nothing

End Sub

