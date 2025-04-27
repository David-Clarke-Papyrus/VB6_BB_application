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
Attribute VB_Exposed = False
Option Explicit
Event eMax(Max As Long)
Event eProgress(Progress As Long)
Dim fs As Scripting.FileSystemObject
Dim txtStream As Scripting.TextStream
Dim strSQL As String
Public Function PossibleIncorrectServiceCharge(pTRID As Long) As Boolean
      Dim OpenResult As Integer
      Dim cmd As ADODB.Command
      Dim par As ADODB.Parameter
      Dim r As Long
      '-------------------------------
18540     OpenResult = oPC.OpenDBSHort
      '-------------------------------
18550     Set cmd = New ADODB.Command
18560     cmd.CommandText = "HasPossibleIncorrectServiceCharge"
18570     cmd.commandType = adCmdStoredProc
          
18580     Set par = Nothing
18590     Set par = cmd.CreateParameter("@TRID", adInteger, adParamInput, , pTRID)
18600     cmd.Parameters.Append par
          
18610     Set par = Nothing
18620     Set par = cmd.CreateParameter("@OK", adInteger, adParamOutput)
18630     cmd.Parameters.Append par
          
18640     cmd.ActiveConnection = oPC.COShort
18650     cmd.execute
          
18660     r = FNB(cmd.Parameters("@OK"))
          
18670     If r = 1 Then
18680         PossibleIncorrectServiceCharge = True
18690     Else
18700         PossibleIncorrectServiceCharge = False
18710     End If
18720     Set cmd = Nothing
      '---------------------------------------------------
18730     If OpenResult = 0 Then oPC.DisconnectDBShort
      '---------------------------------------------------


18740     Exit Function
          

End Function
Friend Function ReservationExists(pTPID As Long, pPID As String) As Boolean
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_ReservationExists"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pPID", adVarChar, , 15)
    cmd.Parameters.Append par
    par.Value = pPID
    Set par = cmd.CreateParameter("@pTPID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = pTPID
    Set par = cmd.CreateParameter("@COUNT", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    ReservationExists = (cmd.Parameters(1) > 0)
    Set cmd = Nothing

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.ReservationExists(pTPID,pPID)", Array(pTPID, pPID)
End Function

Public Function IssueInvoiceLine(pPID As String, pbDeleted As Boolean, pCOLID As Long, pPIID As Long, pProforma As Boolean, pQty As Long) As Boolean
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "UpdateInvoiceLine"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PID", adGUID, adParamInput, , pPID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@LineDeleted", adBoolean, adParamInput, , pbDeleted)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@IL_COL_ID", adInteger, adParamInput, , pCOLID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@PIID", adInteger, adParamInput, , pPIID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@Proforma", adBoolean, adParamInput, , pProforma)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@QTY", adInteger, adParamInput, , pQty)
    cmd.Parameters.Append par
    Set par = Nothing
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.IssueInvoiceLine(pPID,pbDeleted,pCOLID,pPIID,pProforma,pQty)", Array(pPID, _
         pbDeleted, pCOLID, pPIID, pProforma, pQty)
End Function
Public Function IssueInvoice(pInvoice As a_Invoice, IsPreDelivery As Boolean, Optional pStatus As enStatus) As String
    On Error GoTo errHandler
Dim strSQL As String
Dim rsIL As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsPI As ADODB.Recordset
'Dim rsSP As ADODB.Recordset
Dim rsAPP As ADODB.Recordset
Dim rsSales As ADODB.Recordset
Dim rsWants As ADODB.Recordset
Dim rsCOL As ADODB.Recordset
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

        PostInvoiceLines pInvoice.InvoiceID, pStatus, IsPreDelivery
        Exit Function


EXITH:

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.IssueInvoice(pInvoice,pStatus)", Array(pInvoice, _
         pStatus)
End Function
Public Function IssueGDN(pGDN As a_GDN, Optional pStatus As enStatus) As String
    On Error GoTo errHandler

   
        PostGDNLines pGDN.GDNID, pStatus
        Exit Function

EXITH:

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.IssueGDN(pGDN,pWasProforma,pToProForma,pStatus)", Array(pGDN, _
          pStatus)
End Function
Public Sub CalculateInvoiceTableSummaryData(pTRID As Long, pPaidAsCash As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "CalculateInvoiceTableSummaryData"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TRID", adInteger, adParamInput, , pTRID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@AmtPaidAsCash", adInteger, adParamInput, , pPaidAsCash)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CalculateInvoiceTableSummaryData(pTRID,pPaidAsCash)", Array(pTRID, _
         pPaidAsCash)
End Sub

Public Sub CalculateGDNTableSummaryData(pTRID As Long, pPaidAsCash As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "CalculateGDNTableSummaryData"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TRID", adInteger, adParamInput, , pTRID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@AmtPaidAsCash", adInteger, adParamInput, , pPaidAsCash)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CalculateGDNTableSummaryData(pTRID,pPaidAsCash)", Array(pTRID, _
         pPaidAsCash)
End Sub

Public Sub AUTOGenerateReturnFromGDN(pINVID As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "CreateApproReturnForInvonAppro"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TRID", adInteger, adParamInput, , pINVID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@TRSTATUS", adInteger, adParamInput, , 4)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@ARID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing

    Set par = cmd.CreateParameter("@POSITION", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.AUTOGenerateReturnFromGDN(pINVID)", pINVID
End Sub

Public Function CancelInvoice(pInvoice As a_Invoice, Optional pDontResetStatus As Boolean) As Boolean
    On Error GoTo errHandler

        If Not pDontResetStatus Then
            pInvoice.SetStatus stVOID
        End If
            
        PostInvoiceLines pInvoice.InvoiceID, pInvoice.Status, pInvoice.IsPreDelivery
        
        Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CancelInvoice(pInvoice,pDontResetStatus)", Array(pInvoice, _
         pDontResetStatus)
End Function
Public Function CancelGDN(pGDN As a_GDN, Optional pDontResetStatus As Boolean) As Boolean
    On Error GoTo errHandler

        If Not pDontResetStatus Then
            pGDN.SetStatus stVOID
        End If
            
        PostGDNLines pGDN.GDNID, pGDN.Status
        
        Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CancelGDN(pInvoice,pDontResetStatus)", Array(pGDN, _
         pDontResetStatus)
End Function

Private Function PostInvoiceLines(pINVOICEID As Long, pNewStatus As enStatus, IsPreDelivery As Boolean)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "PostInvoicelines"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@INVOICEID", adInteger, adParamInput, , pINVOICEID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@STATUS", adInteger, adParamInput, , pNewStatus)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@IsPreDelivery", adBoolean, adParamInput, , IsPreDelivery)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.PostInvoiceLines(pInvoiceID,pNewStatus)", Array(pINVOICEID, pNewStatus)
End Function
Private Function PostGDNLines(pGDNID As Long, pNewStatus As enStatus)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "PostGDNLines"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    
    Set par = cmd.CreateParameter("@INVOICEID", adInteger, adParamInput, , pGDNID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@STATUS", adInteger, adParamInput, , pNewStatus)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.PostGDNLines(pGDNID,pNewStatus)", Array(pGDNID, pNewStatus)
End Function

Private Function PostCO(pCOID As Long, pNewStatus As enStatus)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "PostCO"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    
    Set par = cmd.CreateParameter("@COID", adInteger, adParamInput, , pCOID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@STATUS", adInteger, adParamInput, , pNewStatus)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.PostCO(pCOID,pNewStatus)", Array(pCOID, pNewStatus)
End Function
Public Function PostTransfer(pTransferID As Long, pNewStatus As enStatus, pIsCancellation As Boolean)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "PostTransfer"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    
    Set par = cmd.CreateParameter("@TRANSFERID", adInteger, adParamInput, , pTransferID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@STATUS", adInteger, adParamInput, , pNewStatus)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@IsCancellation", adBoolean, adParamInput, , pIsCancellation)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    ErrorIn "z_StockManager.PostTransfer(pTransferID,pNewStatus)", Array(pTransferID, pNewStatus)
End Function

Private Function PostPO(pPOID As Long, pNewStatus As enStatus, pForeign As Boolean, pSupplierID As Long)
    On Error GoTo errHandler
Dim lng As Long
Dim strPos As String
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "PostPO"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@POID", adInteger, adParamInput, , pPOID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@STATUS", adInteger, adParamInput, , pNewStatus)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@FOREIGN", adBoolean, adParamInput, , pForeign)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@SUPPLIERID", adInteger, adParamInput, , pSupplierID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@ERR", adInteger, adParamOutput, , lng)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@ERRPOS", adVarChar, adParamOutput, 10, strPos)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort  'THis is used only by PBKS_POSSVR so it must use this connection - DO NOT CHANGE
    
    cmd.execute
    
    If FNN(cmd.Parameters("@err")) <> 0 Then
        MsgBox "Error in PostPO " & FNN(cmd.Parameters("@Err")) & " at position " & FNS(cmd.Parameters("@ERRPOS"))
    End If
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.PostPO(pPOID,pNewStatus,pForeign,pSupplierID)", Array(pPOID, pNewStatus, _
         pForeign, pSupplierID)
End Function
Public Function IssueCreditNote(pCN As a_CN, pWasProforma As Boolean, pToProForma As Boolean)
    On Error GoTo errHandler
Dim strSQL As String
Dim rsCNL As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsPI As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsSales As ADODB.Recordset
Dim OpenResult As Integer
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
    oPC.COShort.execute "UPDATE tTR SET TR_Status = 4,TR_CaptureDate = GetDate() WHERE TR_ID = " & pCN.TRID
    Set rsCNL = New ADODB.Recordset
    strSQL = "SELECT CNL_P_ID,CNL_PI_ID,CNL_QTY,CNL_QTYDAM,CNL_PRICE,CNL_SalesValue,CNL_VATVALUE,IL_DISCOUNTRATE,IL_VATRATE FROM tCNL LEFT JOIN tILINE ON CNL_IL_ID = IL_ID WHERE CNL_TR_ID = " & pCN.TRID
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rsCNL.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
    
    Do While Not rsCNL.eof
    
        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QtyOnHand,P_QtyCopiesOnHand,p_QtyReserved,P_QtyExpectedBack,P_QtyTotalSold,P_COST FROM tProduct Where P_ID = '" & rsCNL.Fields("CNL_P_ID") & "'"
        rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        Set rsPI = New ADODB.Recordset
        strSQL = "SELECT PI_OnReserve,PI_FirstSoldDate,PI_ExpectedBack,PI_SoldDate,PI_SoldTo_ID FROM tPI Where PI_ID = " & rsCNL.Fields("CNL_PI_ID")
        rsPI.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QTYRESERVED,STP_QTYONHAND,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED FROM tStoreP Where STP_P_ID = '" & rsCNL.Fields("CNL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
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
                If Not rsSP.eof Then
                    rsSP.Fields("STP_QTYONHAND") = FNN(rsSP.Fields("STP_QTYONHAND")) + rsCNL.Fields("CNL_QTY") - FNN(rsCNL.Fields("CNL_QTYDAM"))
                End If
                rsP.Fields("P_QtyOnHand") = FNN(rsP.Fields("P_QTYONHAND")) + FNN(rsCNL.Fields("CNL_QTY")) - FNN(rsCNL.Fields("CNL_QTYDAM"))
                rsP.Fields("P_QtyTotalSold") = FNN(rsP.Fields("P_QtyTotalSold")) - FNN(rsCNL.Fields("CNL_QTY"))
            End If
            Set rsSales = New ADODB.Recordset
            rsSales.Open "SELECT * FROM tSALES_CYr WHERE SCY_P_ID = '" & rsCNL.Fields("CNL_P_ID") & "' AND SCY_Dte = '" & ReverseDate(pCN.DOCDate) & "' AND SCY_TP_ID = " & pCN.Customer.ID, oPC.COShort, adOpenDynamic, adLockOptimistic
            If rsSales.eof Then
                    rsSales.AddNew
                    rsSales.Fields("SCY_P_ID") = rsCNL.Fields("CNL_P_ID")
                    rsSales.Fields("SCY_Dte") = pCN.DOCDate
                    rsSales.Fields("SCY_TP_ID") = pCN.Customer.ID
            End If
            rsSales.Fields("SCY_QTY") = FNN(rsSales.Fields("SCY_QTY")) - FNN(rsCNL.Fields("CNL_QTY"))
            rsSales.Fields("SCY_SaleValue") = FNN(rsSales.Fields("SCY_SaleValue")) - (FNN(rsCNL.Fields("CNL_QTY")) * (FNDBL((rsCNL.Fields("CNL_PRICE")) * (100 - FNN(rsCNL.Fields("IL_DISCOUNTRATE"))) / 100)))
            rsSales.Fields("SCY_COSTValue") = FNN(rsSales.Fields("SCY_COSTValue")) - (FNN(rsCNL.Fields("CNL_QTY")) * (FNDBL((rsP.Fields("P_COST")) * (100 - FNN(rsCNL.Fields("IL_DISCOUNTRATE"))) / 100)))
            rsSales.Fields("SCY_SaleValueEXVAT") = FNN(rsSales.Fields("SCY_SaleValueEXVAT")) - (FNN(rsCNL.Fields("CNL_QTY")) * (FNDBL((rsCNL.Fields("CNL_PRICE")) * (100 - FNN(rsCNL.Fields("IL_DISCOUNTRATE"))) / 100))) / ((100 + FNDBL(rsCNL.Fields("IL_VATRATE"))) / 100)
            rsSales.Fields("SCY_COSTValueEXVAT") = FNN(rsSales.Fields("SCY_COSTValueEXVAT")) - (FNN(rsCNL.Fields("CNL_QTY")) * (FNDBL((rsP.Fields("P_COST")) * (100 - FNN(rsCNL.Fields("IL_DISCOUNTRATE"))) / 100))) / ((100 + FNDBL(rsCNL.Fields("IL_VATRATE"))) / 100)
            rsSales.Update
            
            
        End If
        
        
        If Not rsSP.eof Then rsSP.Update
        If Not rsP.eof Then rsP.Update
        If Not rsPI.eof Then rsPI.Update
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
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.IssueCreditNote(pCN,pWasProforma,pToProForma)", Array(pCN, pWasProforma, _
         pToProForma)
End Function

Public Function CancelCreditNote(pCN As a_CN)
    On Error GoTo errHandler
Dim strSQL As String
Dim rsCNL As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsPI As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim OpenResult As Integer

    Set rsCNL = New ADODB.Recordset
    strSQL = "SELECT CNL_P_ID,CNL_PI_ID,CNL_QTY FROM tCNL WHERE CNL_TR_ID = " & pCN.TRID
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rsCNL.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
    
    Do While Not rsCNL.eof
    
        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QtyOnHand,P_QtyTotalSold,P_QtyCopiesOnHand,p_QtyReserved,P_QtyExpectedBack FROM tProduct Where P_ID = '" & rsCNL.Fields("CNL_P_ID") & "'"
        rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        Set rsPI = New ADODB.Recordset
        strSQL = "SELECT PI_OnReserve,PI_FirstSoldDate,PI_ExpectedBack,PI_SoldDate FROM tPI Where PI_ID = " & rsCNL.Fields("CNL_PI_ID")
        rsPI.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QTYRESERVED,STP_QTYONHAND,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED FROM tStoreP Where STP_P_ID = '" & rsCNL.Fields("CNL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
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
                rsP.Fields("P_QtyOnHand") = FNN(rsP.Fields("P_QTYONHAND")) - FNN(rsCNL.Fields("CNL_QTY"))
                rsP.Fields("P_QtyTotalSold") = FNN(rsP.Fields("P_QtyTotalSold")) - FNN(rsCNL.Fields("CNL_QTY"))
            End If
        End If
        If Not rsSP.eof Then rsSP.Update
        If Not rsP.eof Then rsP.Update
        If Not rsPI.eof Then rsPI.Update
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
    oPC.COShort.execute "UPDATE tTR SET TR_STATUS = 6 WHERE TR_ID = " & pCN.TRID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CancelCreditNote(pCN)", pCN
End Function

Public Function IssueCO(pCO As a_CO)
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim OpenResult As Integer


    PostCO pCO.TRID, stISSUED

'''''-------------------------------
''''    OpenResult = oPC.OpenDBSHort
'''''-------------------------------
''''
''''    oPC.COShort.execute "UPDATE tTR SET TR_Status = 3,TR_CaptureDate = GetDate() WHERE TR_ID = " & pCO.TRID
''''    Set rs = New ADODB.Recordset
''''    Set rsP = New ADODB.Recordset
''''    Set rsSP = New ADODB.Recordset
''''    strSQL = "SELECT COL_P_ID,COL_QTY FROM tCOL WHERE COL_TR_ID = " & pCO.TRID
''''    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
''''
''''    Do While Not rs.EOF
''''
''''        strSQL = "SELECT STP_QTYRESERVED,STP_QTYONHAND,STP_QTYONBACKORDER FROM tStoreP Where STP_P_ID = '" & rs.Fields("COL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
''''        rsSP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
''''        If Not rsSP.EOF Then
''''            rsSP.Fields("STP_QTYONBACKORDER") = FNN(rsSP.Fields("STP_QTYONBACKORDER")) + FNN(rs.Fields("COL_QTY"))
''''            rsSP.Update
''''        End If
''''        rsSP.Close
''''        strSQL = "SELECT P_OBSOLETE,P_QTYONBACKORDER FROM tPRODUCT WHERE P_ID = '" & rs.Fields("COL_P_ID") & "'"
''''        rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
''''        rsP!P_QTYONBACKORDER = FNN(rsP!P_QTYONBACKORDER) + FNN(rs.Fields("COL_QTY"))
''''        rsP!P_OBSOLETE = 0
''''        If Not rsP.EOF Then rsP.Update
''''        rsP.Close
''''        rs.MoveNext
''''    Loop
''''    rs.Close
''''    Set rsP = Nothing
''''    Set rs = Nothing
''''    Set rsSP = Nothing
'''''---------------------------------------------------
''''    If OpenResult = 0 Then oPC.DisconnectDBShort
'''''---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.IssueCO(pCO)", pCO
End Function
Public Function CancelCO(pCO As a_CO)
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim OpenResult As Integer

    If Not pCO.CanCancel Then Exit Function
    
    PostCO pCO.TRID, stCANCELLED
''''    Set rs = New ADODB.Recordset
''''    strSQL = "SELECT COL_P_ID,COL_QTY FROM tCOL WHERE COL_TR_ID = " & pCO.TRID
''''
'''''-------------------------------
''''    OpenResult = oPC.OpenDBSHort
'''''-------------------------------
''''    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
''''
''''    oPC.COShort.BeginTrans
''''    Do While Not rs.EOF
''''        Set rsSP = New ADODB.Recordset
''''        strSQL = "SELECT STP_QtyOnBackorder,STP_QTYRESERVED,STP_QTYONHAND,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED FROM tStoreP Where STP_P_ID = '" & rs.Fields("COL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
''''        rsSP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
''''            rsSP.Fields("STP_QtyOnBackorder") = FNN(rsSP.Fields("STP_QtyOnBackorder")) - FNN(rs.Fields("COL_QTY"))
''''        If Not rsSP.EOF Then rsSP.Update
''''
''''        strSQL = "SELECT P_QTYONBACKORDER FROM tPRODUCT WHERE P_ID = '" & rs.Fields("COL_P_ID") & "'"
''''        Set rsP = New ADODB.Recordset
''''        rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
''''        rsP!P_QTYONBACKORDER = FNN(rsP!P_QTYONBACKORDER) - FNN(rs.Fields("COL_QTY"))
''''        If Not rsP.EOF Then rsP.Update
''''        rsP.Close
''''        Set rsP = Nothing
''''
''''
''''        rs.MoveNext
''''    Loop
''''    rs.Close
''''    Set rs = Nothing
''''    rsSP.Close
''''    Set rsSP = Nothing
''''    oPC.COShort.execute "UPDATE tTR SET TR_STATUS = 6 WHERE TR_ID = " & pCO.TRID
''''    oPC.COShort.CommitTrans
'''''---------------------------------------------------
''''    If OpenResult = 0 Then oPC.DisconnectDBShort
'''''---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CancelCO(pCO)", pCO
End Function
Public Function MarkCOLsActionedForProduct(pPID As String)
    On Error GoTo errHandler
Dim OpenResult As Integer
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "UPDATE tCOL SET COL_ACTIONTAKEN = 3 FROM tCOL JOIN tTR ON COL_TR_ID = TR_ID  WHERE COL_P_ID = '" & pPID & "' AND (COL_ACTIONTAKEN = 0 Or COL_ACTIONTAKEN = 2) AND TR_STATUS = 3"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.MarkCOLsActionedForProduct(pPID)", pPID
End Function
Public Sub MarkProductObsolete(pPID As String)
    On Error GoTo errHandler
Dim OpenResult As Integer
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "UPDATE tPRODUCT SET P_Obsolete = 1 WHERE P_ID = '" & pPID & "'"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.MarkProductObsolete(pPID)", pPID
End Sub

Public Function MarkCOL(pCOLID As Long) ', pPID As String, pQty As Long)
    On Error GoTo errHandler
Dim strSQL As String
Dim rsP As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim lngQty As Long
Dim strPID As String
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "UPDATE tCOL SET COL_ActionTaken = 3 WHERE COL_ID = " & pCOLID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.MarkCOL(pCOLID)", Array(pCOLID)
End Function

Public Function IssuePO(pPO As a_PO, pStatus As Long)
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsCOL As ADODB.Recordset
Dim OpenResult As Integer


    PostPO pPO.TRID, pStatus, pPO.ISForeignCurrency, pPO.Supplier.ID

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.IssuePO(pPO)", pPO
End Function

Public Function IssueAPP(pAPP As a_APP)
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsCOL As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "UPDATE tTR SET TR_Status = 3,TR_CaptureDate = GetDate() WHERE TR_ID = " & pAPP.TRID
    Set rs = New ADODB.Recordset
    strSQL = "SELECT APPL_P_ID,APPL_QTY,APPL_COL_ID FROM tAPPL WHERE APPL_TR_ID = " & pAPP.TRID
    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic

    Do While Not rs.eof

        Set rsSTP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnAppro,STP_QtyOnHand FROM tStoreP Where STP_P_ID = '" & rs.Fields("APPL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSTP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        If Not rsSTP.eof Then
            rsSTP.Fields("STP_QtyOnAppro") = FNN(rsSTP.Fields("STP_QtyOnAppro")) + FNN(rs.Fields("APPL_QTY"))
            rsSTP.Fields("STP_QtyOnHand") = FNN(rsSTP.Fields("STP_QtyOnHand")) - FNN(rs.Fields("APPL_QTY"))
            rsSTP.Update
        End If

        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QtyOnBackorder,P_LastApproToTPID,P_QtyOnHand,P_QtyOnAppro FROM tProduct Where P_ID = '" & rs.Fields("APPL_P_ID") & "'"
        rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        If Not rsP.eof Then
            rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) - FNN(rs.Fields("APPL_QTY"))
            rsP!P_QtyOnAppro = FNN(rsP!P_QtyOnAppro) + FNN(rs.Fields("APPL_QTY"))
            rsP!P_LastApproToTPID = pAPP.Customer.ID
            If oPC.CanGenerateApprosFromOrderFulfilment Then
                rsP!P_QtyOnBackorder = FNN(rsP!P_QtyOnBackorder) - FNN(rs.Fields("APPL_QTY"))
            End If
            rsP.Update
        End If
        
        Set rsCOL = New ADODB.Recordset
        strSQL = "SELECT * FROM tCOL Where COL_ID = " & rs.Fields("APPL_COL_ID")
        rsCOL.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        If Not rsCOL.eof Then
            rsCOL!COL_QtyDispatched = FNN(rsCOL!COL_QtyDispatched) + FNN(rs.Fields("APPL_QTY"))
            If FNN(rsCOL!COL_QtyDispatched) >= FNN(rsCOL!COL_Qty) Then
                rsCOL!COL_Fulfilled = "FUL"
            Else
                rsCOL!COL_Fulfilled = "OS"
            End If
            rsCOL.Update
        End If


        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    rsSTP.Close
    Set rsSTP = Nothing
    rsP.Close
    Set rsP = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.IssueAPP(pAPP)", pAPP
End Function
Public Function CancelAPP(pAPP As a_APP)
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
Dim OpenResult As Integer

    If Not pAPP.CanCancel Then Exit Function

    Set rs = New ADODB.Recordset
    strSQL = "SELECT APPL_P_ID,APPL_QTY FROM tAPPL WHERE APPL_TR_ID = " & pAPP.TRID
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic

    Do While Not rs.eof

        Set rsSTP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnAppro,STP_QtyOnHand FROM tStoreP Where STP_P_ID = '" & rs.Fields("APPL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSTP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
            rsSTP.Fields("STP_QtyOnAppro") = FNN(rsSTP.Fields("STP_QtyOnAppro")) - FNN(rs.Fields("APPL_QTY"))
            rsSTP.Fields("STP_QtyOnHand") = FNN(rsSTP.Fields("STP_QtyOnHand")) + FNN(rs.Fields("APPL_QTY"))
        If Not rsSTP.eof Then rsSTP.Update

        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QtyOnHand,P_QtyOnAppro FROM tProduct Where P_ID = '" & rs.Fields("APPL_P_ID") & "'"
        rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
            rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) + FNN(rs.Fields("APPL_QTY"))
            rsP!P_QtyOnAppro = FNN(rsP!P_QtyOnAppro) - FNN(rs.Fields("APPL_QTY"))
        If Not rsP.eof Then rsP.Update
        rs.MoveNext
    Loop
    oPC.COShort.execute "UPDATE tTR SET TR_Status = 6 WHERE TR_ID = " & pAPP.TRID
    rs.Close
    Set rs = Nothing
    rsSTP.Close
    Set rsSTP = Nothing
    rsP.Close
    Set rsP = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CancelAPP(pAPP)", pAPP
End Function


'Public Function IssueDEL(pDel As a_Delivery)
'    On Error GoTo errHandler
'Dim strSQL As String
'Dim strPos As String
'
'Dim rs As ADODB.Recordset
'Dim rsP As ADODB.Recordset
'Dim rsSP As ADODB.Recordset
'Dim rsPOL As ADODB.Recordset
'Dim rsCOL As ADODB.Recordset
'Dim rsPO As ADODB.Recordset
'Dim rs2 As ADODB.Recordset
'
'Dim bFulfilled As Boolean
'Dim lngtmpCost As Long
'Dim lngQtyReceived As Long
'Dim lngOldCost As Currency
'Dim lngNewCost As Currency
'Dim lngRoundTo As Long
''Must update    P_LastPriceDelivered,P_LastQtyDelivered,P_QtyOnHand,P_LastDateDelivered
''               POL_QtyReceivedSofar,POL_FulfilledStatus
''               STP_QtyOnHand,STP_QTYONORDER,STP_LastReceived,STP_FirstReceived
''Must generate COLAllocation rows
'Dim OpenResult As Integer
'Dim lngTotalDeliveryCost As Long
'Dim lngTotalQtyItems As Long
'Dim lngAverageExtraCost As Long
'Dim dblNewLocalPrice As Double
'Dim bMarkASFulfilled As Boolean
'
''Here we can prepare for ways of allocating extra cahrges e.g. freight to the base v
'    If oPC.GetProperty("ApplyExtraChargesToCostType") = "BY_QTY" Then
'        lngTotalDeliveryCost = pDel.BatchTotalExtras
'        lngTotalQtyItems = pDel.BatchQtyTotal
'        If lngTotalQtyItems > 0 Then
'            lngAverageExtraCost = lngTotalDeliveryCost / lngTotalQtyItems
'        Else
'            lngAverageExtraCost = 0
'        End If
'    Else
'            lngAverageExtraCost = 0
'    End If
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    oPC.COShort.execute "UPDATE tTR SET TR_Status = 4,TR_CaptureDate = GetDate() WHERE TR_ID = " & pDel.TRID
'
'
''----------------------------
'        'Distribute Extra Cost Over Stock Items and Calculate new average cost for each item   ----- Execute stored procedure to do this
'    DistributeExtraCostOverStockItems pDel.TRID, pDel.TRID
''------------------------------
'',DELL_MBCodeChanged
'     Set rs = New ADODB.Recordset
'    strSQL = "SELECT ISNULL(TR_TP_ID,0) TPID ,DEL_CURR_ID,DELL_ID,DELL_Discount,DELL_P_ID,DELL_QTYFIRM,DELL_QTYSS,DELL_PRICECHANGED,DELL_MBCodeChanged,DELL_MBCODE,DELL_PRODUCTTYPE_ID,DELL_QTYSHORT,DELL_FOREIGNPRICE,DELL_PRICE, " _
'            & " DELL_PRICESELL,DELL_POL_ID " _
'            & " FROM tDELL " _
'            & " LEFT JOIN tPOL ON DELL_POL_ID = POL_ID " _
'            & " LEFT JOIN tTR ON POL_TR_ID = TR_ID " _
'            & " LEFT JOIN tDEL ON DEL_ID = " & pDel.TRID _
'            & " WHERE DELL_TR_ID = " & pDel.TRID
'    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
'    Do While Not rs.eof
'        Set rsP = New ADODB.Recordset
'        strSQL = "SELECT P_ID,P_LastFCFactor,P_SupplierID,P_RRP,P_SP,P_Cost,P_QtyOnOrder,P_DefaultDeliveryDays,P_LastPriceDelivered,P_LastQtyDelivered, " _
'                  & "          P_QtyOnHand,P_LastDateDelivered,P_Obsolete,P_LastForeignPrice,P_LastFCID, P_ProductType_ID,P_USPRICE,P_UKPRICE,P_EUPRICE,P_MultibuyCode FROM tProduct " _
'                  & "                  Where P_ID = '" & rs.fields("DELL_P_ID") & "'"
'        rsP.open strSQL, oPC.COShort, adOpenDynamic, adLockPessimistic
'
'        If rsP.eof Then GoTo ENDOFLOOP
'        If rs!DEL_CURR_ID = oPC.Configuration.DefaultCurrencyID Then
''        'Determine average cost
'            If pDel.Supplier.VATable = True Then
'                lngNewCost = ((FNN(rs!DELL_Price) * ((100 - FNDBL(rs!DELL_Discount)) / 100)) + lngAverageExtraCost) / oPC.Configuration.DefaultCurrency.Divisor
'                lngNewCost = lngNewCost / ((100 + oPC.Configuration.VATRate) / 100)
'            Else
'                lngNewCost = ((FNN(rs!DELL_Price) * ((100 - FNDBL(rs!DELL_Discount)) / 100)) + lngAverageExtraCost) / oPC.Configuration.DefaultCurrency.Divisor
'            End If
'        Else
'            If pDel.Supplier.VATable = True Then
'                lngNewCost = ((FNN(rs!DELL_Price) * ((100 - FNDBL(rs!DELL_Discount)) / 100)) / oPC.Configuration.DefaultCurrency.Divisor)
'                lngNewCost = lngNewCost / ((100 + oPC.Configuration.VATRate) / 100)
'            Else
'                lngNewCost = ((FNN(rs!DELL_Price) * ((100 - FNDBL(rs!DELL_Discount)) / 100)) / oPC.Configuration.DefaultCurrency.Divisor)
'            End If
'            If FNN(rs!DEL_CURR_ID) = oPC.Configuration.Currencies.FindBySysname("USD").ID Then
'                rsP.fields("P_USPRICE") = FNDBL(rs!DELL_ForeignPrice)
'            Else
'                If FNN(rs!DEL_CURR_ID) = oPC.Configuration.Currencies.FindBySysname("GBP").ID Then
'                    rsP.fields("P_UKPRICE") = FNDBL(rs!DELL_ForeignPrice)
'                Else
'                    If FNN(rs!DEL_CURR_ID) = oPC.Configuration.Currencies.FindBySysname("EUR").ID Then
'                        rsP.fields("P_EUPRICE") = FNDBL(rs!DELL_ForeignPrice)
'                    End If
'                End If
'            End If
'            rsP!P_LastForeignPrice = rs!DELL_ForeignPrice
'            rsP!P_LastFCID = rs!DEL_CURR_ID
'            If FNN(rs!DEL_CURR_ID) > 0 Then rsP!P_LastFCFactor = oPC.Configuration.Currencies.FindCurrencyByID(FNN(rs!DEL_CURR_ID)).Factor
'        End If
'        If oPC.SetPricesInGRN Then
'            If rs!DEL_CURR_ID <> oPC.Configuration.DefaultCurrencyID Then
'                If oPC.UpdatePriceOnForeignDelivery Then
'                    dblNewLocalPrice = (FNDBL(rs!DELL_ForeignPrice) * oPC.Configuration.Currencies.FindCurrencyByID(FNN(rs!DEL_CURR_ID)).FactorInv) / oPC.Configuration.Currencies.FindCurrencyByID(FNN(rs!DEL_CURR_ID)).Divisor
'                    lngRoundTo = oPC.Configuration.RoundingRules.GetRoundTo(CLng(dblNewLocalPrice * 100))
'                    dblNewLocalPrice = (RoundUp(CLng(dblNewLocalPrice), lngRoundTo))
'                    If (rsP!P_SP <> dblNewLocalPrice * 100 * ((100 + oPC.Configuration.VATRate) / 100)) Or (FNS(rsP!P_MultibuyCode) <> FNS(rs!DELL_MBCode)) Then
'                        rsP!P_SP = dblNewLocalPrice * 100 * ((100 + oPC.Configuration.VATRate) / 100)
'                        rsP!P_RRP = dblNewLocalPrice * 100 * ((100 + oPC.Configuration.VATRate) / 100)
'                        rs!DELL_PriceChanged = True
'                        If rsP!P_MultibuyCode <> FNS(rs!DELL_MBCode) Then rs!DELL_MBCodeChanged = True
'                        rs.Update
'                    End If
'                End If
'            Else
'                If FNN(rsP!P_SP) <> FNN(rs!DELL_PRICESELL) Or (FNS(rsP!P_MultibuyCode) <> FNS(rs!DELL_MBCode)) Then
'                    If FNN(rsP!P_SP) <> FNN(rs!DELL_PRICESELL) Then
'                        InsertAuditRecord "SP", "Receiving", Format(FNN(rsP!P_SP) / 100, "###,##0.00"), Format(FNN(rs!DELL_PRICESELL) / 100, "###,##0.00"), gSTAFFID, rsP!P_ID, 0
'                    End If
'                    rsP!P_SP = FNN(rs!DELL_PRICESELL)
'                    rs!DELL_PriceChanged = True
'                    If rsP!P_MultibuyCode <> FNS(rs!DELL_MBCode) Then rs!DELL_MBCodeChanged = True
'                    rs.Update
'                End If
'                rsP!P_RRP = FNN(rs!DELL_Price)
'            End If
'        End If
'
''        rsP!P_LastQtyDelivered = FNN(rs!DELL_QTYFIRM) + FNN(rs!DELL_QTYSS)
''        rsP!P_QtyOnOrder = NonNegative_Long(FNN(rsP!P_QtyOnOrder) - FNN(rs!DELL_QTYFIRM) - FNN(rs!DELL_QTYSS))
''        rsP!P_LastDateDelivered = pDEL.DOCDate
''        rsP!P_LastPriceDelivered = lngNewCost * oPC.Configuration.DefaultCurrency.Divisor
'
'
'
'        If FNN(rs!DELL_ProductType_ID) > 0 Then
'            rsP!P_ProductType_ID = FNN(rs!DELL_ProductType_ID)
'        End If
'
'        If rsP!P_MultibuyCode <> FNS(rs!DELL_MBCode) Then
'            InsertAuditRecord "MB", "Receiving", FNS(rsP!P_MultibuyCode), FNS(rs!DELL_MBCode), gSTAFFID, rsP!P_ID, 0
'        End If
'        ''update MULTIBUY
'        rsP!P_MultibuyCode = FNS(rs!DELL_MBCode)
'
'        rsP!P_OBSOLETE = 0
'
'        If oPC.SetSupplierIDFROMDEL = True Then
'            If Not pDel.Supplier.DoNotOrderFrom Then
'                If pDel.Supplier.ID > 0 Then
'                    rsP!P_SUPPLIERID = pDel.Supplier.ID
'                End If
'            End If
'        End If
'
'        Set rsSP = New ADODB.Recordset
'        strSQL = "SELECT STP_QTYRESERVED,STP_FirstReceived,STP_QTYONHAND,STP_LastReceived,STP_QTYONORDER,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED FROM tStoreP Where STP_P_ID = '" & rs.fields("DELL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
'        rsSP.open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
'        If Not rsSP.eof Then
'            rsSP.fields("STP_QtyOnHand") = FNN(rsSP.fields("STP_QtyOnHand")) + (FNN(rs!DELL_QTYFIRM) + FNN(rs!DELL_QTYSS))
'            If rs.fields("DELL_POL_ID") > 0 Then
'                rsSP!STP_QtyOnOrder = NonNegative_Long(FNN(rsSP!STP_QtyOnOrder) - FNN(rs!DELL_QTYFIRM) - FNN(rs!DELL_QTYSS))
'            End If
'            rsSP.fields("STP_LastReceived") = pDel.DOCDate
'            If IsNull(rsSP.fields("STP_FirstReceived")) Then
'                rsSP.fields("STP_FirstReceived") = pDel.DOCDate
'            End If
'        End If
'
'        'markup the POL_Fulfilled value in tPOL and P_DefaultDeliveryDays in tProduct'''''''''''''''''
'        Set rsPOL = New ADODB.Recordset
'        strSQL = "SELECT TR_DATE,POL_QtyReceivedSofar,POL_Fulfilled,POL_TR_ID,POL_QTYSS,POL_QTYFirm FROM tPOL JOIN tTR ON POL_TR_ID = TR_ID Where POL_ID = " & rs.fields("DELL_POL_ID")
'        rsPOL.open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
'            If Not rsPOL.eof Then
'                rsPOL.fields("POL_QtyReceivedSofar") = FNN(rsPOL.fields("POL_QtyReceivedSofar")) + (FNN(rs!DELL_QTYFIRM) + FNN(rs!DELL_QTYSS))
'                bMarkASFulfilled = oPC.GetProperty("MarkPOLasFulfilledWhenShortSupplied") = "TRUE"
'                If (rsPOL.fields("POL_QTYSS") + FNN(rsPOL.fields("POL_QTYFirm"))) > FNN(rsPOL.fields("POL_QtyReceivedSofar")) + IIf(bMarkASFulfilled, FNN(rs.fields("DELL_QTYSHORT")), 0) Then
'                    rsPOL.fields("POL_FulFilled") = "OS"
'                Else
'                    rsPOL.fields("POL_FulFilled") = "FUL"
'                End If
'            End If
'
'
'        If Not rsPOL.eof Then
'            rsPOL.Update
'            rsP!P_DefaultDeliveryDays = DateDiff("d", rsPOL!TR_DATE, pDel.DOCDate)
'        End If
'        '''''''''''''''''''''''''''''
'
'        ''''probably obsolete
'        If Not rsSP.eof Then
'            rsSP.fields("STP_QTYONORDER") = NonNegative_Lng(FNN(rsSP.fields("STP_QTYONORDER")) - (FNN(rs!DELL_QTYFIRM) + FNN(rs!DELL_QTYSS)))
'            rsSP.Update
'        End If
'        '''''''''''''''''''''''''''''''''''''
'
'ENDOFLOOP:
'          rsP.fields("P_QtyOnHand") = FNN(rsP.fields("P_QtyOnHand")) + (FNN(rs!DELL_QTYFIRM) + FNN(rs!DELL_QTYSS))
'        If Not rsP.eof Then rsP.Update
'        rs.MoveNext
'    Loop
'
'    On Error Resume Next
'    rs2.Close
'    rs.Close
'    rsSP.Close
'    rsP.Close
'    rsPOL.Close
'    On Error GoTo errHandler
'
'    Set rsPOL = Nothing
'    Set rsSP = Nothing
'    Set rsP = Nothing
'    Set rs = Nothing
'    Set rs2 = Nothing
'         strSQL = "UPDATE tPRODUCT SET " _
'                & " P_QtyOnOrder = TOTOS, " _
'                & " P_LastDateDelivered = DateReceived, " _
'                & " P_LastPriceDelivered = PRICE, " _
'                & " P_LastQtyDelivered = QtyFirm + QtySS " _
'                & " FROM tPRODUCT JOIN vPOLs_Aggr_ByDEL ON PID = P_ID WHERE DELID =  " & pDel.TRID
'
'        oPC.COShort.execute strSQL
'    UpdateFulfilledStatus pDel.TRID
'
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_StockManager.IssueDEL(pDEL)", pDel
'End Function
Public Function IssueDEL(pDel As a_Delivery)
49140     On Error GoTo errHandler
      Dim cmd As ADODB.Command
      Dim par As ADODB.Parameter
      Dim OpenResult As Integer
      '-------------------------------
49150     OpenResult = oPC.OpenDBSHort
      '-------------------------------
49160     Set cmd = New ADODB.Command
49170     cmd.CommandText = "ISSUEDEL"
49180     cmd.commandType = adCmdStoredProc
49190     cmd.CommandTimeout = 120
          
49200     Set par = cmd.CreateParameter("@DELID", adInteger, adParamInput, , pDel.TRID)
49210     cmd.Parameters.Append par
49220     Set par = Nothing
          
49230     cmd.ActiveConnection = oPC.COShort
49240     cmd.execute
          
          
49250     Set cmd = Nothing
      '---------------------------------------------------
49260     If OpenResult = 0 Then oPC.DisconnectDBShort
      '---------------------------------------------------

49270     Exit Function
49280     Exit Function
errHandler:
49290     If ErrMustStop Then Debug.Assert False: Resume
49300     ErrorIn "z_StockManager.IssueDEL(pDEL)", pDel
End Function
Public Sub UpdateFulfilledStatus(DELID As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "UpdatePOLFulfilledStatus"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@DELID", adInteger, adParamInput, , DELID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.UpdateFulfilledStatus(DELID)", DELID
End Sub
Public Function DistributeExtraCostOverStockItems(SourceDOcID As Long, TargetDocID As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    cmd.CommandText = "DistributeExtraCostOverStockItems"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@SourceDocID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = SourceDOcID
    Set par = cmd.CreateParameter("@TargetDocID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = TargetDocID
    Set par = cmd.CreateParameter("@STYLE", adVarChar, , 10, adParamInput)
    cmd.Parameters.Append par
    par.Value = "A"
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.DistributeExtraCostOverStockItems(SourceDOcID,TargetDocID)", _
         Array(SourceDOcID, TargetDocID)
End Function


Public Function IssueAPPR(pAPPR As a_APPR)
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsAPPL As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsAPP As ADODB.Recordset
Dim bFulfilled As Boolean
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "UPDATE tTR SET TR_Status = 3,TR_CaptureDate = GetDate() WHERE TR_ID = " & pAPPR.TRID
    Set rs = New ADODB.Recordset
    strSQL = "SELECT APPRL_APPL_ID,APPRL_Qty FROM tAPPRL WHERE APPRL_TR_ID = " & pAPPR.TRID
    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic

    Do While Not rs.eof
        
        Set rsAPPL = New ADODB.Recordset
        strSQL = "SELECT APPL_TR_ID,APPL_P_ID,APPL_QTY,APPL_QTYReturned,APPL_Fulfilled FROM tAPPL WHERE APPL_ID = " & rs!APPRL_APPL_ID
        rsAPPL.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
            rsAPPL!APPL_QtyReturned = FNN(rsAPPL!APPL_QtyReturned) + FNN(rs!APPRL_Qty)
            If FNN(rsAPPL!APPL_Qty) <= FNN(rsAPPL!APPL_QtyReturned) Then
                rsAPPL!APPL_Fulfilled = "FUL"
            End If
        If Not rsAPPL.eof Then rsAPPL.Update
        
        
        strSQL = "Select APPL_QTY,APPL_QTYReturned from tAPPL WHERE APPL_TR_ID = " & rsAPPL!APPL_TR_ID
        Set rsAPP = New ADODB.Recordset
        rsAPP.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
        bFulfilled = True
        Do While Not rsAPP.eof
            If FNN(rsAPP!APPL_Qty) > FNN(rsAPP!APPL_QtyReturned) Then
                bFulfilled = False
                Exit Do
            End If
            rsAPP.MoveNext
        Loop
        If bFulfilled Then
            oPC.COShort.execute "UPDATE tTR SET TR_STATUS = 4 WHERE TR_ID = " & rsAPPL!APPL_TR_ID
        End If
        rsAPP.Close
        Set rsAPP = Nothing

        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_OBSOLETE,P_QtyOnHand,P_QtyOnAppro FROM tProduct Where P_ID = '" & rsAPPL.Fields("APPL_P_ID") & "'"
        rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        If Not rsP.eof Then
            rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) + FNN(rs!APPRL_Qty)
            rsP!P_QtyOnAppro = FNN(rsP!P_QtyOnAppro) - FNN(rs!APPRL_Qty)
            rsP!P_OBSOLETE = 0
            rsP.Update
        End If
        rsP.Close
        Set rsP = Nothing


        rs.MoveNext
        rsAPPL.Close
        Set rsAPPL = Nothing
        
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.IssueAPPR(pAPPR)", pAPPR
End Function
Public Function CancelAPPR(pAPPR As a_APPR)
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsAPPL As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsAPP As ADODB.Recordset
Dim bFulfilled As Boolean
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    strSQL = "SELECT APPRL_APPL_ID,APPRL_Qty FROM tAPPRL WHERE APPRL_TR_ID = " & pAPPR.TRID
    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic

    Do While Not rs.eof
        Set rsAPPL = New ADODB.Recordset
        strSQL = "SELECT APPL_TR_ID,APPL_P_ID,APPL_QTY,APPL_QTYReturned,APPL_Fulfilled FROM tAPPL WHERE APPL_ID = " & rs!APPRL_APPL_ID
        rsAPPL.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
            rsAPPL!APPL_QtyReturned = FNN(rsAPPL!APPL_QtyReturned) - FNN(rs!APPRL_Qty)
            rsAPPL!APPL_Fulfilled = "OS"
        If Not rsAPPL.eof Then rsAPPL.Update
        oPC.COShort.execute "UPDATE tTR SET TR_STATUS = 3 WHERE TR_ID = " & rsAPPL!APPL_TR_ID 'It cant be COMPLETE but has to be ISSUED only
        
        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnHand,STP_QtyOnAppro FROM tStoreP Where STP_P_ID = '" & rsAPPL.Fields("APPL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        If Not rsSP.eof Then
            rsSP.Fields("STP_QtyOnHand") = FNN(rsSP.Fields("STP_QtyOnHand")) - FNN(rs.Fields("APPRL_Qty"))
            rsSP.Fields("STP_QtyOnAppro") = FNN(rsSP.Fields("STP_QtyOnAppro")) + FNN(rs.Fields("APPRL_Qty"))
            rsSP.Update
        End If
        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QtyOnHand,P_QtyOnAppro FROM tProduct Where P_ID = '" & rsAPPL.Fields("APPL_P_ID") & "'"
        rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        If Not rsP.eof Then
            rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) - FNN(rs!APPRL_Qty)
            rsP!P_QtyOnAppro = FNN(rsP!P_QtyOnAppro) + FNN(rs!APPRL_Qty)
            rsP.Update
        End If
        rs.MoveNext
    Loop
    oPC.COShort.execute "UPDATE tTR SET TR_STATUS = 6 WHERE TR_ID = " & pAPPR.TRID 'It cant be COMPLETE but has to be ISSUED only
    rs.Close
    Set rs = Nothing
    rsSP.Close
    Set rsSP = Nothing
    rsP.Close
    Set rsP = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CancelAPPR(pAPPR)", pAPPR
End Function

Public Function CancelPO(pPO As a_PO)
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsCOL As ADODB.Recordset
Dim rsPOL As ADODB.Recordset
Dim rsMAXPOL As ADODB.Recordset
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer

    PostPO pPO.TRID, stCANCELLED, pPO.ISForeignCurrency, pPO.Supplier.ID


    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CancelPO(pPO)", pPO
End Function
'Public Function CancelPOL(pPOLID As Long, pPID As String, pQty As Long)
'    On Error GoTo ErrHandler
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'Dim strSQL As String
'Dim rsP As ADODB.Recordset
'Dim rsSTP As ADODB.Recordset
'    oPC.COShort.BeginTrans
'    oPC.COShort.execute "UPDATE tPOL SET POL_Fulfilled = 'CAN' WHERE POL_ID = " & pPOLID
'
'    Set rsP = New ADODB.Recordset
'    strSQL = "SELECT P_QTYONORDER FROM tPRODUCT WHERE P_ID = '" & pPID & "'"
'    rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
'        rsP!P_QtyOnOrder = NonNegative_Lng(FNN(rsP!P_QtyOnOrder) - pQty)
'    If Not rsP.EOF Then rsP.Update
'    rsP.Close
'    Set rsP = Nothing
'
'        Set cmd = New ADODB.Command
'        cmd.CommandText = "sp_ResetLastPO"
'        cmd.CommandType = adCmdStoredProc
'        Set par = cmd.CreateParameter("PID", adGUID, adParamInput)
'        cmd.Parameters.Append par
'        par.Value = pPID
'        cmd.ActiveConnection = oPC.COShort
'        cmd.execute
'        Set cmd = Nothing
'
'    Set rsSTP = New ADODB.Recordset
'    strSQL = "SELECT STP_QTYONORDER FROM tStoreP WHERE STP_P_ID = '" & pPID & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
'    rsSTP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
'        rsSTP!STP_QtyOnOrder = NonNegative_Lng(FNN(rsSTP!STP_QtyOnOrder) - pQty)
'    If Not rsSTP.EOF Then rsSTP.Update
'    rsSTP.Close
'    Set rsSTP = Nothing
'    oPC.COShort.CommitTrans
'    Exit Function
'
'ErrHandler:
'    RlsObjs rsSTP, rsP
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_StockManager.CancelPOL(pPOLID,pPID,pQty)", Array(pPOLID, pPID, pQty), EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
'End Function
'Public Function MarkPOLFulfilled(pPOLID As Long, pPID As String, pQty As Long)
'    On Error GoTo ErrHandler
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'Dim strSQL As String
'Dim rsP As ADODB.Recordset
'Dim rsSTP As ADODB.Recordset
'    oPC.COShort.BeginTrans
'    oPC.COShort.execute "UPDATE tPOL SET POL_Fulfilled = 'FUL' WHERE POL_ID = " & pPOLID
'
'    Set rsP = New ADODB.Recordset
'    strSQL = "SELECT P_QTYONORDER FROM tPRODUCT WHERE P_ID = '" & pPID & "'"
'    rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
'    rsP!P_QtyOnOrder = NonNegative_Lng(FNN(rsP!P_QtyOnOrder) - pQty)
'    If Not rsP.EOF Then rsP.Update
'    rsP.Close
'    Set rsP = Nothing
'
'
'    Set rsSTP = New ADODB.Recordset
'    strSQL = "SELECT STP_QTYONORDER FROM tStoreP WHERE STP_P_ID = '" & pPID & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
'    rsSTP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
'    rsSTP!STP_QtyOnOrder = NonNegative_Lng(FNN(rsSTP!STP_QtyOnOrder) - pQty)
'    If Not rsSTP.EOF Then rsSTP.Update
'    rsSTP.Close
'    Set rsSTP = Nothing
'    oPC.COShort.CommitTrans
'    Exit Function
'
'ErrHandler:
'    RlsObjs rsSTP, rsP
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_StockManager.MarkPOLFulfilled(pPOLID,pPID,pQty)", Array(pPOLID, pPID, pQty), EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
'End Function
Public Function MarkPOLStatus(pPOLID As Long, pStatus As String) As Long
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_MARKPOLSTATUS"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@POLID", adInteger, adParamInput, , pPOLID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@STATUS", adVarChar, adParamInput, 3, pStatus)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CODE", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    MarkPOLStatus = cmd.Parameters(2)
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.MarkPOLStatus(pPOLID,pStatus)", Array(pPOLID, pStatus)
End Function

Public Function MarkCOLStatus(pCOLID As Long, pStatus As String) As Long
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_MARKCOLSTATUS"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@COLID", adInteger, adParamInput, , pCOLID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@STATUS", adVarChar, adParamInput, 3, pStatus)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CODE", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    MarkCOLStatus = cmd.Parameters(2)
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.MarkCOLStatus(pCOLID,pStatus)", Array(pCOLID, pStatus)
End Function

Public Function AdjustStock(pPID As String, pCount As Long, pSTAFFID As Long, pNote As String, Optional pDateTime As Date) As Long
    On Error GoTo errHandler
Dim rsP As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
Dim strSQL As String
Dim lngCalc As Long
Dim lngAdj As Long
Dim OpenResult As Integer
Dim zSQL As New z_SQL
Dim intCLST As Long
'Dim intLDC As Integer
Dim intLPD As Long
Dim intc As Long
    AdjustStock = 0
    
    If pDateTime = CDate(0) Then
        pDateTime = Now()
    End If
 '   zSQL.SwitchTriggers "disable"   taken out because it looks like it could cause lock ups -as experienced by PLE
    Set rsP = New ADODB.Recordset
    strSQL = "SELECT P_QtyOnHand,P_DateLastStockTake,P_QtyLastStockTake,Isnull(P_CostLastStocktake,0) as P_CostLastStocktake, Isnull(P_LastPriceDelivered,0) as P_LastPriceDelivered FROM tProduct WHERE P_ID = '" & pPID & "'"
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        lngCalc = FNN(rsP!P_QtyOnHand)
        intCLST = rsP!P_CostLastStocktake
        intLPD = rsP!P_LastPriceDelivered
        If intCLST = 0 Then
            intc = intLPD
        Else
            intc = intCLST
        End If
    If lngCalc <> pCount Then
        rsP!P_QtyOnHand = pCount
        rsP.Update
        
        lngAdj = pCount - lngCalc
        oPC.COShort.execute "INSERT INTO tADJ (ADJ_PID,ADJ_Adjustment,ADJ_Counted,ADJ_Date,ADJ_StaffID,ADJ_NOTE) VALUES ('" & pPID & "'," & lngAdj & "," & pCount & ", '" & ReverseDateTime(pDateTime) & "'," & pSTAFFID & ",'" & Replace(pNote, "'", "''") & "')"
        
        Set rsSTP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnHand,STP_DateLastStockTake,STP_QtyLastStockTake FROM tStoreP WHERE STP_P_ID = '" & pPID & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSTP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        If Not rsSTP.eof Then
            rsSTP!STP_QTYONHAND = pCount
            rsSTP.Update
        End If
        rsSTP.Close
    End If
   
    rsP.Close
    Set rsP = Nothing
    Set rsSTP = Nothing
 '   zSQL.SwitchTriggers "Enable"

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:

    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.AdjustStock(pPID,pCount,pStaffID,pNote,pDateTime)", Array(pPID, pCount, _
         pSTAFFID, pNote, pDateTime)
         
End Function
Public Function CANCELDELL(pDELLID As Long)
    On Error GoTo errHandler

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CANCELDELL(pDELLID)", pDELLID
End Function
Public Function CancelDEL(pDel As a_Delivery)
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsDELL As ADODB.Recordset
Dim rsPOL As ADODB.Recordset
Dim rsCOL As ADODB.Recordset
Dim rsPO As ADODB.Recordset
Dim rsMAXDELL As ADODB.Recordset
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim bFulfilled As Boolean
Dim OpenResult As Integer

'Must update    P_LastPriceDelivered,P_LastQtyDelivered,P_QtyOnHand,P_LastDateDelivered
'               POL_QtyReceivedSofar,POL_FulfilledStatus
'               STP_QtyOnHand,STP_QTYONORDER,STP_LastReceived,STP_FirstReceived
'MsgBox "dont forget STP_LastReceived,STP_FirstReceived"
    Set rs = New ADODB.Recordset
    strSQL = "SELECT DELL_ID,DELL_P_ID,DELL_QTYFIRM,DELL_QTYSS,DELL_PRICE,DELL_POL_ID FROM tDELL WHERE DELL_TR_ID = " & pDel.TRID
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
    oPC.COShort.BeginTrans
    Do While Not rs.eof


        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QTYRESERVED,STP_FirstReceived,STP_QTYONHAND,STP_LastReceived,STP_QTYONORDER,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED FROM tStoreP Where STP_P_ID = '" & rs.Fields("DELL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        If Not rsSP.eof Then
            rsSP.Fields("STP_QtyOnHand") = FNN(rsSP.Fields("STP_QtyOnHand")) - (rs!DELL_QTYFIRM + rs!DELL_QTYSS)
            If FNN(rs.Fields("DELL_POL_ID")) > 0 Then
                rsSP.Fields("STP_QTYONORDER") = FNN(rsSP.Fields("STP_QTYONORDER")) + (rs!DELL_QTYFIRM + rs!DELL_QTYSS)
            End If
            rsSP.Fields("STP_LastReceived") = CDate(0)
            If rsSP.Fields("STP_FirstReceived") = pDel.DOCDate Then
                rsSP.Fields("STP_FirstReceived") = CDate(0)
            End If
        End If
        If Not rsSP.eof Then rsSP.Update
        rsSP.Close
        
        Set cmd = New ADODB.Command
        cmd.CommandText = "sp_ResetLastDelivery"
        cmd.commandType = adCmdStoredProc
        Set par = cmd.CreateParameter("PID", adGUID, adParamInput)
        cmd.Parameters.Append par
        par.Value = rs.Fields("DELL_P_ID")
        cmd.ActiveConnection = oPC.COShort
        cmd.execute
        Set cmd = Nothing
'''''''''''''''''''''''''''''''''''
'        Set rsMAXDELL = New ADODB.Recordset
'        strSQL = "SELECT MAX(DELL_ID) as MAXID FROM ActiveDELLs WHERE DELL_P_ID = '" & rs!DELL_P_ID & "'"
'        rsMAXDELL.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
'        If Not rsMAXDELL.EOF Then
'            Set rsDELL = New ADODB.Recordset
'            rsDELL.Open "SELECT TR_Date,DELL_PRICE,DELL_QtyFirm,DELL_QtySS,TR_Date FROM tDELL JOIN tTR ON DELL_TR_ID = TR_ID WHERE DELL_ID = " & rsMAXDELL.Fields("MAXID"), oPC.COShort, adOpenKeyset, adLockPessimistic
'        End If
        Set rsP = New ADODB.Recordset
        rsP.Open "SELECT P_QtyOnHand,P_LastQtyDelivered,P_LastDateDelivered,P_LastPriceDelivered FROM tProduct Where P_ID = '" & rs.Fields("DELL_P_ID") & "'", oPC.COShort, adOpenKeyset, adLockPessimistic
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
        If Not rsP.eof Then rsP.Update
        rsP.Close
        Set rsP = Nothing
        strSQL = "UPDATE tPRODUCT SET " _
                & " P_QtyOnOrder = TOTOS, " _
                & " P_LastDateDelivered = DateReceived, " _
                & " P_LastPriceDelivered = PRICE, " _
                & " P_LastQtyDelivered = QtyFirm + QtySS " _
                & " FROM tPRODUCT JOIN vPOLs_Aggr_ByDEL ON PID = P_ID WHERE DELID =  " & pDel.TRID
        
        oPC.COShort.execute strSQL
''''''''''''''''''''''''''''''
        
        Set rsPOL = New ADODB.Recordset
        strSQL = "SELECT POL_QtyReceivedSofar,POL_Fulfilled,POL_TR_ID,POL_QTYSS,POL_QTYFirm FROM tPOL Where POL_ID = " & rs.Fields("DELL_POL_ID")
        rsPOL.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        If Not rsPOL.eof Then
            rsPOL.Fields("POL_QtyReceivedSofar") = FNN(rsPOL.Fields("POL_QtyReceivedSofar")) - (FNN(rs!DELL_QTYFIRM) + FNN(rs!DELL_QTYSS))
            rsPOL.Fields("POL_FulFilled") = "OS"
            rsPOL.Update
        End If

        oPC.COShort.execute "UPDATE tTR SET TR_STATUS = 6 WHERE TR_ID = " & pDel.TRID
        rs.MoveNext
    Loop
    Set rsPOL = Nothing
    Set rsSP = Nothing
    Set rsP = Nothing
    Set rs = Nothing
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CancelDEL(pDEL)", pDel
End Function



Public Function IssueTF(pTF As a_TF)
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------


    oPC.COShort.execute "UPDATE tTR SET TR_Status = 4,TR_CaptureDate = GetDate() WHERE TR_ID = " & pTF.TRID

    Set rs = New ADODB.Recordset
    strSQL = "SELECT TFRL_P_ID,TFRL_QTY,TFRL_PRICE FROM tTFRL WHERE TFRL_TR_ID = " & pTF.TRID
    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic

    Do While Not rs.eof

        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_OBSOLETE,P_QtyOnHand,P_QtyOnAppro,P_SP,P_RRP FROM tProduct Where P_ID = '" & rs.Fields("TFRL_P_ID") & "'"
        rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
            If pTF.InOut = "OUT" Then
                rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) - FNN(rs.Fields("TFRL_QTY"))
            Else
                rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) + FNN(rs.Fields("TFRL_QTY"))
                rsP!P_SP = FNN(rs.Fields("TFRL_PRICE"))
                rsP!P_RRP = rsP!P_SP
                rsP!P_OBSOLETE = 0
            End If
        If Not rsP.eof Then rsP.Update
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    'rsSTP.Close
    Set rsSTP = Nothing
    rsP.Close
    Set rsP = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
'H:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.IssueTF(pTF)", pTF
End Function
Public Function CancelTF(pTF As a_TF)
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set rs = New ADODB.Recordset
    strSQL = "SELECT TFRL_P_ID,TFRL_QTY FROM tTFRL WHERE TFRL_TR_ID = " & pTF.TRID
    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
    
    Do While Not rs.eof

        Set rsSTP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnAppro,STP_QtyOnHand FROM tStoreP Where STP_P_ID = '" & rs.Fields("TFRL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSTP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        If Not rsSTP.eof Then
            If pTF.InOut = "OUT" Then
                rsSTP.Fields("STP_QtyOnHand") = FNN(rsSTP.Fields("STP_QtyOnHand")) + FNN(rs.Fields("TFRL_QTY"))
            Else
                rsSTP.Fields("STP_QtyOnHand") = FNN(rsSTP.Fields("STP_QtyOnHand")) - FNN(rs.Fields("TFRL_QTY"))
            End If
            rsSTP.Update
        End If

        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QtyOnHand,P_QtyOnAppro FROM tProduct Where P_ID = '" & rs.Fields("TFRL_P_ID") & "'"
        rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
            If pTF.InOut = "OUT" Then
                rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) + FNN(rs.Fields("TFRL_QTY"))
            Else
                rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) - FNN(rs.Fields("TFRL_QTY"))
            End If
        If Not rsP.eof Then rsP.Update
        rs.MoveNext
    Loop
    oPC.COShort.execute "UPDATE tTR SET TR_Status = 6 WHERE TR_ID = " & pTF.TRID
    rs.Close
    Set rs = Nothing
    rsSTP.Close
    Set rsSTP = Nothing
    rsP.Close
    Set rsP = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CancelTF(pTF)", pTF
End Function

Public Function GenerateReturnsPerPubNew(pTPID As Long, pMinDate As Date, pMaxDate As Date, pSTAFFID As Long, Optional bSupplementary As Boolean) As Integer
    On Error GoTo errHandler
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
Dim iCount As Long
Dim strDefaultCo As String
Dim OpenResult As Integer

    'Get titles that have been received SeeSafe on supplier invoices prior to or on the pDATE parameter.
    'The Titles should have QtyOnHand > 0
    If bSupplementary Then
'    strSQL = "SELECT SUM(dbo.tRL.RL_QtyApproved - dbo.tRL.RL_QtyReturned) AS QTYSS, dbo.tProduct.P_ID, dbo.tR.R_ApprovalRef, dbo.tProduct.P_QtyOnHand, " _
'            & " dbo.tProduct.P_RRP , dbo.tProduct.P_Title " _
'            & " FROM dbo.tTR INNER JOIN " _
'            & " dbo.tR INNER JOIN " _
'            & " dbo.tRL ON dbo.tR.R_ID = dbo.tRL.RL_TR_ID INNER JOIN " _
'            & " dbo.tProduct ON dbo.tRL.RL_P_ID = dbo.tProduct.P_ID ON dbo.tTR.TR_ID = dbo.tR.R_ID " _
'            & " Where (dbo.tRL.RL_QtyApproved > dbo.tRL.RL_QtyReturned) AND (isnull(tR.R_Type,'S') <> 'E') AND tTR.TR_TP_ID = " & pTPID _
'            & " GROUP BY  dbo.tProduct.P_ID, " _
'            & " dbo.tR.R_ApprovalRef, dbo.tProduct.P_QtyOnHand, dbo.tProduct.P_RRP, dbo.tProduct.P_Title" _
'            & " Having (dbo.tProduct.P_QtyOnHand > 0)"
            
            
    strSQL = "SELECT DISTINCT ISNULL(dbo.tDELL.DELL_QtySS, 0) - ISNULL(Qty, 0) AS QTYSS, dbo.tProduct.P_ID, " _
                    & " dbo.tProduct.P_QtyOnHand,dbo.tProduct.P_RRP, dbo.tDELL.DELL_Discount AS Discount, " _
                    & " dbo.tDEL.DEL_SupplierInvoiceRef, dbo.tDEL.DEL_SupplierInvoiceDate,dbo.tDELL.DELL_ID " _
            & " FROM         dbo.tTR AS tTR_1 INNER JOIN " _
                    & " dbo.tDELL ON tTR_1.TR_ID = dbo.tDELL.DELL_TR_ID INNER JOIN " _
                    & " dbo.tDEL ON dbo.tDELL.DELL_TR_ID = dbo.tDEL.DEL_ID INNER JOIN " _
                    & " dbo.tProduct ON dbo.tProduct.P_ID = dbo.tDELL.DELL_P_ID LEFT OUTER JOIN " _
                    & " dbo.vReturnsPerDELL_Appr ON dbo.tDELL.DELL_ID = dbo.vReturnsPerDELL_Appr.DELLID " _
            & " WHERE     (tTR_1.TR_Status IN (3, 4)) AND (dbo.tDEL.DEL_SupplierInvoiceDate >= CONVERT(DATETIME,   '" & ReverseDate(pMaxDate) & "', 102)) " _
                    & " AND (dbo.tDEL.DEL_SupplierInvoiceDate <= CONVERT(DATETIME,   '" & ReverseDate(pMinDate) & "', 102)) " _
                    & " AND (tTR_1.TR_TP_ID =   " & pTPID & " ) " _
                    & " AND (dbo.tProduct.P_QtyOnHand > 0) " _
                    & " AND (dbo.tDELL.DELL_QTYSS - ISNULL(dbo.vReturnsPerDELL_Appr.QTY, 0) > 0)"
            
            
    Else

    
    strSQL = "SELECT DISTINCT ISNULL(dbo.tDELL.DELL_QtySS, 0) - ISNULL(Qty, 0) AS QTYSS, dbo.tProduct.P_ID, " _
                    & " dbo.tProduct.P_QtyOnHand,dbo.tProduct.P_RRP, dbo.tDELL.DELL_Discount AS Discount, " _
                    & " dbo.tDEL.DEL_SupplierInvoiceRef, dbo.tDEL.DEL_SupplierInvoiceDate,dbo.tDELL.DELL_ID " _
            & " FROM         dbo.tTR AS tTR_1 INNER JOIN " _
                    & " dbo.tDELL ON tTR_1.TR_ID = dbo.tDELL.DELL_TR_ID INNER JOIN " _
                    & " dbo.tDEL ON dbo.tDELL.DELL_TR_ID = dbo.tDEL.DEL_ID INNER JOIN " _
                    & " dbo.tProduct ON dbo.tProduct.P_ID = dbo.tDELL.DELL_P_ID LEFT OUTER JOIN " _
                    & " dbo.vReturnsPerDELL ON dbo.tDELL.DELL_ID = dbo.vReturnsPerDELL.DELLID " _
            & " WHERE     (tTR_1.TR_Status IN (3, 4)) AND (dbo.tDEL.DEL_SupplierInvoiceDate >= CONVERT(DATETIME,   '" & ReverseDate(pMaxDate) & "', 102)) " _
                    & " AND (dbo.tDEL.DEL_SupplierInvoiceDate <= CONVERT(DATETIME,   '" & ReverseDate(pMinDate) & "', 102)) " _
                    & " AND (tTR_1.TR_TP_ID =   " & pTPID & " ) " _
                    & " AND (dbo.tProduct.P_QtyOnHand > 0) " _
                    & " AND (dbo.tDELL.DELL_QTYSS - ISNULL(dbo.vReturnsPerDELL.QTY, 0) > 0)"
    
    
 
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
    
    If rs.eof Then
        GenerateReturnsPerPubNew = 0
        Exit Function
    Else
        GenerateReturnsPerPubNew = 1
    End If
    RaiseEvent eMax(rs.RecordCount)
    
    oPC.COShort.BeginTrans
    
    Set rsTR = New ADODB.Recordset
  '  rsTR.CursorLocation = adUseClient
    rsTR.Open "Select * FROM tTR WHERE TR_ID = 0", oPC.COShort, adOpenKeyset, adLockOptimistic
    rsTR.AddNew
        rsTR!TR_DATE = ReverseDate(Date)
        rsTR!TR_TP_ID = pTPID
        rsTR!TR_CaptureDate = ReverseDateTime(Now)
        strDefaultCo = oPC.Configuration.DefaultCompany.CompanyCode
        rsTR!TR_CODE = strDefaultCo & "R" & oCode.GetNextCode(enReturn, "GETANDWRITE")
        rsTR!TR_Type = 11
        If bSupplementary Then
            rsTR!TR_Status = 3
        Else
            rsTR!TR_Status = 2
        End If
        rsTR!TR_StaffID = pSTAFFID
        rsTR!TR_COMP_ID = oPC.Configuration.DefaultCOMPID
    rsTR.Update
    rsTR.Bookmark = rsTR.Bookmark
    lngTRID = rsTR("TR_ID")
    
    Set rsR = New ADODB.Recordset
    rsR.Open "Select * FROM tR WHERE R_ID = 0", oPC.COShort, adOpenKeyset, adLockOptimistic
    rsR.AddNew
        rsR!R_ID = lngTRID
        rsR!R_Status = 0
        If bSupplementary Then
            rsR!R_Type = "E"
        Else
            rsR!R_Type = "P"
        End If
    rsR.Update
    
    Set rsRL = New ADODB.Recordset
    rsRL.Open "Select * FROM tRL WHERE RL_ID = 0", oPC.COShort, adOpenKeyset, adLockOptimistic
    i = 0
    Do While Not rs.eof
        rsRL.AddNew
            rsRL!RL_P_ID = rs!P_ID
            rsRL!RL_TR_ID = lngTRID
            rsRL!RL_Price = rs!P_RRP
            rsRL!RL_DELL_ID = rs!DELL_ID
            rsRL!RL_QtySystem = rs!P_QtyOnHand
            rsRL!RL_SiREF = FNS(rs!DEL_SUpplierInvoiceRef)
            rsRL!RL_QtyRequested = GetMin(CLng(rs!P_QtyOnHand), CLng(rs!QtySS))
            If bSupplementary Then
                rsRL!RL_QtyApproved = GetMin(CLng(rs!P_QtyOnHand), CLng(rs!QtySS))
'                rsRL!RL_SiREF = FNS(rs!R_ApprovalRef)
           End If
        rsRL.Update
        rs.MoveNext
        i = i + 1
        RaiseEvent eProgress(i)
    Loop
    
    If Not bSupplementary Then
'        strSQL = "SELECT DELL_ID,DELL_QTYSS,RL_ID,DEL_SupplierInvoiceRef,DEL_SupplierInvoiceDate,DELL_Price,DELL_Discount " _
'                & " FROM tDEL " _
'                & " INNER JOIN tDELL on DELL_TR_ID = DEL_ID " _
'                & " INNER JOIN tRL ON RL_P_ID = DELL_P_ID " _
'                & " INNER JOIN tTR ON DELL_TR_ID = TR_ID " _
'                & " WHERE RL_TR_ID = " & lngTRID & " AND (TR_STATUS = 3 or TR_Status = 4) " _
'                & " AND DEL_SupplierInvoiceDate >= {d '" & ReverseDate(pMaxDate) & "'} " _
'                & " AND DEL_SupplierInvoiceDate <= {d '" & ReverseDate(pMinDate) & "'} " _
'                & " ORDER BY RL_ID,DEL_SupplierInvoiceDate"
                
'        strSQL = "SELECT  dbo.tDELL.DELL_ID, dbo.tDELL.DELL_QtySS - ISNULL(dbo.tRL.RL_QtyReturned, 0) as DELL_QTYSS, dbo.tRL.RL_ID, dbo.tDEL.DEL_SupplierInvoiceRef, " _
'            & " dbo.tDEL.DEL_SupplierInvoiceDate , dbo.tDELL.DELL_Price, dbo.tDELL.DELL_Discount " _
'            & " FROM         dbo.tDEL INNER JOIN " _
'                                  & " dbo.tDELL ON dbo.tDELL.DELL_TR_ID = dbo.tDEL.DEL_ID INNER JOIN " _
'                                  & " dbo.tRL ON dbo.tRL.RL_P_ID = dbo.tDELL.DELL_P_ID INNER JOIN " _
'                                  & " dbo.tTR ON dbo.tDELL.DELL_TR_ID = dbo.tTR.TR_ID INNER JOIN " _
'                                  & " dbo.tTR AS tTR_1 ON dbo.tRL.RL_TR_ID = tTR_1.TR_ID AND dbo.tTR.TR_TP_ID = tTR_1.TR_TP_ID " _
'            & " WHERE     (dbo.tRL.RL_TR_ID = " & lngTRID & " ) " _
'                        & " AND (dbo.tTR.TR_Status IN (3, 4)) " _
'                        & " AND (dbo.tDEL.DEL_SupplierInvoiceDate <= {d '" & ReverseDate(pMinDate) & "'} ) " _
'                        & " AND (dbo.tDEL.DEL_SupplierInvoiceDate >= {d '" & ReverseDate(pMaxDate) & "'} ) " _
'                        & " AND dbo.tDELL.DELL_QtySS - ISNULL(dbo.tRL.RL_QtyReturned, 0) > 0 " _
'            & " ORDER BY dbo.tRL.RL_ID, dbo.tDEL.DEL_SupplierInvoiceDate"
'        rs.Close
'        rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
'        Do While Not rs.EOF
'            lngCurrentRLID = rs!RL_ID
'            strRef = ""
'            iCount = 0
'            Do While lngCurrentRLID = rs!RL_ID And iCount < 4
'                If rs!DELL_QtySS > 0 Then
'                    strRef = strRef & IIf(Len(strRef) > 0, ",", "") & FNS(rs!DEL_SupplierInvoiceRef) & ": " & Format(rs!DEL_SupplierInvoiceDate, "dd/mm/yyyy") & "(" & rs!DELL_QtySS & ")"
'                    lngPrice = rs!DELL_Price
'                    dblDiscount = rs!DELL_Discount
'                End If
'                oPC.COShort.execute "UPDATE tDELL SET DELL_RL_ID = " & lngCurrentRLID & " WHERE DELL_ID = " & rs!DELL_ID
'                rs.MoveNext
'                iCount = 1 + 1
'                If rs.EOF Then Exit Do
'            Loop
'            oPC.COShort.execute "UPDATE tRL SET RL_SIRef = '" & Right(strRef, 300) & "' ,RL_Discount = " & dblDiscount & " , RL_Price = " & lngPrice & " WHERE RL_ID = " & lngCurrentRLID
'        Loop
    End If
    rs.Close
    Set rs = Nothing
    rsTR.Close
    Set rsTR = Nothing
    rsR.Close
    Set rsR = Nothing
    rsRL.Close
    Set rsRL = Nothing
    
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.GenerateReturnsPerPubNew(pTPID,pMinDate,pMaxDate,pStaffID," & _
        "bSupplementary)", Array(pTPID, pMinDate, pMaxDate, pSTAFFID, bSupplementary)
End Function
Public Function UpdateReturnLineRequested(pRLID As Long, pQty As Long)
    On Error GoTo errHandler
Dim strSQL As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "UPdate tRL SET tRL.RL_QTYRequested = " & pQty & " WHERE RL_ID = " & pRLID
    oPC.COShort.execute strSQL
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.UpdateReturnLineRequested(pRLID,pQty)", Array(pRLID, pQty)
End Function
Public Function UpdateReturnLineReturning(pRLID As Long, pQty As Long)
    On Error GoTo errHandler
Dim strSQL As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "UPdate tRL SET tRL.RL_QTYReturned = " & pQty & " WHERE RL_ID = " & pRLID
    oPC.COShort.execute strSQL
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.UpdateReturnLineReturning(pRLID,pQty)", Array(pRLID, pQty)
End Function
Public Function UpdateReturnLineApproved(pRLID As Long, pQty As Long)
    On Error GoTo errHandler
Dim strSQL As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "UPdate tRL SET tRL.RL_QTYApproved = " & pQty & " WHERE RL_ID = " & pRLID
    oPC.COShort.execute strSQL
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.UpdateReturnLineApproved(pRLID,pQty)", Array(pRLID, pQty)
End Function
Public Function UpdateReturnLineNote(pRLID As Long, pNote As String)
    On Error GoTo errHandler
Dim strSQL As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "UPdate tRL SET tRL.RL_Memo = '" & pNote & "' WHERE RL_ID = " & pRLID
    oPC.COShort.execute strSQL
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.UpdateReturnLineNote(pRLID,pNote)", Array(pRLID, pNote)
End Function
Public Function UpdateReturnMemo(pRID As Long, pNote As String)
    On Error GoTo errHandler
Dim strSQL As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "UPdate tTR SET tTR.TR_Note = '" & pNote & "' WHERE TR_ID = " & pRID
    oPC.COShort.execute strSQL
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.UpdateReturnMemo(pRID,pNote)", Array(pRID, pNote)
End Function

Public Function GenerateReturns_Step1(pRID As Long, pRs As ADODB.Recordset, pSTAFFID As Long) As Boolean
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim oSQL As New z_SQL
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oSQL.DropTable "tmpReturns"
    GenerateReturns_Step1 = True
    oPC.COShort.BeginTrans
    oPC.COShort.execute "Create table tmpReturns (RLID INT,Counted Int,Requested Int)"
    Set rs = pRs
    rs.MoveFirst
    Do While Not rs.eof
        oPC.COShort.execute "INSERT INTO tmpReturns (RLID,COUNTED,Requested) VALUES (" & rs.Fields(0) & "," & rs.Fields("Counted") & "," & rs.Fields("Requested") & ")"
'        If pRS!System <> pRS!COunted Then
'            AdjustStock rs!pID, pRS!COunted, pStaffID
'        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    strSQL = "UPdate tRL SET tRL.RL_QTYReturned = tmpReturns.Requested,tRL.RL_QTYApproved = tmpReturns.Requested,tRL.RL_QTYRequested = tmpReturns.Requested,tRL.RL_QTYCounted = tmpReturns.Counted FROM tRL, tmpReturns   WHERE RL_TR_ID = " & pRID & " AND  RL_ID = RLID"
    oPC.COShort.execute strSQL
    strSQL = "DELETE FROM tRL WHERE RL_QTYRequested < 1 and  RL_TR_ID = " & pRID
    oPC.COShort.execute strSQL
    oPC.COShort.execute "Update tTR Set TR_Status = 3,TR_STAFFID = " & pSTAFFID & "  WHERE TR_ID = " & pRID
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.GenerateReturns_Step1(pRID,pRs,pStaffID)", Array(pRID, pRs, pSTAFFID)
End Function
Public Function GenerateReturns_Step2(pRID As Long, pRs As ADODB.Recordset, pSTAFFID As Long, pApprovalRef As String, pApprovalDate As Date) As Boolean
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim oSQL As New z_SQL
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oSQL.DropTable "tmpReturns"
    GenerateReturns_Step2 = True
    oPC.COShort.BeginTrans
    oPC.COShort.execute "Create table tmpReturns (RLID INT,Approved Int,Returned Int)"
    Set rs = pRs
    rs.MoveFirst
    Do While Not rs.eof
        oPC.COShort.execute "INSERT INTO tmpReturns (RLID,Approved,Returned) VALUES (" & rs.Fields("RLID") & "," & rs.Fields("Approved") & "," & rs.Fields("Returned") & ")"
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    '''NOTE: rather keep history, filter out on printing
  '  strSQL = "DELETE FROM tRL WHERE RL_QTYReturned < 1 and  RL_TR_ID = " & pRID
  '  oPC.COShort.execute strSQL

    strSQL = "UPdate tRL SET tRL.RL_QTYReturned = tmpReturns.Returned,tRL.RL_QTYApproved = tmpReturns.Approved FROM tRL, tmpReturns   WHERE RL_TR_ID = " & pRID & " AND  RL_ID = RLID"
    oPC.COShort.execute strSQL
    strSQL = "update tProduct set P_QTYOnHand = ISNULL(a.P_QTYOnHand,0) - b.Dif FROM tPRODUCT a " _
        & " JOIN (select P_ID,sum(tRL.RL_QtyReturned) dif From tPRODUCT JOIN tRL ON RL_P_ID = P_ID Where RL_TR_ID = " & pRID _
        & " group by RL_P_ID,P_ID) b ON a.P_ID = b.P_ID"
    oPC.COShort.execute strSQL
    strSQL = "UPdate tSTOREP SET tStoreP.STP_QTYOnHand = tStoreP.STP_QTYOnHand - tRL.RL_QtyReturned FROM tStoreP JOIN tRL ON RL_P_ID = STP_P_ID WHERE RL_TR_ID = " & pRID
    oPC.COShort.execute strSQL
    


    oPC.COShort.execute "Update tTR Set TR_Status = 4,TR_Date = '" & ReverseDate(Date) & "',TR_CaptureDate = '" & ReverseDateTime(Now) & "' WHERE TR_ID = " & pRID
    oPC.COShort.execute "Update tR Set R_ApprovalREf = '" & pApprovalRef & "' ,R_ApprovalTermDate =  '" & ReverseDate(pApprovalDate) & "'  WHERE R_ID = " & pRID
    oPC.COShort.CommitTrans
    GenerateReturns_Step2 = True
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.GenerateReturns_Step2(pRID,pRs,pStaffID,pApprovalRef,pApprovalDate)", _
         Array(pRID, pRs, pSTAFFID, pApprovalRef, pApprovalDate)
End Function
Public Function IssueReturn(pTRID As Long)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
        strSQL = "UPdate tProduct SET tProduct.P_QTYOnHand = tProduct.P_QTYOnHand - tRL.RL_QtyReturned FROM tPRODUCT JOIN tRL ON RL_P_ID = P_ID  WHERE RL_TR_ID = " & pTRID
        oPC.COShort.execute strSQL
        strSQL = "UPdate tSTOREP SET tStoreP.STP_QTYOnHand = tStoreP.STP_QTYOnHand - tRL.RL_QtyReturned FROM tStoreP JOIN tRL ON RL_P_ID = STP_P_ID WHERE RL_TR_ID = " & pTRID
        oPC.COShort.execute strSQL
        oPC.COShort.execute "Update tTR Set TR_Status = 4 ,TR_CaptureDate = GetDate() WHERE TR_ID = " & pTRID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.IssueReturn(pTRID)", pTRID
End Function
Public Function CountProductsPerPTID(pPTID As Long) As Long
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_CountProductsPerPT"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PTID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = pPTID
    Set par = cmd.CreateParameter("@COUNT", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    CountProductsPerPTID = cmd.Parameters(1)
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CountProductsPerPTID(pPTID)", pPTID
End Function

Public Function CountProductsPerSECID(pPTID As Long) As Long
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_CountProductsPerSEC"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PTID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = pPTID
    Set par = cmd.CreateParameter("@COUNT", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    CountProductsPerSECID = cmd.Parameters(1)
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CountProductsPerSECID(pPTID)", pPTID
End Function

Public Function CountCustomersPerCTID(pCTID As Long) As Long
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_CountCustomersPerCT"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@CTID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = pCTID
    Set par = cmd.CreateParameter("@COUNT", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    CountCustomersPerCTID = cmd.Parameters(1)
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CountCustomersPerCTID(pCTID)", pCTID
End Function

Public Sub MergePTs(pPTOld As Long, pPTNew As Long)
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "Update tProduct SET P_PRODUCTTYPE_ID = " & pPTNew & " WHERE P_PRODUCTTYPE_ID = " & pPTOld
    oPC.COShort.execute "DELETE FROM tPT WHERE PT_ID = " & pPTOld
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_StockManager.MergePTs(pPTOld,pPTNew)", Array(pPTOld, pPTNew)
    Exit Sub
errHandler:
    ErrorIn "z_StockManager.MergePTs(pPTOld,pPTNew)", Array(pPTOld, pPTNew)
End Sub
Public Sub MergeCurrencies(pLose As Long, pKeep As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "MergeCurrencies"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@KEEP", adInteger, adParamInput, , pKeep)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@LOSE", adInteger, adParamInput, , pLose)
    cmd.Parameters.Append par
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.MergeCurrencies(pLose,pKeep)", Array(pLose, pKeep)
End Sub

Public Sub MergeSECs(pPTOld As Long, pPTNew As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "MergeSections"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@OLDID", adInteger, adParamInput, , pPTOld)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@NEWID", adInteger, adParamInput, , pPTNew)
    cmd.Parameters.Append par
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.MergeSECs(pPTOld,pPTNew)", Array(pPTOld, pPTNew)
End Sub
Public Sub MergeCTs(pOld As Long, pNew As Long)
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "Update tTP_IG  SET TPIG_IG_ID = " & pNew & " WHERE TPIG_IG_ID = " & pOld & " AND  NOT EXISTS (SELECT TPIG_IG_ID FROM tTP_IG a WHERE TPIG_TP_ID = tTP_IG.TPIG_TP_ID AND a.TPIG_IG_ID = " & pNew & ")"
    oPC.COShort.execute "DELETE tTP_IG   WHERE TPIG_IG_ID = " & pOld '& " AND  NOT EXISTS (SELECT TPIG_IG_ID FROM tTP_IG a WHERE a.TPIG_TP_ID = TPIG_TP_ID AND a.TPIG_IG_ID = " & pNew & ")"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.MergeCTs(pOld,pNew)", Array(pOld, pNew)
End Sub
Public Sub DeleteUnusedPTs()
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "DELETE dbo.tPT FROM dbo.tPT LEFT OUTER JOIN dbo.tProduct ON dbo.tProduct.P_ProductType_ID = dbo.tPT.PT_ID Where (dbo.tProduct.P_ProductType_ID Is Null)"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.DeleteUnusedPTs"
End Sub
Public Sub DeleteUnusedCurrs()
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
   ' oPC.COShort.execute "DELETE dbo.tCurrency FROM dbo.tCurrency LEFT OUTER JOIN dbo.tProduct ON dbo.tProduct.P_ProductType_ID = dbo.tPT.PT_ID Where (dbo.tProduct.P_ProductType_ID Is Null)"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.DeleteUnusedCurrs"
End Sub
Public Sub DeleteUnusedSECs()
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "DELETE FROM tDICT WHERE DICT_ID IN (SELECT ID FROM vUNUSEDSECTIONS)"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.DeleteUnusedSECs"
End Sub
Public Sub DeleteUnusedCTs()
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "DELETE dbo.tDICT FROM dbo.tDICT LEFT OUTER JOIN dbo.tTP_IG ON dbo.tDICT.DICT_ID = dbo.tTP_IG.TPIG_IG_ID WHERE ISNULL(DICT_TYPE,0) = 2 AND TPIG_IG_ID IS NULL"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.DeleteUnusedCTs"
End Sub

Public Function SetMemo(val As String, pTRID As Long) As Boolean
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "UPDATE tTR SET TR_NOTE = '" & Left(Replace(val, "'", "''"), 5000) & "' WHERE TR_ID = " & pTRID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.SetMemo(val,pTRID)", Array(val, pTRID)
End Function
Public Function SetDispatchModeID(ID As Long, pTRID As Long) As Boolean
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "UPDATE tPO SET PO_DispatchModeID = " & CStr(ID) & " WHERE PO_ID = " & pTRID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.SetDispatchModeID(ID,pTRID)", Array(ID, pTRID)
End Function

Public Function SetCORef(val As String, pTRID As Long) As Boolean
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "UPDATE tCO SET CO_ORDERNUM = '" & val & "' WHERE CO_ID = " & pTRID
    oPC.COShort.execute "UPDATE tCOL SET COL_REF = '" & val & "' WHERE COL_TR_ID = " & pTRID & " AND ISNULL(COL_REF,'') = ''"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.SetCORef(val,pTRID)", Array(val, pTRID)
End Function
'Public Function SetINVRef(val As String, pTRID As Long) As Boolean
'    On Error GoTo ErrHandler
'Dim OpenResult As Integer
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    oPC.COShort.execute "UPDATE tInvoice SET I_Ref = '" & val & "' WHERE I_ID = " & pTRID
'    oPC.COShort.execute "UPDATE tCOL SET COL_REF = '" & val & "' WHERE COL_TR_ID = " & pTRID & " AND ISNULL(COL_REF,'') = ''"
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_StockManager.SetCORef(val,pTRID)", Array(val, pTRID)
'End Function

Public Function SetForAttnCO(val As String, pTRID As Long) As Boolean
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "UPDATE tCO SET CO_ForAttn = '" & val & "' WHERE CO_ID = " & pTRID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.SetForAttn(val,pTRID)", Array(val, pTRID)
End Function
Public Function SetForAttnINV(val As String, pTRID As Long) As Boolean
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "UPDATE tInvoice SET I_ForAttn = '" & val & "' WHERE I_ID = " & pTRID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.SetForAttnINV(val,pTRID)", Array(val, pTRID)
End Function
Public Function SetForAttnQU(val As String, pTRID As Long) As Boolean
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "UPDATE tQU SET QU_ForAttn = '" & val & "' WHERE QU_ID = " & pTRID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.SetForAttnQU(val,pTRID)", Array(val, pTRID)
End Function

Public Function SetCOMemoandOrderRef(pMemo As String, pOrderRef As String, pTRID As Long) As Boolean
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "UPDATE tTR SET TR_NOTE = '" & pMemo & "' WHERE TR_ID = " & pTRID
    oPC.COShort.execute "UPDATE tCO SET CO_ORDERNUM = '" & pOrderRef & "' WHERE CO_ID = " & pTRID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.SetCOMemoandOrderRef(pMemo,pOrderRef,pTRID)", Array(pMemo, pOrderRef, _
         pTRID)
End Function

Public Sub DeleteRelatedDocs(pTP As Long)
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "DELETE FROM dbo.tTR Where TR_TP_ID = pTP"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.DeleteRelatedDocs(pTP)", pTP
End Sub

Public Function CancelR(PID As Long)
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    strSQL = "SELECT RL_P_ID,RL_QTYRETURNED FROM tRL WHERE RL_TR_ID = " & PID
    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
    
    oPC.COShort.BeginTrans
    Do While Not rs.eof
        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnBackorder,STP_QTYRESERVED,STP_QTYONHAND,STP_QTYCOPIESONHAND,STP_QTYCOPIESRESERVED FROM tStoreP Where STP_P_ID = '" & rs.Fields("RL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        If Not rsSP.eof Then
            rsSP.Fields("STP_QTYONHAND") = FNN(rsSP.Fields("STP_QTYONHAND")) + FNN(rs.Fields("RL_QTYRETURNED"))
            rsSP.Update
        End If
        strSQL = "SELECT P_QTYONHAND FROM tPRODUCT WHERE P_ID = '" & rs.Fields("RL_P_ID") & "'"
        Set rsP = New ADODB.Recordset
        rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) + FNN(rs.Fields("RL_QTYRETURNED"))
        If Not rsP.eof Then rsP.Update
        rsP.Close
        Set rsP = Nothing
        
        
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    rsSP.Close
    Set rsSP = Nothing
    oPC.COShort.execute "UPDATE tTR SET TR_STATUS = 6 WHERE TR_ID = " & PID
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CancelR(PID)", PID
End Function

Public Sub ReverseReturnToRequested(pTRID As Long)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
        oPC.COShort.BeginTrans
    
        strSQL = "UPdate tProduct SET tProduct.P_QTYOnHand = ISNULL(tProduct.P_QTYOnHand,0) + tRL.RL_QtyReturned FROM tPRODUCT JOIN tRL ON RL_P_ID = P_ID  WHERE RL_TR_ID = " & pTRID
        oPC.COShort.execute strSQL
        strSQL = "UPdate tSTOREP SET tStoreP.STP_QTYOnHand = IsNull(tStoreP.STP_QTYOnHand,0) + tRL.RL_QtyReturned FROM tStoreP JOIN tRL ON RL_P_ID = STP_P_ID WHERE RL_TR_ID = " & pTRID
        oPC.COShort.execute strSQL

    oPC.COShort.execute "Update tTR Set TR_Status = 3 WHERE TR_ID = " & pTRID
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.ReverseReturnToRequested(pTRID)", pTRID
End Sub
Public Function CancelRL(pRLID As Long, pPID As String, pQty As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSQL As String
Dim rsP As ADODB.Recordset
Dim rsSTP As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.BeginTrans
    oPC.COShort.execute "UPDATE tRL SET RL_STatus = 'CAN' WHERE RL_ID = " & pRLID
    
    Set rsP = New ADODB.Recordset
    strSQL = "SELECT P_QTYONHAND FROM tPRODUCT WHERE P_ID = '" & pPID & "'"
    rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) + IIf(IsNull(pQty), 0, pQty)
    If Not rsP.eof Then rsP.Update
    rsP.Close
    Set rsP = Nothing
    
    
    Set rsSTP = New ADODB.Recordset
    strSQL = "SELECT STP_QTYONHAND FROM tStoreP WHERE STP_P_ID = '" & pPID & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
    rsSTP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
        rsSTP!STP_QTYONHAND = FNN(rsSTP!STP_QTYONHAND) + IIf(IsNull(pQty), 0, pQty)
    If Not rsSTP.eof Then rsSTP.Update
    rsSTP.Close
    Set rsSTP = Nothing
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CancelRL(pRLID,pPID,pQty)", Array(pRLID, pPID, pQty)
End Function

Public Function GetLabelsToPrint(pType As String, pStart As Boolean, Optional pTRID As Long, _
                Optional pCode As String, Optional pCodeSet As String) As XArrayDB
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Static XA As New XArrayDB

Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.CursorLocation = adUseClient
    If pType = "D" Then
        rs.Open "SELECT * FROM vLabels_D WHERE TRID = " & pTRID, oPC.COShort, adOpenStatic, adLockReadOnly
    ElseIf pType = "T" Then
        rs.Open "SELECT * FROM vLabels_T WHERE TRID = " & pTRID, oPC.COShort, adOpenStatic, adLockReadOnly
    ElseIf pType = "P" Then
        If Len(pCode) = 13 Then
            rs.Open "SELECT * FROM vLabels_P WHERE P_EAN = '" & pCode & "'", oPC.COShort, adOpenStatic, adLockReadOnly
        Else
            rs.Open "SELECT * FROM vLabels_P WHERE P_CODE = '" & pCode & "'", oPC.COShort, adOpenStatic, adLockReadOnly
        End If
    ElseIf pType = "S" Then   'a set of codes
        rs.Open "SELECT * FROM vLabels_P WHERE P_EAN in ('" & Replace(pCodeSet, ",", "','") & "')", oPC.COShort, adOpenStatic, adLockReadOnly
    End If
    If pStart Then XA.ReDim 1, 0, 1, 23
    Do While Not rs.eof
        XA.InsertRows XA.UpperBound(1) + 1, 1
        XA(XA.UpperBound(1), 1) = Trim(rs!CodeF)       'FormatCode(FNS(rs!P_Code), True)
        XA(XA.UpperBound(1), 2) = FNS(rs!P_Title)
        XA(XA.UpperBound(1), 3) = Format(FNN(rs!P_SP) / oPC.Configuration.DefaultCurrency.Divisor, "R###,##0.00")
        XA(XA.UpperBound(1), 4) = CStr(FNN(rs!QtySS) + FNN(rs!QtyFirm))
        XA(XA.UpperBound(1), 5) = CStr(FNN(rs!RepriceCount))
        XA(XA.UpperBound(1), 6) = CStr(FNN(rs!RepriceCount) + FNN(rs!QtySS) + FNN(rs!QtyFirm))
        XA(XA.UpperBound(1), 7) = False
        XA(XA.UpperBound(1), 8) = FNS(rs!P_EAN)
        XA(XA.UpperBound(1), 9) = FNN(rs!P_SP)
        XA(XA.UpperBound(1), 10) = FNS(rs!AcNo)
        XA(XA.UpperBound(1), 11) = FNS(rs!MainCategoryCode)
        XA(XA.UpperBound(1), 12) = FNS(rs!MainCategory)
        XA(XA.UpperBound(1), 13) = FNS(rs!P_MainAuthor)
        XA(XA.UpperBound(1), 14) = FNS(rs!ProductCode)
        XA(XA.UpperBound(1), 15) = FNS(rs!PT_Code)
        XA(XA.UpperBound(1), 16) = FNS(rs!P_DateRecordAdded)
        XA(XA.UpperBound(1), 17) = IIf(IsNull(rs!P_LastDateDelivered), "", Format(FND(rs!P_LastDateDelivered), "MM/YY"))
        XA(XA.UpperBound(1), 18) = FNS(rs!AllCategoriesShort)
        XA(XA.UpperBound(1), 19) = FNS(rs!P_Length)
        XA(XA.UpperBound(1), 20) = FNS(rs!P_Width)
        XA(XA.UpperBound(1), 21) = FNS(rs!P_Code)
        XA(XA.UpperBound(1), 22) = FNS(rs!Seq)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set GetLabelsToPrint = XA
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.GetLabelsToPrint(pType,pStart,pTRID,pCODE)", Array(pType, pStart, pTRID, _
         pCode)
End Function

'Public Function GetApproSlips(pAPPID As Long) As ADODB.Recordset
'    On Error GoTo ErrHandler
'Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'Dim OpenResult As Integer
''-------------------------------
''    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    rs.Open "SELECT * FROM vAPPROSLIPS WHERE TR_ID = " & pAPPID & " ORDER BY AUTHOR", oPC.COShort, adOpenForwardOnly, adLockReadOnly
'    Set GetApproSlips = rs
''    Set rs = Nothing
''---------------------------------------------------
''    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'    Exit Function
'ErrHandler:
'    RlsObjs rs
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_StockManager.GetApproSlips(pAPPID)", pAPPID
'End Function

Sub FixMissingInvoiceRefsOnReturn(pRID As Long, pMinDate As Date, pMaxDate As Date)
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim strRef As String
Dim iCount As Integer
Dim lngPrice As Long
Dim dblDiscount As Double
Dim lngCurrentRLID As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT DELL_QTYSS,RL_ID,DEL_SupplierInvoiceRef,DEL_SupplierInvoiceDate,DELL_Price,DELL_Discount " _
            & " FROM tDEL " _
            & " INNER JOIN tDELL on DELL_TR_ID = DEL_ID " _
            & " INNER JOIN tRL ON RL_P_ID = DELL_P_ID " _
            & " INNER JOIN tTR ON DELL_TR_ID = TR_ID " _
            & " WHERE RL_TR_ID = " & pRID & " AND (TR_STATUS = 3 or TR_Status = 4) " _
            & " AND DEL_SupplierInvoiceDate >= {d '" & ReverseDate(pMaxDate) & "'} " _
            & " AND DEL_SupplierInvoiceDate <= {d '" & ReverseDate(pMinDate) & "'} " _
            & " ORDER BY RL_ID,DEL_SupplierInvoiceDate"
    rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
    Do While Not rs.eof
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
            If rs.eof Then Exit Do
        Loop
        oPC.COShort.execute "UPDATE tRL SET RL_SIRef = '" & strRef & "' ,RL_Discount = " & dblDiscount & " , RL_Price = " & lngPrice & " WHERE RL_ID = " & lngCurrentRLID
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.FixMissingInvoiceRefsOnReturn(pRID,pMinDate,pMaxDate)", Array(pRID, _
         pMinDate, pMaxDate)
End Sub

Public Function CancelAPPRL(pAPPRLID As Long) ', pPID As String, pQty As Long)
    On Error GoTo errHandler
Dim strSQL As String
Dim rsP As ADODB.Recordset
Dim rsSP As ADODB.Recordset
Dim rsAPPL As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim rsAPP As ADODB.Recordset
Dim lngQty As Long
Dim strPID As String
Dim bFulfilled As Boolean
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tAPPRL WHERE APPRL_ID = " & pAPPRLID

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic

    'oPC.COShort.BeginTrans
        Set rsAPPL = New ADODB.Recordset
        strSQL = "SELECT APPL_TR_ID,APPL_P_ID,APPL_QTY,APPL_QTYReturned,APPL_Fulfilled FROM tAPPL WHERE APPL_ID = " & rs!APPRL_APPL_ID
        rsAPPL.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
            rsAPPL!APPL_QtyReturned = FNN(rsAPPL!APPL_QtyReturned) - FNN(rs!APPRL_Qty)
            If FNN(rsAPPL!APPL_Qty) > FNN(rsAPPL!APPL_QtyReturned) Then
                rsAPPL!APPL_Fulfilled = "OS"
            End If
        If Not rsAPPL.eof Then rsAPPL.Update
        
        strSQL = "Select APPL_QTY,APPL_QTYReturned from tAPPL WHERE APPL_TR_ID = " & rsAPPL!APPL_TR_ID
        Set rsAPP = New ADODB.Recordset
        rsAPP.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
        bFulfilled = True
        Do While Not rsAPP.eof
            If FNN(rsAPP!APPL_Qty) > FNN(rsAPP!APPL_QtyReturned) Then
                bFulfilled = False
                Exit Do
            End If
            rsAPP.MoveNext
        Loop
        If bFulfilled Then
            oPC.COShort.execute "UPDATE tTR SET TR_STATUS = 4 WHERE TR_ID = " & rsAPPL!APPL_TR_ID
        Else
            oPC.COShort.execute "UPDATE tTR SET TR_STATUS = 3 WHERE TR_ID = " & rsAPPL!APPL_TR_ID
        End If
        rsAPP.Close
        Set rsAPP = Nothing
        Set rsSP = New ADODB.Recordset
        strSQL = "SELECT STP_QtyOnHand,STP_QtyOnAppro FROM tStoreP Where STP_P_ID = '" & rsAPPL.Fields("APPL_P_ID") & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        rsSP.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            rsSP.Fields("STP_QtyOnHand") = FNN(rsSP.Fields("STP_QtyOnHand")) - FNN(rs.Fields("APPRL_Qty"))
            rsSP.Fields("STP_QtyOnAppro") = FNN(rsSP.Fields("STP_QtyOnAppro")) + FNN(rs.Fields("APPRL_Qty"))
        If Not rsSP.eof Then rsSP.Update
        rsSP.Close
        Set rsSP = Nothing

        Set rsP = New ADODB.Recordset
        strSQL = "SELECT P_QtyOnHand,P_QtyOnAppro FROM tProduct Where P_ID = '" & rsAPPL.Fields("APPL_P_ID") & "'"
        rsP.Open strSQL, oPC.COShort, adOpenKeyset, adLockPessimistic
            rsP!P_QtyOnHand = FNN(rsP!P_QtyOnHand) - FNN(rs!APPRL_Qty)
            rsP!P_QtyOnAppro = FNN(rsP!P_QtyOnAppro) + FNN(rs!APPRL_Qty)
        If Not rsP.eof Then rsP.Update
        rsP.Close
        Set rsP = Nothing


        rs.MoveNext
        rsAPPL.Close
        Set rsAPPL = Nothing
        
   ' oPC.COShort.CommitTrans
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CancelAPPRL(pAPPRLID)", Array(pAPPRLID)
End Function
Public Sub RejectReturn(pQty As Long, pRLID As Long, pNote As String)
    On Error GoTo errHandler

    'subtract existing qty from pQty, adjust the OH balance by adding the difference
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_RejectReturn"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@RLID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = pRLID
    Set par = cmd.CreateParameter("@TOTALQTYREJECTED", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = pQty
    Set par = cmd.CreateParameter("@NOTE", adVarChar, adParamInput, 100)
    cmd.Parameters.Append par
    par.Value = pNote
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.RejectReturn(pQty,pRLID,pNote)", Array(pQty, pRLID, pNote)
End Sub
Public Function RediarizePOLS(pPID As String, pNewETA As String, pReason As String) As Boolean
    On Error GoTo errHandler
'Sets the POL_ETA date to the new date for all POLs for the specified PID
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT POL_ID,'" & pReason & "' as Note,'n' as f3,'" & pNewETA & "' as f4,POL_P_ID as PID,'' as dummy2 FROM tPOL JOIN tTR on TR_ID = POL_TR_ID WHERE POL_P_ID = '" & pPID & "' and POL_Fulfilled = 'os' and (TR_Status = 3 or TR_Status = 2)"
    Set rs = New ADODB.Recordset
    oPC.COShort.BeginTrans
    rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
    If Not rs.eof Then
        ActionODPO rs
    End If
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.RediarizePOLS(pPID,pNewETA,pReason)", Array(pPID, pNewETA, pReason)
End Function
Public Function RediarizeCOLS(pPID As String, pNewETA As String, pReason As String) As Boolean
    On Error GoTo errHandler
'Sets the POL_ETA date to the new date for all POLs for the specified PID
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT COL_ID,'" & pReason & "' as Note,'n' as f3,'" & pNewETA & "' as f4 FROM tCOL JOIN tTR on TR_ID = COL_TR_ID WHERE COL_P_ID = '" & pPID & "' and COL_Fulfilled = 'OS' and (TR_Status = 3 or TR_Status = 2)"
    Set rs = New ADODB.Recordset
    oPC.COShort.BeginTrans
    rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
    If Not rs.eof Then
        ActionODCO rs
    End If
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.RediarizeCOLS(pPID,pNewETA,pReason)", Array(pPID, pNewETA, pReason)
End Function

Public Function CANCELPOLs(pPID As String, pReason As String) As Boolean
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT POL_ID,'" & pReason & "' as f2,'C' as f3,'' as f4 FROM tPOL JOIN tTR on TR_ID = POL_TR_ID WHERE POL_P_ID = '" & pPID & "' and POL_Fulfilled = 'os' and (TR_Status = 3 or TR_Status = 2)"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
    ActionODPO rs
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CancelPOLS(pPID,pReason)", Array(pPID, pReason)
End Function
Public Function CancelPOLSSimple(pPID As String, strReason As String)
Dim rs As ADODB.Recordset
    strSQL = "SELECT POL_ID,'" & strReason & "' as f2,'C' as f3,'' as f4 FROM tPOL JOIN tTR on TR_ID = POL_TR_ID WHERE POL_P_ID = '" & pPID & "' and POL_Fulfilled = 'os' and (TR_Status = 3 or TR_Status = 2)"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
    Do While Not rs.eof
        oPC.COShort.execute "INSERT INTO tPOLA (POLA_POL_ID,POLA_DATE,POLA_CODE,POLA_REPORT) VALUES (" & rs.Fields(0) & ",'" & ReverseDateTime(Date) & "','" & FNS(rs.Fields(3)) & "','" & strReason & "')"
        rs.MoveNext
    Loop
        oPC.COShort.execute "UPDATE tPOL SET POL_Fulfilled = 'CAN' WHERE ISNULL(POL_Fulfilled,'OS') IN ('OS','') AND POL_P_ID = '" & pPID & "'"
End Function
Public Function CANCELCOLs(pPID As String, pReason As String) As Boolean
    On Error GoTo errHandler
'Sets the POL_ETA date to the new date for all POLs for the specified PID
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT POL_ID,'" & pReason & "' as f2,'C' as f3,'' as f4 FROM tCOL JOIN tTR on TR_ID = COL_TR_ID WHERE COL_P_ID = '" & pPID & "' and COL_Fulfilled = 'OS' and (TR_Status = 3 or TR_Status = 2)"
    Set rs = New ADODB.Recordset
    oPC.COShort.BeginTrans
    rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
    ActionODCO rs
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CancelCOLS(pPID,pReason)", Array(pPID, pReason)
End Function
Public Function CancelCOLSSimple(pPID As String, strReason As String)
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    'strSQL = "SELECT COL_ID,'" & strREason & "' as f2,'C' as f3,'' as f4 FROM tCOL JOIN tTR on TR_ID = COL_TR_ID WHERE COL_P_ID = '" & pPID & "' and COL_Fulfilled = 'os' and (TR_Status = 3 or TR_Status = 2)"
    strSQL = "SELECT COL_ID FROM tCOL JOIN tTR on TR_ID = COL_TR_ID WHERE COL_P_ID = '" & pPID & "' and COL_Fulfilled = 'os' and (TR_Status = 3 or TR_Status = 2)"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
    Do While Not rs.eof
        oPC.COShort.execute "INSERT INTO tCOLA (COLA_COL_ID,COLA_DATE,COLA_REPORT) VALUES (" & rs.Fields("COL_ID") & ",'" & ReverseDateTime(Date) & "','" & strReason & "')"
        oPC.COShort.execute "UPDATE tCOL SET COL_Fulfilled = 'CAN' WHERE ISNULL(COL_Fulfilled,'OS') IN ('OS','') AND COL_ID = " & rs.Fields("COL_ID")
        rs.MoveNext
    Loop
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function
Public Function RediarizeCOLSSimple(pPID As String, pNewETA As Date) As Boolean
    On Error GoTo errHandler
'Sets the POL_ETA date to the new date for all POLs for the specified PID
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT COL_ID FROM tCOL JOIN tTR on TR_ID = COL_TR_ID WHERE COL_P_ID = '" & pPID & "' and COL_Fulfilled = 'OS' and (TR_Status = 3 or TR_Status = 2)"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
    Do While Not rs.eof
        oPC.COShort.execute "UPDATE tCOL SET COL_ETA = '" & ReverseDate(pNewETA) & "' WHERE COL_ID = " & rs.Fields("COL_ID")
        rs.MoveNext
    Loop
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.RediarizeCOLSSimple(pPID,pNewETA)", Array(pPID, pNewETA)
End Function
Public Function RediarizePOLSSimple(pPID As String, pNewETA As Date) As Boolean
    On Error GoTo errHandler
'Sets the POL_ETA date to the new date for all POLs for the specified PID
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT POL_ID FROM tPOL JOIN tTR on TR_ID = POL_TR_ID WHERE POL_P_ID = '" & pPID & "' and POL_Fulfilled = 'OS' and (TR_Status = 3 or TR_Status = 2)"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
    Do While Not rs.eof
        oPC.COShort.execute "UPDATE tPOL SET POL_ETA = '" & ReverseDate(pNewETA) & "' WHERE POL_ID = " & rs.Fields("POL_ID")
        rs.MoveNext
    Loop
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.RediarizePOLSSimple(pPID,pNewETA)", Array(pPID, pNewETA)
End Function

Public Function LogTransmission(pTRID As Long, pMsg As String)
    On Error GoTo errHandler
Dim strLog As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strLog = "UPDATE tTR SET TR_LOG = RIGHT('" & pMsg & "' + CHAR(10) + ISNULL(TR_LOG,''),500) WHERE TR_ID = " & pTRID
    oPC.COShort.execute strLog
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.LogTransmission(pTRID,pMsg)", Array(pTRID, pMsg)
End Function
Public Function CustomerIndexClashes(lngTPID As Long, strAcno As String) As Boolean
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_CustomerClash"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = lngTPID
    Set par = cmd.CreateParameter("@ACNO", adVarChar, adParamInput, 20)
    cmd.Parameters.Append par
    par.Value = strAcno
    Set par = cmd.CreateParameter("@R", adInteger, adParamOutput)
    cmd.Parameters.Append par
    par.Value = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    CustomerIndexClashes = cmd.Parameters(2)
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function


    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CustomerIndexClashes(lngTPID,strACno)", Array(lngTPID, strAcno)
End Function


Public Sub Cashup(pZSessionGUID As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_CashUp"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@ZGUID", adGUID, adParamInput)
    cmd.Parameters.Append par
    par.Value = pZSessionGUID
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.Cashup(pZSessionGUID)", pZSessionGUID
End Sub

Public Sub UpdateDispatchLog(pDocCOde As String, pMsg As String)
    On Error GoTo errHandler
Dim strSQL As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "UPDATE tTR SET TR_LOG = TR_LOG + '" & vbCrLf & pMsg & "' WHERE TR_CODE = '" & pDocCOde & "'"
    oPC.COShort.execute strSQL
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.UpdateDispatchLog(pDocCode,pMsg)", Array(pDocCOde, pMsg)
End Sub
Public Sub UpdateBookStatus(pStatus As String, pPID As String)
    On Error GoTo errHandler
Dim strSQL As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "UPDATE tProduct SET P_Status = '" & pStatus & "' WHERE P_ID = '" & pPID & "'"
    oPC.COShort.execute strSQL
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.UpdateBookStatus(pStatus,pPID)", Array(pStatus, pPID)
End Sub

Public Sub RollbackTransaction(Conn As ADODB.Connection)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim Res As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "RollbackTransaction_Safe"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@RES", adInteger, adParamOutput)
    cmd.Parameters.Append par
    cmd.ActiveConnection = Conn
    
    cmd.execute
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.RollbackTransaction(Conn)", Conn
End Sub

Public Sub GenerateConsolidatedGRN(Conn As ADODB.Connection)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim Res As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "ProduceConsolidatedGRN"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@RES", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = Conn
    
    cmd.execute
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.GenerateConsolidatedGRN(Conn)", Conn
End Sub

Public Sub GetMissingExchangeNumbers(pTILLPOINT As String, pMissing As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim Res As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "CheckSkippedExchanges"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TILLPOINT", adVarChar, adParamInput, 20, pTILLPOINT)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@MISSING", adVarChar, adParamOutput, 3000)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    pMissing = cmd.Parameters(1).Value
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.GetMissingExchangeNumbers(pTILLPOINT,pMissing)", Array(pTILLPOINT, _
         pMissing)
End Sub

Public Sub ImportCashSales(pFilename As String, pDate As Date)
    On Error GoTo errHandler
    ImportOneFile pFilename, 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.ImportCashSales(pFilename,pDate)", Array(pFilename, pDate)
End Sub
Private Sub ImportOneFile(pFilePath As String, lngFNID As Long)
    On Error GoTo errHandler
Dim lngLineCount As Long
Dim strCode As String
Dim strQty As String
Dim lngQty As Long
Dim oBatch As z_SQL
Dim strCodeOut As String
Dim strEANOut As String
Dim strLine As String
Dim fs As New FileSystemObject
Dim arLine() As String

    Set oBatch = New z_SQL
    Set txtStream = fs.OpenTextFile(pFilePath)
    
    lngLineCount = 0
    Do While Not txtStream.AtEndOfStream
        txtStream.ReadLine
        lngLineCount = lngLineCount + 1
    Loop
    lngLineCount = 0
    txtStream.Close
    Set txtStream = Nothing
    Set txtStream = fs.OpenTextFile(pFilePath)
    lngLineCount = 0
    Do While Not txtStream.AtEndOfStream
        strLine = txtStream.ReadLine
        If strLine = "" Then GoTo skip
        arLine = Split(strLine, ",")
        strCode = Trim(arLine(0))
        If UBound(arLine) = 0 Then
            lngQty = 1
        Else
            lngQty = CLng(Trim(arLine(1)))
        End If
        If IsGoodCode(strCode, strCodeOut, strEANOut) Then
            lngLineCount = lngLineCount + 1
            retval = oBatch.RunSQL("INSERT INTO TMPIMPORTSALES (Code,QTY) VALUES ('" & CStr(strCode) & "'," & lngQty & ")")
        End If
skip:
    Loop
    Set oBatch = Nothing
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.ImportOneFile(pFilePath,lngFNID)", Array(pFilePath, lngFNID)
End Sub

Public Sub CreateSalesFromImport(pDate As Date, pormessage As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim Res As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_IMPORTSALES_CREATE"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@DTE", adDate, adParamInput, , pDate)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@CODE", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@POSITION", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    If cmd.Parameters(1).Value <> 0 Then   'or found
        pormessage = "or in Stored procedure sp_IMPORTSALES_CREATE at position " & FNS(cmd.Parameters(2).Value)
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CreateSalesFromImport(pDate,pormessage)", Array(pDate, pormessage)
End Sub

Public Sub RecalculateTPBalance(pTPID As Long, pTot As Double, pCur As Double, p30 As Double, p60 As Double, p90 As Double, p120 As Double)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim Res As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
   ' cmd.CommandText = "RecalculateTPBalance"
    cmd.CommandText = "AgeInvoices"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput, , pTPID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@Tot", adNumeric, adParamOutput)
    par.Precision = 15
    par.NumericScale = 2
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@Cur", adNumeric, adParamOutput)
    par.Precision = 15
    par.NumericScale = 2
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@30", adNumeric, adParamOutput)
    par.Precision = 15
    par.NumericScale = 2
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@60", adNumeric, adParamOutput)
    par.Precision = 15
    par.NumericScale = 2
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@90", adNumeric, adParamOutput)
    par.Precision = 15
    par.NumericScale = 2
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@120", adNumeric, adParamOutput)
    par.Precision = 15
    par.NumericScale = 2
    cmd.Parameters.Append par
    Set par = Nothing
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    
    pTot = FNCURR(cmd.Parameters("@Tot").Value)
    pCur = FNCURR(cmd.Parameters("@Cur").Value)
    p30 = FNCURR(cmd.Parameters("@30").Value)
    p60 = FNCURR(cmd.Parameters("@60").Value)
    p90 = FNCURR(cmd.Parameters("@90").Value)
    p120 = FNCURR(cmd.Parameters("@120").Value)
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.RecalculateTPBalance(pTPID,pTot,pCur,p30,p60,p90,p120)", Array(pTPID, _
         pTot, pCur, p30, p60, p90, p120)
End Sub

Public Sub CreateInvoiceFromExchange(pEXCHID As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "CreateInvoiceFromExchange"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , pEXCHID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@TRSTATUS", adInteger, adParamInput, , 3)  'Create the invoice as issued
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort  'THis is used only by PBKS_POSSVR so it must use this connection - DO NOT CHANGE
    
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CreateInvoiceFromExchange(pEXCHID)", pEXCHID
End Sub

Public Sub MarkSubstitute(pOriginal As String, pSubstitute As String, pViceVersa As Boolean)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "MarkSubstitute"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Original", adGUID, adParamInput, , pOriginal)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@Substitute", adGUID, adParamInput, , pSubstitute)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@ViceVersa", adBoolean, adParamInput, , IIf(pViceVersa, 1, 0))
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.MarkSubstitute(pOriginal,pSubstitute,pViceVersa)", Array(pOriginal, _
         pSubstitute, pViceVersa)
End Sub

Public Sub InsertInvoiceSubstitutions(pINVOICEID As Long, pILID As Long, pPID As String, pCOLID As Long, pQty As Long, pSubstituteFor As String, FromInvoiceorGDN As String)
    On Error GoTo errHandler
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "InsertInvoiceSubstitution"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@INVID", adInteger, adParamInput, , pINVOICEID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@ILID", adInteger, adParamInput, , pILID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@PID", adGUID, adParamInput, , pPID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@COLID", adInteger, adParamInput, , pCOLID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@QTY", adInteger, adParamInput, , pQty)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@SubstituteFor", adGUID, adParamInput)
    cmd.Parameters.Append par
    If pSubstituteFor > "" Then
        par.Value = pSubstituteFor
    End If
    Set par = Nothing
    Set par = cmd.CreateParameter("@FromInvoiceorGDN", adVarChar, adParamInput, 1, FromInvoiceorGDN)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort  'THis is used only by PBKS_POSSVR so it must use this connection - DO NOT CHANGE
    
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.InsertInvoiceSubstitutions(pInvoiceID,pILID,pPID,pCOLID,pQty," & _
  "pSubstituteFor)", Array(pINVOICEID, pILID, pPID, pCOLID, pQty, pSubstituteFor), EA_NORERAISE

End Sub

Public Sub InsertAuditRecord(pAuditPoint As String, pReason As String, pOldValue As String, pNewValue As String, pSMID As Long, pPID As String, pTPID As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "InsertAuditRecord"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@AUDITPOINT", adVarChar, adParamInput, 5, pAuditPoint)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@Reason", adVarChar, adParamInput, 100, pReason)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@OldValues", adVarChar, adParamInput, 100, pOldValue)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@NewValues", adVarChar, adParamInput, 100, pNewValue)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@SMID", adInteger, adParamInput, , pSMID)
    cmd.Parameters.Append par
    Set par = Nothing
    If pPID = "" Then
        Set par = cmd.CreateParameter("@PID", adGUID, adParamInput)
    Else
        Set par = cmd.CreateParameter("@PID", adGUID, adParamInput, , pPID)
    End If
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput, , pTPID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.InsertAuditRecord(pAuditPoint,pReason,pOldValue,pNewValue,pSMID)", _
         Array(pAuditPoint, pReason, pOldValue, pNewValue, pSMID)
End Sub
Public Function AuditDaysession(pZID As String) As String
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "CheckSkippedExchangesByZID"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 120
    Set par = cmd.CreateParameter("@ZID", adGUID, adParamInput, , FNS(pZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@MISSING", adVarChar, adParamOutput, 3000)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    AuditDaysession = cmd.Parameters("@Missing").Value
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.AuditDaysession(pZID)", pZID
End Function
Public Function CloseDaysession(pZID As String, pDte As Date) As String
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "CloseDaysessionByZID"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@ZID", adVarChar, adParamInput, 40, FNS(pZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Dte", adDate, adParamInput, , pDte)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CloseDaysession(pZID)", pZID
End Function
Public Function CloseOperatorsession(pXID As String, pDte As Date) As String
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "CloseOperatorsessionByXID"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@XID", adVarChar, adParamInput, 40, FNS(pXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Dte", adDate, adParamInput, , pDte)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CloseOperatorsession(pXID)", pXID
End Function

Public Function AuditAll(dteSince As Date, strTillpoint As String) As String
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter


    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "CheckSkippedExchanges"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@SINCE", adDate, adParamInput, , dteSince)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TILLPOINT", adVarChar, adParamInput, 20, strTillpoint)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@MISSING", adVarChar, adParamOutput, 3000)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    AuditAll = cmd.Parameters("@Missing").Value
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault
    Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.AuditAll(dteSince,strTillpoint)", Array(dteSince, strTillpoint)
End Function

'Public Function FormatISBN13(p As String) As String
'    On Error GoTo ErrHandler
'Dim OpenResult As Integer
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'Dim f As String
'
'    If Not IsNumeric(p) Then
'        FormatISBN13 = p
'        Exit Function
'    End If
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    Set cmd = New ADODB.Command
'    cmd.CommandText = "FormatISBN13"
'    cmd.CommandType = adCmdStoredProc
'
'    Set par = cmd.CreateParameter("@IN", adVarChar, adParamInput, 15, p)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@OUT", adVarChar, adParamOutput, 20, f)
'    cmd.Parameters.Append par
'
'    cmd.ActiveConnection = oPC.COShort
'    cmd.execute
'
'    FormatISBN13 = cmd.Parameters("@OUT").Value
'
'    Set cmd = Nothing
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'    Exit Function
'
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_StockManager.FormatISBN13(p)", p
'End Function
Public Function PreviousBillings(PID As String, TPID As Long) As ADODB.Recordset
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "PreviousBillings"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 50, PID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput, , TPID)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    Set PreviousBillings = cmd.execute
    
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.PreviousBillings(PID,TPID)", Array(PID, TPID)
End Function

Public Sub PasteQuotationLine(PID As String, Qty As Long, Price As Long, DiscountRate As Double, _
VATRate As Double, lngQUID As Long, strRef As String, ExtraChargePID As String, ExtraChargeValue As Long, _
FCPrice As Long, FCFactor As Double, FCID As Long)
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "PasteQuotationLine"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 50, PID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Qty", adInteger, adParamInput, , Qty)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Price", adInteger, adParamInput, , Price)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Discountrate", adDouble, adParamInput, , DiscountRate)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@VATRate", adDouble, adParamInput, , VATRate)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@QUID", adInteger, adParamInput, , lngQUID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@REF", adVarChar, adParamInput, 30, strRef)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@EXTRACHARGEPID", adVarChar, adParamInput, 40, ExtraChargePID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@EXTRACHARGEVALUE", adInteger, adParamInput, , ExtraChargeValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@FCPrice", adInteger, adParamInput, , FCPrice)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@FCFactor", adDouble, adParamInput, , FCFactor)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@FCID", adInteger, adParamInput, , FCID)
    cmd.Parameters.Append par
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub

End Sub
Public Sub PasteTransferLine(PID As String, Qty As Long, Price As Long, DiscountRate As Double, _
VATRate As Double, lngTRID As Long)
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "PasteTransferLine"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 50, PID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Qty", adInteger, adParamInput, , Qty)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Price", adInteger, adParamInput, , Price)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Discountrate", adDouble, adParamInput, , DiscountRate)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@VATRate", adDouble, adParamInput, , VATRate)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TFRID", adInteger, adParamInput, , lngTRID)
    cmd.Parameters.Append par
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub

End Sub

Public Sub PasteApproLine(PID As String, Qty As Long, Price As Long, DiscountRate As Double, VATRate As Double, lngQUID As Long, strRef As String)
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "PasteApproLine"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 50, PID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Qty", adInteger, adParamInput, , Qty)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Price", adInteger, adParamInput, , Price)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Discountrate", adDouble, adParamInput, , DiscountRate)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@VATRate", adDouble, adParamInput, , VATRate)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@QUID", adInteger, adParamInput, , lngQUID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@REF", adVarChar, adParamInput, 30, strRef)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub

End Sub

Public Sub PasteInvoiceLine(PID As String, QtyFirm As Long, QtySS As Long, Price As Long, DiscountRate As Double, _
VATRate As Double, lngQUID As Long, strRef As String, ExtraChargePID As String, ExtraChargeValue As Long, _
FCPrice As Long, FCFactor As Double, FCID As Long)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "PasteInvoiceLine"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 50, PID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@QtyFirm", adInteger, adParamInput, , QtyFirm)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@QtySS", adInteger, adParamInput, , QtySS)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Price", adInteger, adParamInput, , Price)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Discountrate", adDouble, adParamInput, , DiscountRate)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@VATRate", adDouble, adParamInput, , VATRate)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@QUID", adInteger, adParamInput, , lngQUID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@REF", adVarChar, adParamInput, 30, strRef)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@EXTRACHARGEPID", adVarChar, adParamInput, 40, ExtraChargePID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ExtrachargeValue", adInteger, adParamInput, , ExtraChargeValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@FCPrice", adInteger, adParamInput, , FCPrice)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@FCFactor", adDouble, adParamInput, , FCFactor)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@FCID", adInteger, adParamInput, , FCID)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.PasteInvoiceLine(PID,QtyFirm,QtySS,Price,DiscountRate,VATRate,lngQUID," & _
        "strRef,ExtraChargePID,ExtraChargeValue,FCPrice,FCFactor,FCID)", Array(PID, QtyFirm, QtySS, Price, _
         DiscountRate, VATRate, lngQUID, strRef, ExtraChargePID, ExtraChargeValue, FCPrice, FCFactor, FCID), _
         EA_NORERAISE
    
End Sub
Public Sub PasteGDNLine(PID As String, QtyFirm As Long, QtySS As Long, Price As Long, DiscountRate As Double, _
VATRate As Double, lngQUID As Long, strRef As String, ExtraChargePID As String, ExtraChargeValue As Long, _
FCPrice As Long, FCFactor As Double, FCID As Long)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "PasteInvoiceLine"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 50, PID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@QtyFirm", adInteger, adParamInput, , QtyFirm)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@QtySS", adInteger, adParamInput, , QtySS)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Price", adInteger, adParamInput, , Price)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Discountrate", adDouble, adParamInput, , DiscountRate)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@VATRate", adDouble, adParamInput, , VATRate)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@QUID", adInteger, adParamInput, , lngQUID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@REF", adVarChar, adParamInput, 30, strRef)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@EXTRACHARGEPID", adVarChar, adParamInput, 40, ExtraChargePID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ExtrachargeValue", adInteger, adParamInput, , ExtraChargeValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@FCPrice", adInteger, adParamInput, , FCPrice)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@FCFactor", adDouble, adParamInput, , FCFactor)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@FCID", adInteger, adParamInput, , FCID)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.PasteGDNLine(PID,QtyFirm,QtySS,Price,DiscountRate,VATRate,lngQUID,strRef," & _
        "ExtraChargePID,ExtraChargeValue,FCPrice,FCFactor,FCID)", Array(PID, QtyFirm, QtySS, Price, _
         DiscountRate, VATRate, lngQUID, strRef, ExtraChargePID, ExtraChargeValue, FCPrice, FCFactor, FCID), _
         EA_NORERAISE
End Sub

Public Sub PastePOLine(PID As String, QtyFirm As Long, QtySS As Long, Price As Long, DiscountRate As Double, _
VATRate As Double, lngQUID As Long, strRef As String, pETA As Date)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "PastePOLine"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 50, PID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@QtyFirm", adInteger, adParamInput, , QtyFirm)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@QtySS", adInteger, adParamInput, , QtySS)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Price", adInteger, adParamInput, , Price)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Discountrate", adDouble, adParamInput, , DiscountRate)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@VATRate", adDouble, adParamInput, , VATRate)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@QUID", adInteger, adParamInput, , lngQUID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@REF", adVarChar, adParamInput, 30, strRef)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ETA", adDate, adParamInput, , pETA)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.PastePOLine(PID,QtyFirm,QtySS,Price,DiscountRate,VATRate,lngQUID,strRef," & _
        "ExtraChargePID,ExtraChargeValue)", Array(PID, QtyFirm, QtySS, Price, DiscountRate, VATRate, lngQUID, _
         strRef)
End Sub

Public Function CreateNewOrder(pTPID As Long) As Long
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "[CreateCO]"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput, , pTPID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@ORDERNUM", adVarChar, adParamInput, 20, "")
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@MSG", adVarChar, adParamInput, 500, "")
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@NEWTRID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    iReturn = FNN(cmd.Parameters(3))
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    CreateNewOrder = iReturn
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CreateNewOrder(pTPID)", pTPID
End Function
Public Function CreateNewPOOrder(pTPID As Long) As Long
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "CreatePO"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput, , pTPID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@MSG", adVarChar, adParamInput, 500, "")
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@NEWTRID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    iReturn = FNN(cmd.Parameters(2))
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    CreateNewPOOrder = iReturn
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CreateNewPOOrder(pTPID)", pTPID
End Function
Public Function CreateNewTransfer(pSTID As Long, pInOut As String) As Long
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "CreateTFR"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@StoreID", adInteger, adParamInput, , pSTID)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@InOut", adVarChar, adParamInput, 5, pInOut)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@MSG", adVarChar, adParamInput, 500, "")
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@NEWTRID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    iReturn = FNN(cmd.Parameters(3))
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    CreateNewTransfer = iReturn
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CreateNewTransfer(pSTID)", pSTID
End Function
Public Function CreateNewTransferFromFile(pXMLDoc As String, pInOut As String, pMsg As String) As Long
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "CreateTFRFromFile"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@ARGXML", adVarChar, adParamInput, 20000, pXMLDoc)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@InOut", adVarChar, adParamInput, 5, pInOut)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@MSG", adVarChar, adParamInput, 500, pMsg)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@NEWTRID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    iReturn = FNN(cmd.Parameters(3))
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    CreateNewTransferFromFile = iReturn
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_StockManager.CreateNewTransferFromFile(pSTID)", pSTID
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_StockManager.CreateNewTransferFromFile(pStoreCode,pInOut,pMsg)", Array(pStoreCode, _
'         pInOut, pMsg)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CreateNewTransferFromFile(pXMLDoc,pInOut,pMsg)", Array(pXMLDoc, pInOut, _
         pMsg)
End Function
Public Sub appendTransferRows(pTFRID As Long, par() As String)
Dim i As Integer
    MsgBox par(0, 0)
End Sub
Public Sub PasteCOLine(PID As String, QtyFirm As Long, QtySS As Long, Price As Long, DiscountRate As Double, _
VATRate As Double, lngQUID As Long, strRef As String, ExtraChargePID As String, ExtraChargeValue As Long, pETA As Date, _
FCPrice As Long, FCFactor As Double, FCID As Long)
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "PasteCOLine"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 50, PID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@QtyFirm", adInteger, adParamInput, , QtyFirm)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@QtySS", adInteger, adParamInput, , QtySS)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Price", adInteger, adParamInput, , Price)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Discountrate", adDouble, adParamInput, , DiscountRate)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@VATRate", adDouble, adParamInput, , VATRate)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@QUID", adInteger, adParamInput, , lngQUID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@REF", adVarChar, adParamInput, 30, strRef)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@EXTRACHARGEPID", adVarChar, adParamInput, 40, ExtraChargePID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@EXTRACHARGEVALUE", adInteger, adParamInput, , ExtraChargeValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ETA", adDate, adParamInput, , pETA)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@FCPrice", adInteger, adParamInput, , FCPrice)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@FCFactor", adDouble, adParamInput, , FCFactor)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@FCID", adInteger, adParamInput, , FCID)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub

End Sub

Public Function CreateNewInvoice(pTPID As Long, IsProforma As Boolean, IsPreDelivery As Boolean) As Long
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "[CreateINV]"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput, , pTPID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@PF", adBoolean, adParamInput, , IsProforma)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@IsPreDelivery", adBoolean, adParamInput, , IsPreDelivery)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@NEWTRID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    iReturn = FNN(cmd.Parameters(3))
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    CreateNewInvoice = iReturn
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CreateNewInvoice(pTPID)", pTPID
End Function

'Public Sub PasteINVLine(PID As String, QtyFirm As Long, QtySS As Long, Price As Long, DiscountRate As Double, VATRate As Double, lngQUID As Long, strRef As String)
'Dim OpenResult As Integer
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'Dim f As String
'
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    Set cmd = New ADODB.Command
'    cmd.CommandText = "[PasteInvoiceLine]"
'    cmd.CommandType = adCmdStoredProc
'
'    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 50, PID)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@QtyFirm", adInteger, adParamInput, , QtyFirm)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@QtySS", adInteger, adParamInput, , QtySS)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@Price", adInteger, adParamInput, , Price)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@Discountrate", adDouble, adParamInput, , DiscountRate)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@VATRate", adDouble, adParamInput, , VATRate)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@QUID", adInteger, adParamInput, , lngQUID)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@REF", adVarChar, adParamInput, 30, strRef)
'    cmd.Parameters.Append par
'
'    cmd.ActiveConnection = oPC.COShort
'    cmd.execute
'
'    Set cmd = Nothing
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'
'    Exit Sub
'
'End Sub

'=======
Public Function CreateNewQuotation(pTPID As Long) As Long
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "[CreateQuotation]"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput, , pTPID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@NEWTRID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    iReturn = FNN(cmd.Parameters(1))
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    CreateNewQuotation = iReturn
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CreateNewQuotation(pTPID)", pTPID
End Function
Public Function CreateNewAppro(pTPID As Long) As Long
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "[CreateAppro]"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput, , pTPID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@NEWTRID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    iReturn = FNN(cmd.Parameters(1))
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    CreateNewAppro = iReturn
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CreateNewAppro(pTPID)", pTPID
End Function

'Public Sub PasteQULine(PID As String, QtyFirm As Long, QtySS As Long, Price As Long, DiscountRate As Double, VATRate As Double, lngQUID As Long, strRef As String)
'    On Error GoTo errHandler
'Dim OpenResult As Integer
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'Dim f As String
'
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    Set cmd = New ADODB.Command
'    cmd.CommandText = "PasteQULine"
'    cmd.CommandType = adCmdStoredProc
'
'    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 50, PID)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@QtyFirm", adInteger, adParamInput, , QtyFirm)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@QtySS", adInteger, adParamInput, , QtySS)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@Price", adInteger, adParamInput, , Price)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@Discountrate", adDouble, adParamInput, , DiscountRate)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@VATRate", adDouble, adParamInput, , VATRate)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@QUID", adInteger, adParamInput, , lngQUID)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@REF", adVarChar, adParamInput, 30, strRef)
'    cmd.Parameters.Append par
'
'    cmd.ActiveConnection = oPC.COShort
'    cmd.execute
'
'    Set cmd = Nothing
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_StockManager.PasteQULine(PID,QtyFirm,QtySS,Price,DiscountRate,VATRate,lngQUID,strRef)", _
'         Array(PID, QtyFirm, QtySS, Price, DiscountRate, VATRate, lngQUID, strRef)
'End Sub
'
'''Public Function ExportToXML_CustomerAdvice(rsPO As adodb.Recordset, DispatchMode As enTransmitType, pFilename As String, Optional pDestinationEmailAddress As String, Optional pWholeMessage As String, Optional pQtyCopies As Integer) As Boolean
'''    On Error GoTo ErrHandler
'''Dim oTF As New z_TextFile
'''Dim strPath As String
'''Dim strBillto As String
'''Dim strDelto As String
'''Dim strFOFile As String
'''Dim strPDFFile As String
'''Dim strXML As String
'''Dim strCommand As String
'''Dim i As Integer
'''Dim strHTML As String
'''Dim fs As New FileSystemObject
'''Dim objXSL As New MSXML2.DOMDocument60
'''Dim opXMLDOC As New MSXML2.DOMDocument60
'''Dim objXMLDOC  As New MSXML2.DOMDocument60
'''Dim strWorkingFolder As String
'''Dim strDispatchMethod As String
'''Dim oDC As a_DocumentControl
'''Dim strExecutable As String
'''
'''    If DispatchMode = enMail Then
'''        If oPC.UsesOutlookForPOEmail Then
'''            strWorkingFolder = oPC.LocalFolder & "TEMP\"
'''            strDispatchMethod = "Outlook"
'''        Else
'''            strWorkingFolder = oPC.SharedFolderRoot & "\Emails\"
'''            strDispatchMethod = "Emailing"
'''        End If
'''    ElseIf DispatchMode = enPrint Then
'''        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
'''        strDispatchMethod = "Printing"
'''    ElseIf DispatchMode = enView Then
'''        strWorkingFolder = oPC.SharedFolderRoot & "\PDF\"
'''        strDispatchMethod = ""
'''    Else
'''        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
'''        strDispatchMethod = "Printing"
'''    End If
'''                        p 1
'''    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
'''    If oDC Is Nothing And strDispatchMethod <> "Emailing" And strDispatchMethod <> "Outlook" Then
'''        ExportToXML = False
'''        Exit Function
'''    End If
'''                        p 2
'''    If pQtyCopies < 1 Then pQtyCopies = 1
'''
'''    Set xMLDoc = New ujXML
'''                        p 21
'''    With xMLDoc
'''        .docProgID = "MSXML2.DOMDocument"
'''        .docInit "COA_DOC"
'''            .chCreate "MessageType"
'''                .elText = "SALESORDERADVICE"
'''            .elCreateSibling "MessageCreationDate"
'''                .elText = Format(Now(), "yyyymmddHHNN")
'''            .elCreateSibling "DestinationAddress"
'''                If oPC.TestMode Then
'''                    .elText = oPC.EmailAddressForTesting
'''                    pDestinationEmailAddress = oPC.EmailAddressForTesting
'''                Else
'''                    If Me.Customer Is Nothing Then
'''                        .elText = ""
'''                        pDestinationEmailAddress = ""
'''                    Else
'''                        .elText = Me.OrderToAddress.Email
'''                        pDestinationEmailAddress = Me.Supplier.OrderToAddress.Email
'''                    End If
'''                End If
'''                        p 23
'''            .elCreateSibling "TemplateName"
'''                .elText = "PO_DOC"
'''            .elCreateSibling "SendersEmail"
'''                If oPC.EmailFrom > "" Then
'''                    .elText = oPC.EmailFrom
'''                Else
'''                    .elText = Me.SendersEmail
'''                End If
'''            .elCreateSibling "CopyCount"
'''                .elText = pQtyCopies
'''            .elCreateSibling "Printer"
'''                If Not oDC Is Nothing Then .elText = oDC.PrinterName
'''            .elCreateSibling "Status"
'''                .elText = StatusForPrinting
'''            .elCreateSibling "DocCode"
'''                .elText = Me.DOCCode
'''            .elCreateSibling "DocDate", True
'''                .elText = Me.DOCDate
'''            .elCreateSibling "Sender", True
'''                .elText = oPC.Configuration.DefaultCompany.CompanyName
'''            .elCreateSibling "SenderAddress", True
'''                .elText = Replace(oPC.Configuration.DefaultCompany.StreetAddress, Chr(13) & Chr(10), Chr(10))
'''            .elCreateSibling "SupplierName", True
'''                .elText = Supplier.NameAndCode(35)
'''            .elCreateSibling "SupplierWithAddress", True
'''                If Supplier.OrderToAddress Is Nothing Then
'''                    .elText = ""
'''                Else
'''                    .elText = Replace(Supplier.OrderToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
'''                End If
'''            .elCreateSibling "SupplierPhone", True
'''                If Supplier.BillTOAddress Is Nothing Then
'''                    .elText = ""
'''                Else
'''                    .elText = IIf(Me.Supplier.BillTOAddress.Phone > "", "Phone: " & Supplier.BillTOAddress.Phone, "")
'''                End If
'''            .elCreateSibling "SupplierFax", True
'''                If Supplier.BillTOAddress Is Nothing Then
'''                    .elText = ""
'''                Else
'''                    .elText = IIf(Supplier.BillTOAddress.Fax > "", "Fax: " & Supplier.BillTOAddress.Fax, "")
'''                End If
'''            .elCreateSibling "ACNO"
'''                .elText = IIf(Me.Supplier.AcNo > "", "Ac/no. " & Me.Supplier.AcNo, "")
'''            .elCreateSibling "BillTo", True
'''                If Me.BillToAddressID > 0 Then
'''                    .elText = Replace(oPC.Configuration.Stores.FindStoreByID(Me.BillToAddressID).BillAddress, Chr(13) & Chr(10), Chr(10))
'''                End If
'''            .elCreateSibling "DelTo", True
'''                If Me.DELTOStoreID > 0 Then
'''                    .elText = Replace(oPC.Configuration.Stores.FindStoreByID(Me.DELTOStoreID).DelAddress, Chr(13) & Chr(10), Chr(10))
'''                End If
'''                            p 24
'''
'''            For i = 1 To Me.POLines.Count
'''               ' If POLines(i).Fulfilled <> "CAN" Then
'''                    .elCreateSibling "DetailLine", True
'''                    .chCreate "SKU"
'''                    .elText = POLines(i).ProductCodeForExport
'''                    .elCreateSibling "Title", True
'''                    If POLines(i).Fulfilled <> "CAN" Then
'''                       .elText = POLines(i).TitleAuthor
'''                    Else
'''                       .elText = "***CANCELLED***" & POLines(i).TitleAuthor
'''                    End If
'''                    .elCreateSibling "QtyFirm", True
'''                        .elText = POLines(i).QtyFirmF
'''                    .elCreateSibling "QtySS", True
'''                        .elText = POLines(i).QtySSF
'''                    .elCreateSibling "Price", True
'''                        .elText = POLines(i).PriceF(bForeign)
'''                    .elCreateSibling "DiscountRate", True
'''                        .elText = POLines(i).DiscountF
'''                    .elCreateSibling "Reference", True
'''                        .elText = POLines(i).Ref
'''                    .elCreateSibling "Extension", True
'''                        .elText = POLines(i).PLessDiscExtF(bForeign)
'''                    .elCreateSibling "Note", True
'''                        .elText = POLines(i).Note
'''                    .navUP
'''                'End If
'''            Next i
'''                            p 25
'''            .elCreateSibling "TotalText", True
'''                .elText = "Total"
'''            .elCreateSibling "TotalNumbers", True
'''                .elText = TotalPayableF(bForeign)
'''            .elCreateSibling "Memo", True
'''                .elText = Memo
'''            .elCreateSibling "CompanyRegistration", True
'''                .elText = oPC.Configuration.DefaultCompany.CoRegistrationNumber
'''            .elCreateSibling "VATNumber", True
'''                .elText = oPC.Configuration.DefaultCompany.VatNumber
'''            .elCreateSibling "StaffMember", True
'''                .elText = IIf(Me.Signature > "", Me.Signature, "")
'''            .elCreateSibling "OrderMessage", True
'''                .elText = oPC.Configuration.OrderText
'''    End With
'''                            p 26
''''FINALLY PRODUCE THE .XML FILE
'''    strXML = strWorkingFolder & "PO_" & Me.DOCCode & ".xml"
'''    With xMLDoc
'''        If fs.FileExists(strXML) Then
'''            fs.DeleteFile strXML
'''        End If
'''        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
'''    End With
'''                            p 27
'''    If DispatchMode = enMail Then
'''        'WRITE THE .HTML FILE
'''            objXSL.async = False
'''            objXSL.validateOnParse = False
'''            objXSL.resolveExternals = False
'''            strPath = oPC.SharedFolderRoot & "\Templates\PO_DOC_HTML.xslt"
'''            Set fs = New FileSystemObject
'''            If fs.FileExists(strPath) Then
'''                objXSL.Load strPath
'''            End If
'''
'''            If fs.FileExists(strWorkingFolder & "PO_" & Me.DOCCode & ".HTML") Then
'''                fs.DeleteFile strWorkingFolder & "PO_" & Me.DOCCode & ".HTML", True
'''            End If
'''            oTF.OpenTextFileToAppend strWorkingFolder & "PO_" & Me.DOCCode & ".HTML"
'''            pWholeMessage = xMLDoc.docObject.transformNode(objXSL)
'''            oTF.WriteToTextFile pWholeMessage 'xMLDoc.docObject.transformNode(objXSL)
'''            oTF.CloseTextFile
'''    End If
'''
'''                        p 27
''''WRITE THE .PDF FILE IF NECESSARY
''''Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file
'''    If (Me.Supplier.DispatchMethod = "M" And DispatchMode = enMail) Or DispatchMode = enView Then
'''        Set objXSL = Nothing
'''        Set objXSL = New MSXML2.DOMDocument60
'''        objXSL.async = False
'''        objXSL.validateOnParse = False
'''        objXSL.resolveExternals = False
'''        strPath = oPC.SharedFolderRoot & "\Templates\PO_DOC_FO.xsl"
'''        Set fs = New FileSystemObject
'''        If fs.FileExists(strPath) Then
'''            objXSL.Load strPath
'''        End If
'''                        p 6
'''        Set opXMLDOC = New MSXML2.DOMDocument60
'''        opXMLDOC.async = False
'''        opXMLDOC.validateOnParse = False
'''        opXMLDOC.resolveExternals = False
'''        xMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC
'''
'''        strFOFile = strWorkingFolder & "PO_" & Me.DOCCode & ".FO"
'''        strPDFFile = strWorkingFolder & "PO_" & Me.DOCCode & ".PDF"
'''                                p 7
'''        docWriteTostream strFOFile, opXMLDOC, "UNICODE"
'''
''''Stage 2 Convert the .FO file to .PDF and clean up
'''        strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFile & " " & strPDFFile
'''        'ShellandWait strCommand, 100
'''        F_7_AB_1_ShellAndWaitSimple strCommand
'''        If fs.FileExists(strFOFile) Then
'''            fs.DeleteFile strFOFile
'''        End If
'''    End If
'''                        p 8
''''Log the dispatch
'''    Set oSM = New z_StockManager
'''    Set oFSO = New FileSystemObject
'''    pFilename = strPDFFile
'''
'''    If DispatchMode <> enMail And DispatchMode <> enView Then
'''        oSM.LogTransmission Me.TRID, "Dispatched: by " & strDispatchMethod & " : " & Format(Date, "dd/mm/yyyy") & vbCrLf
'''        Log = "Dispatched: by " & strDispatchMethod & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
'''    End If
'''    Set oSM = Nothing
'''    Set oFSO = Nothing
'''                        p 9
'''    oTF.CloseTextFile
'''    ExportToXML = True
'''
'''    Screen.MousePointer = vbDefault
'''    If DispatchMode = enView Then
'''        strExecutable = GetPDFExecutable(strPDFFile)
'''        Shell strExecutable & " " & strPDFFile
'''    End If
'''
'''    Exit Function
'''ErrHandler:
'''    If ErrMustStop Then Debug.Assert False: Resume
'''    ErrorIn "a_PO.ExportToXML(bForeign)", bForeign, , , "strPos", Array(strPos)
'''End Function
'''Public Sub docWriteTostream(ByVal FilePath As String, obj As MSXML2.DOMDocument60, _
'''                Optional ByVal CharSet As String = "UNICODE")
'''    On Error GoTo ErrHandler
'''    Dim s As Object
'''    Set s = CreateObject("ADODB.Stream")
'''    With s
'''        If CharSet <> "" Then .CharSet = CharSet
'''        .Open
'''        .WriteText obj.xml
'''        .SaveToFile FilePath, 2 'adSaveCreateOverWrite
'''        .Close
'''    End With
'''    Exit Sub
'''ErrHandler:
'''    ErrorIn "a_PO.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
'''End Sub
Public Sub SavePurchaseOrdemindeport(x As XArrayDB, WSName As String)
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim i As Long

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "DELETE FROM tReminder WHERE  WSNAME = '" & WSName & "'"
    For i = 1 To x.UpperBound(1)
        If (x.Value(i, 14) = "R") Then
            If rs Is Nothing Then
                Set rs = New ADODB.Recordset
                rs.Open "SELECT * FROM tReminder", oPC.COShort, adOpenDynamic, adLockOptimistic
            End If
            rs.AddNew
            rs.Fields("POLID") = x(i, 12)
            rs.Fields("WSNAME") = WSName
            rs.Update
        End If
    Next i
    If Not rs Is Nothing Then
        If rs.State <> 0 Then
            rs.Close
        End If
        Set rs = Nothing
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------


End Sub

Public Sub SaveCustomerOrderStatusReport(x As XArrayDB, WSName As String)
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim i As Long

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "DELETE FROM tCustReport WHERE REPTYPE = 'G' AND WSNAME = '" & WSName & "'"
    For i = 1 To x.UpperBound(1)
        If (x.Value(i, 11) > "") Then
'        If X(i, 12) > 0 Then
            If rs Is Nothing Then
                Set rs = New ADODB.Recordset
                rs.Open "SELECT * FROM tCustReport", oPC.COShort, adOpenDynamic, adLockOptimistic
            End If
            rs.AddNew
            rs.Fields("COLID") = x(i, 12)
            rs.Fields("WSNAME") = WSName
            rs.Fields("RePType") = "G"
            rs.Update
        End If
    Next i
    If Not rs Is Nothing Then
        If rs.State <> 0 Then
            rs.Close
        End If
        Set rs = Nothing
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------


End Sub
Public Sub PrintCustomerOrderStatusReport(WSName As String)
Dim ar As arCustReport
Dim rs As ADODB.Recordset
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.Open "Select * FROM vODCOReport WHERE WSNAME = '" & WSName & "'", oPC.COShort, adOpenForwardOnly
    Screen.MousePointer = vbDefault
    If rs.RecordCount > 0 Then
        Set ar = New arCustReport
        ar.Printer.Orientation = ddOLandscape
        ar.component rs  ', frm.chkPagePerSupplier
        ar.Show vbModal
    End If
    rs.Close
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Sub
Public Sub PrintPurchaseOrderReminderReport(WSName As String, bPagePerSupplier As Boolean)
Dim ar As arPOReminder
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.Open "Select * FROM vODPOReport WHERE WSNAME = '" & WSName & "'", oPC.COShort, adOpenForwardOnly
    Screen.MousePointer = vbDefault
    If rs.RecordCount > 0 Then
        Set ar = New arPOReminder
        ar.Printer.Orientation = ddOPortrait
        ar.component rs, bPagePerSupplier
        ar.Show vbModal
    End If
    rs.Close
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Sub

Public Sub ActionODCO(pRs As ADODB.Recordset)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If pRs.eof Then Exit Sub
    Set rs = pRs
    rs.MoveFirst
    Do While Not rs.eof
        If FNS(rs!ACT2) > "" Or FNS(rs!ACT1) > "" Then
            oPC.COShort.execute "INSERT INTO tCOLA (COLA_COL_ID,COLA_DATE,COLA_REPORT) VALUES (" & rs!COLID & ",'" & ReverseDateTime(Now()) & "','" & FNS(rs!ACT2) & "')"
        End If
        Select Case Right(rs!ACT1, 1)
        Case "W"
            oPC.COShort.execute "UPDATE tCOL SET COL_ETA = '" & ReverseDate(DateAdd("ww", Left(rs!ACT1, 1), Date)) & "', COL_LASTACTION = '" & FNS(rs!ACT2) & "',COL_LASTACTIONDATE = '" & ReverseDateTime(Now()) & "' WHERE COL_ID = " & CStr(rs!COLID)
        Case "M"
            oPC.COShort.execute "UPDATE tCOL SET COL_ETA = '" & ReverseDate(DateAdd("m", Left(rs!ACT1, 1), Date)) & "', COL_LASTACTION = '" & FNS(rs!ACT2) & "',COL_LASTACTIONDATE = '" & ReverseDateTime(Now()) & "' WHERE COL_ID = " & CStr(rs!COLID)
        Case Else
            oPC.COShort.execute "UPDATE tCOL SET  COL_LASTACTION = '" & FNS(rs!ACT2) & "',COL_LASTACTIONDATE = '" & ReverseDateTime(Now()) & "' WHERE COL_ID = " & CStr(rs!COLID)
        
        End Select
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.ActionODCO(pRs)", pRs
End Sub
Public Sub ActionODPO(pRs As ADODB.Recordset)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim dteDiarize As Date
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = pRs
    If rs.eof Then Exit Sub
    rs.MoveFirst
    oPC.COShort.execute "DELETE FROM tReminder WHERE  WSNAME = '" & oPC.WorkstationName & "'"
    Do While Not rs.eof
        If FNS(rs.Fields(5)) > "" Or FNS(rs.Fields(3)) > "" Then
            oPC.COShort.execute "INSERT INTO tPOLA (POLA_POL_ID,POLA_DATE,POLA_CODE,POLA_REPORT) VALUES (" & rs.Fields(0) & ",'" & ReverseDateTime(Date) & "','" & FNS(rs.Fields(3)) & "','" & Replace(FNS(rs.Fields(5)), "'", "''") & "')"
        End If
        '''3
        Select Case Right(FNS(rs.Fields(3)), 1)
            Case "W"
                dteDiarize = DateAdd("ww", Left(rs.Fields(3), 1), Date)
            Case "M"
                dteDiarize = DateAdd("m", Left(rs.Fields(3), 1), Date)
        End Select
        
        Select Case FNS(rs.Fields(2))
            Case "C"
                'rs!POL_Fulfilled = "CAN"
                oPC.COShort.execute "UPDATE tPOL SET POL_FULFILLED = 'CAN' WHERE POL_ID = " & rs.Fields("POLID")
                oPC.COShort.execute "UPDATE tPRODUCT SET P_QTYOnOrder =  dbo.NonNegative(P_QTYOnOrder - ISNULL(POL_QtySS,0) - ISNULL(POL_QtyFirm,0)) FROM tPRODUCT JOIN tPOL ON POL_P_ID = P_ID WHERE P_ID = '" & FNS(rs.Fields("PID")) & "'"
              '  oPC.COShort.execute "UPDATE tStoreP SET STP_QTYOnOrder =  dbo.NonNegative(STP_QTYOnOrder - " & FNN(rs!POL_QtySS) + FNN(rs!POL_QtyFirm) - FNN(rs!POL_QTYReceivedSoFar) & ") WHERE STP_P_ID = '" & rs!POL_P_ID & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
            Case "R"
                oPC.COShort.execute "INSERT INTO tReminder (POLID,WSNAME) VALUES(" & rs.Fields("POLID") & ",'" & oPC.WorkstationName & "')"
        End Select
        
        Select Case FNS(rs.Fields(1))
            Case "O"
                oPC.COShort.execute "UPDATE tPRODUCT SET P_Status =  'O' WHERE P_ID = '" & FNS(rs.Fields("PID")) & "'"
            Case "R"
                oPC.COShort.execute "UPDATE tPRODUCT SET P_Status =  'R' WHERE P_ID = '" & FNS(rs.Fields("PID")) & "'"
        End Select
        
            oPC.COShort.execute "UPDATE tPOL SET POL_ETA = '" & ReverseDate(dteDiarize) & "', POL_LASTACTION = '" & Replace(FNS(rs.Fields("Note")), "'", "''") & "',POL_LASTACTIONDATE = '" & ReverseDateTime(Now()) & "' WHERE POL_ID = " & CStr(rs!POL_ID)
           ' rs.Update
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.ActionODPO(pRs)", pRs
End Sub
Public Function CreateInvoiceGDNFromApp(Strguid As String, pTPID As Long, pTRID As Long) As Boolean
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "CreateInvoiceGDNFromAppro"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Strguid", adVarChar, adParamInput, 50, Strguid)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput, , pTPID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@GDNID", adInteger, adParamOutput, , pTRID)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    pTRID = CLng(cmd.Parameters(2))
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CreateInvoiceGDNFromAppR(Strguid,pTPID)", Array(Strguid, pTPID)
End Function
Public Function CreateReturn_FromAppro(Strguid As String, pTRSTATUS As Long, pTPID As Long, pAPPRID As Long, pSTAFFID As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    Set cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "CreateReturn_FromAppro"
    cmd.commandType = adCmdStoredProc


    Set prm = cmd.CreateParameter("@Strguid", adVarChar, adParamInput, 50, Strguid)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TRSTATUS", adChar, adParamInput, 1, pTRSTATUS)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@STAFFID", adInteger, adParamInput, , pSTAFFID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput, , pTPID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@ARID", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    cmd.execute
    
    pAPPRID = cmd.Parameters("@ARID")
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GenerateTRs.GenerateReturn_FromAppro(pTRSTATUS,pAPPRID)", Array(pTRSTATUS, pAPPRID)
End Function

Public Function ActionODPOL(Strguid As String, lngPaid As Long) As Boolean
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "Action_POL_XML"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Strguid", adVarChar, adParamInput, 50, Strguid)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@PAID", adInteger, adParamOutput, , lngPaid)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.ActionODPOL(Strguid,lngPAID)", Array(Strguid, lngPaid)
End Function
Public Function Action_InsertPayments(Strguid As String, lngPaid As Long) As Boolean
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "Action_InsertPayments"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Strguid", adVarChar, adParamInput, 50, Strguid)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@PAID", adInteger, adParamOutput, , lngPaid)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.Action_InsertPayments(Strguid,lngPAID)", Array(Strguid, lngPaid)
End Function

Public Function ActionODCOL(PXMLArgs As String, lngPaid As Long) As Boolean
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
    errSysHandlerSet
    
    ActionODCOL = True

    If Len(PXMLArgs) > 40000 Then
        ActionODCOL = False
        Exit Function
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "Action_COL_XML"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@ARGXML", adVarChar, adParamInput, 40000, PXMLArgs)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@PAID", adInteger, adParamOutput, , lngPaid)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmProductPrev: LoadControls, err repeat = " & CStr(errRepeat) & ", line:" & CStr(Erl())
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmProductPrev: LoadControls after 5 re-attempts"
            MsgBox "Memory error in z_StockManager:ActionODCOL. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't run search."
            Err.Clear
            Exit Function
        End If
    End If

    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.ActionODCOL(pXMLArgs)", PXMLArgs
End Function

Public Sub ActionODPO_NoteandDiary(pcPOL As ch_POL, strDiarize As String, strNote As String)
    On Error GoTo errHandler
Dim dteDiarize As Date
Dim OpenResult As Integer
Dim oPOL As a_POL

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "EXECUTE TRACKINGACTION_1"
    For Each oPOL In pcPOL
        Select Case UCase(Right(strDiarize, 1))
        Case "W"
            dteDiarize = DateAdd("ww", Left(strDiarize, 1), Date)
        Case "M"
            dteDiarize = DateAdd("m", Left(strDiarize, 1), Date)
        End Select
         If FNS(strNote) > "" Then
            oPC.COShort.execute "INSERT INTO tPOLAction (POLA_POL_ID,POLA_NewETA,POLA_REPORT) " _
               & " VALUES (" & oPOL.POLID & ",'" & ReverseDate(dteDiarize) & "','" & strNote & "')"
        End If
       oPC.COShort.execute "UPDATE tPOL SET POL_ETA = '" & ReverseDate(dteDiarize) & "', POL_LASTACTION = '" & strDiarize & "',POL_LASTACTIONDATE = '" & ReverseDateTime(Now()) & "' WHERE POL_ID = " & CStr(oPOL.POLID)
    Next
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.ActionODPO_NoteandDiary(pcPOL,strDiarize,strNote)", Array(pcPOL, _
         strDiarize, strNote)
End Sub

Public Sub MarkForWebExport(PID As String)
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "MarkForWebExport"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 50, PID)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------


    Exit Sub

End Sub

Public Sub SetAddressImmediate(pBillOrDel As String, pTRID As Long, pAddressID As Long)
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    If pBillOrDel = "BILL" Then
        cmd.CommandText = "UPDATE tTR SET TR_ADDRESSDOCID = " & pAddressID & " WHERE TR_ID = " & pTRID
    ElseIf pBillOrDel = "DEL" Then
        cmd.CommandText = "UPDATE tTR SET TR_ADDRESSGOODSID = " & pAddressID & " WHERE TR_ID = " & pTRID
    End If
    cmd.commandType = adCmdText
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager_P.SetAddressImmediate(pAddressID)", pAddressID
End Sub


Public Sub Action_Prod(WSName As String, NewDiaryDate As Date, SupplierMessage As String, CustomerMessage As String, _
PID As String, CANCELPOLs As Boolean, CANCELCOLs As Boolean, NEWPRODUCTSTATUS As String, lngPaid As Long, strSignature As String)
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
Dim i As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "Action_Prod"
    cmd.commandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.COShort
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@WSName", adVarChar, adParamInput, 50, WSName)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@NewDiaryDate", adDate, adParamInput, , NewDiaryDate)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@SupplierMessage", adVarChar, adParamInput, 150, SupplierMessage)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@CustomerMessage", adVarChar, adParamInput, 300, CustomerMessage)
    cmd.Parameters.Append par
       
    Set par = Nothing
    Set par = cmd.CreateParameter("@PID", adGUID, adParamInput, , PID)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@CANCELPOLs", adBoolean, adParamInput, , CANCELPOLs)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@CANCELCOLs", adBoolean, adParamInput, , CANCELCOLs)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@NEWPRODUCTSTATUS", adInteger, adParamInput, , NEWPRODUCTSTATUS)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@Signature", adVarChar, adParamInput, 100, strSignature)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@PAID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.execute

    lngPaid = FNN(cmd.Parameters("@PAID"))
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    ErrPreserve
    If Err = -2147467259 Then   'networkfailure
        OpenResult = oPC.OpenDBSHort(True)
        Resume
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.Action_Prod(NewDiaryDate,SupplierMessage,CustomerMessage,PID,CANCELPOLs," & _
        "CANCELCOLs,NEWPRODUCTSTATUS,lngPAID)", Array(NewDiaryDate, SupplierMessage, CustomerMessage, PID, _
         CANCELPOLs, CANCELCOLs, NEWPRODUCTSTATUS, lngPaid)
End Sub


Public Sub PrepareCustListForCOLReport(PAID As Long)
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
Dim i As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
'    Set cmd = New ADODB.Command
'    cmd.CommandText = "CreateLocalCOLReportTable"
'    cmd.CommandType = adCmdStoredProc
'    cmd.ActiveConnection = oPC.COShort
'    Set par = cmd.CreateParameter("@ARG", adVarChar, adParamInput, 50, Replace(oPC.WorkstationName, "-", ""))
'    cmd.Parameters.Append par
'
'    cmd.execute
'    Set cmd = Nothing
'
'
'    For i = 1 To X.UpperBound(1)
'        Set cmd = New ADODB.Command
'        cmd.CommandText = "PrepareCustListForCOLReport"
'        cmd.CommandType = adCmdStoredProc
'        cmd.ActiveConnection = oPC.COShort
'        Set par = cmd.CreateParameter("@PAID", adInteger, adParamInput, , PAID)
'        cmd.Parameters.Append par
''        Set par = cmd.CreateParameter("@ARG", adVarChar, adParamInput, 50, Replace(oPC.WorkstationName, "-", ""))
''        cmd.Parameters.Append par
''        Set par = cmd.CreateParameter("@COLID", adInteger, adParamInput, , FNN(X(i, 11)))
''        cmd.Parameters.Append par
''        Set par = cmd.CreateParameter("@EmailAddress", adVarChar, adParamInput, 50, X(i, 7))
''        cmd.Parameters.Append par
''        Set par = cmd.CreateParameter("@DispatchMethod", adVarChar, adParamInput, 50, X(i, 8))
''        cmd.Parameters.Append par
''        Set par = cmd.CreateParameter("@Msg", adVarChar, adParamInput, 500, strMsg)
''        cmd.Parameters.Append par
'        cmd.execute
'    Next
'
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.PrepareCustListForCOLReport(PAID)", PAID
End Sub

Public Function GetBOReportData(PAID As Long, pSupplierMessage As String, pCustomerMessage As String, pSignature As String) As ADODB.Recordset
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String
Dim rs As ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.commandType = adCmdText
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "COLReportData"
    cmd.commandType = adCmdStoredProc
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@PAID", adInteger, adParamInput, , PAID)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@SupplierMessage", adVarChar, adParamOutput, 150, pSupplierMessage)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@CustomerMessage", adVarChar, adParamOutput, 300, pCustomerMessage)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@Signature", adVarChar, adParamOutput, 100, pSignature)
    cmd.Parameters.Append par

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic
    pSupplierMessage = FNS(cmd.Parameters(1))
    pCustomerMessage = FNS(cmd.Parameters(2))
    pSignature = FNS(cmd.Parameters(3))
    Set rs.ActiveConnection = Nothing
    Set cmd = Nothing
    Set GetBOReportData = rs

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function
'Public Function GetHistPOLActions(POLID As Long) As ADODB.Recordset
'Dim OpenResult As Integer
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'Dim f As String
'Dim rs As ADODB.Recordset
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    Set cmd = New ADODB.Command
'    cmd.CommandType = adCmdText
'    cmd.CommandTimeout = 0
'    cmd.ActiveConnection = oPC.COShort
'    cmd.CommandText = "COLReportData"
'    cmd.CommandType = adCmdStoredProc
'
'    Set par = Nothing
'    Set par = cmd.CreateParameter("@POLID", adInteger, adParamInput, , POLID)
'    cmd.Parameters.Append par
'
'
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open cmd, , adOpenStatic
'    Set rs.ActiveConnection = Nothing
'    Set cmd = Nothing
'    Set GetHistPOLActions = rs
'
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'End Function

Public Function GenerateReminders(pPAID As Long) As Boolean
    On Error GoTo errHandler
Dim Res As Boolean
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String
Dim objXSL As New MSXML2.DOMDocument60
Dim opXMLDOC As New MSXML2.DOMDocument60
Dim objXMLDOC  As New MSXML2.DOMDocument60
Dim xMLDoc As ujXML
Dim XMLArgs As String
Dim strXMLFilePath As String
Dim strFOFilePath As String
Dim strPDFFilePath As String
Dim strAcno As String
Dim strTmpAcno As String
Dim fs As New FileSystemObject
Dim strCommand As String
Dim cnt As Long
Dim strExecutable As String
Dim strTemplatePath As String
Dim lngTPID As Long
Dim lngTmpTPID As Long

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    GenerateReminders = True
    
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "SELECT * FROM vODPOReport_XML WHERE PAID = " & CStr(pPAID) & " And POLA_NeedReminder <> 0 ORDER BY SupplierACNo,DocCode,TITLE"
    cmd.commandType = adCmdText
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic
    rs.ActiveConnection = Nothing
     If rs.eof Then
        rs.Close
        Set rs = Nothing
        GenerateReminders = False
        Exit Function
    End If
   
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Do While Not rs.eof
        p 2, "topOfLoop"
    '    MsgBox "Top Of Loop"
        strXMLFilePath = oPC.LocalFolder & "Reminders\" & "POReminder_" & StripToAlphanumeric(FNS(rs.Fields("SupplierName"))) & "_" & Format(Date, "MMDDHHmm") & ".xml"
        strFOFilePath = oPC.LocalFolder & "Reminders\" & "POReminder_" & StripToAlphanumeric(FNS(rs.Fields("SupplierName"))) & "_" & Format(Date, "MMDDHHmm") & ".FO"
        strPDFFilePath = oPC.LocalFolder & "Reminders\" & "POReminder_" & StripToAlphanumeric(FNS(rs.Fields("SupplierName"))) & "_" & Format(Date, "MMDDHHmm") & ".PDF"
        Set xMLDoc = New ujXML
        With xMLDoc
            .docProgID = "MSXML2.DOMDocument"
            .docInit "doc_PO_Reminder"
                .chCreate "MessageType"
                    .elText = "PO_Reminder"
                .elCreateSibling "MessageCreationDate"
                    .elText = Format(Now(), "dd-mm-yyyy HH:NN")
                .elCreateSibling "DestinationAddress"
                    .elText = FNS(rs.Fields("Email"))
                .elCreateSibling "TemplateName"
                    .elText = "POReminder_DOC"
                .elCreateSibling "SendersEmail"
                    If oPC.EmailFrom > "" Then
                        .elText = oPC.EmailFrom
                    Else
                        .elText = FNS(rs.Fields("StaffEmail"))
                    End If
                .elCreateSibling "LogoPath"
                If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
                Else
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
                End If
'                .elCreateSibling "DestinationAddress"
'                    .elText = FNS(rs.Fields("Email"))
                .elCreateSibling "Sender"
                    .elText = oPC.Configuration.DefaultCompany.CompanyName
                .elCreateSibling "SenderAddress"
                    .elText = oPC.Configuration.DefaultCompany.StreetAddress
                .elCreateSibling "SupplierName"
                    .elText = FNS(rs.Fields("SupplierName"))
                .elCreateSibling "SupplierWithAddress"
                    .elText = FNS(rs.Fields("SupplierAddress"))
                .elCreateSibling "SupplierPhone"
                    .elText = FNS(rs.Fields("Phone"))
                .elCreateSibling "SupplierFax"
                    .elText = FNS(rs.Fields("Fax"))
                .elCreateSibling "ACNo"
                    .elText = FNS(rs.Fields("SupplierACNo"))
            '    .elCreateSibling "DetailLine", True
                strAcno = FNS(rs.Fields("SupplierACNo"))
                lngTPID = FNN(rs.Fields("TP_ID"))
                lngTmpTPID = lngTPID
                Do While lngTmpTPID = lngTPID And lngTPID <> 0
                             p 3, "topOfInnerLoop"
                            ' MsgBox "TopOfInnerLoop"
                        .elCreateSibling "DetailLine", True
                            .chCreate "DocCode"
                                .elText = FNS(rs.Fields("DocCode"))
                            .elCreateSibling "DocDate", True
                                .elText = FNS(rs.Fields("DocDate"))
                            .elCreateSibling "SKU", True
                                .elText = FNS(rs.Fields("CODEF"))
                            .elCreateSibling "Title", True
                                .elText = FNS(rs.Fields("Title"))
                            .elCreateSibling "Ordered", True
                                .elText = FNS(rs.Fields("Qty"))
                            .elCreateSibling "Outstanding", True
                                .elText = FNS(rs.Fields("QtyOS"))
                            .elCreateSibling "POLReference", True
                                .elText = FNS(rs.Fields("POLRef"))
                            .elCreateSibling "StaffMember", True
                                .elText = FNS(rs.Fields("StaffMember"))
                            .elCreateSibling "Cancelled", True
                                .elText = FNS(rs.Fields("POLA_IsCancelled"))
                        .navUP
                        rs.MoveNext
                        If rs.eof Then
                            lngTPID = 0
                        Else
                            lngTPID = FNN(rs.Fields("TP_ID"))
                        End If
                Loop
    
        'FINALLY PRODUCE THE .XML FILE
                If Not fs.FolderExists(fs.GetParentFolderName(strXMLFilePath)) Then
                    fs.CreateFolder (fs.GetParentFolderName(strXMLFilePath))
                End If
                If fs.FileExists(strXMLFilePath) Then
                    fs.DeleteFile strXMLFilePath
                End If
                .docWriteToFile (Replace(strXMLFilePath, "/", "")), False, "UNICODE", "" ''Had some cases with / in name
'  CREATING THE .PDF file
 'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file
             Set objXSL = Nothing
             Set objXSL = New MSXML2.DOMDocument60
             objXSL.async = False
             objXSL.ValidateOnParse = False
             objXSL.resolveExternals = False
             strTemplatePath = oPC.SharedFolderRoot & "\Templates\POREMINDER_DOC.xsl"
             Set fs = New FileSystemObject
             If fs.FileExists(strTemplatePath) Then
                 objXSL.Load strTemplatePath
             End If
             Set opXMLDOC = New MSXML2.DOMDocument60
             opXMLDOC.async = False
             opXMLDOC.ValidateOnParse = False
             opXMLDOC.resolveExternals = False
             
          '   GoTo ends:
             xMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC
             
             docWriteTostream strFOFilePath, opXMLDOC, "UNICODE"
                    
             'Stage 2 Convert the .FO file to .PDF and clean up
                 '    strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFilePath & " " & strPDFFilePath
             ChDir "\PBKS\Executables\FOP\"
        p 3, "GenerateReminders Pos 3"
             If fs.FileExists(strPDFFilePath) Then
                On Error Resume Next
                fs.DeleteFile strPDFFilePath, True
                If Err Then
                    MsgBox "A file cannot be cleared before being replaced. Please check their is no report document open.", vbOKOnly + vbInformation, "Can't do this"
                    Exit Function
                End If
             End If
             If fs.FileExists(strFOFilePath) Then
                     strCommand = GetFOPCommandstring(strFOFilePath, strPDFFilePath)
                     Res = F_7_AB_1_ShellAndWaitSimple(strCommand, vbHide, 10000)
        
                     p 4, "Pos 4" & CStr(Res) & strCommand
                     If oPC.TestMode = True Then
                         MsgBox "You are in TEST mode:" & vbCrLf & "Current directory = " & CurDir()
                         MsgBox "Command line = " & strCommand
                     End If
                     'IF IT FAILS HERE CHECK THAT JAVA IS INSTALLED
                     'We do the following because the batch file executes and is not waited for by the F_7_AB_1_ShellAndWaitSimple routine
                     'so the FO file is sometimes deleted before the .PDF file is produced.
                     cnt = 0
                     MsgWaitObj 5000
                     Do While Not fs.FileExists(strPDFFilePath) And cnt < 40
                         MsgWaitObj 3000
                         cnt = cnt + 1
                     Loop
                     If oPC.TestMode = True Then
                         MsgBox "You are in TEST mode:" & vbCrLf & "Wait here before deleting .FO and .XML files - Is PDF in position?"
                     End If
                     If fs.FileExists(strFOFilePath) Then
                         fs.DeleteFile strFOFilePath
                     End If
                     If fs.FileExists(strXMLFilePath) Then
                         fs.DeleteFile strXMLFilePath
                     End If
                 
                     Screen.MousePointer = vbDefault
                     If fs.FileExists(strPDFFilePath) Then
                         strExecutable = GetPDFExecutable(strPDFFilePath)
                        If strExecutable = "" Then
                            MsgBox "There is no application set on this computer to open the file: " & strPDFFilePath & ". The document cannot be displayed", vbOKOnly, "Can't do this"
                        Else
                            Shell strExecutable & " " & strPDFFilePath, vbNormalFocus
                        End If
                     End If
            Else
                MsgBox "File: " & strFOFilePath & "Does not exist."
            End If
        End With

    Loop
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.GenerateReminders(pPAID)", pPAID
End Function
Public Function PrintAReminderorCancellation(pTRID As Long, pPAID As Long, pDocType As String) As Boolean
    On Error GoTo errHandler
Dim Res As Boolean
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String
Dim objXSL As New MSXML2.DOMDocument60
Dim opXMLDOC As New MSXML2.DOMDocument60
Dim objXMLDOC  As New MSXML2.DOMDocument60
Dim xMLDoc As ujXML
Dim XMLArgs As String
Dim strXMLFilePath As String
Dim strFOFilePath As String
Dim strPDFFilePath As String
Dim strAcno As String
Dim strTmpAcno As String
Dim fs As New FileSystemObject
Dim strCommand As String
Dim cnt As Long
Dim strExecutable As String
Dim strTemplatePath As String
Dim lngTPID As Long
Dim lngTmpTPID As Long

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    PrintAReminderorCancellation = True
    
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    If pDocType = "R" Then
        cmd.CommandText = "SELECT * FROM vODPOReport_XML WHERE PAID = " & CStr(pPAID) & " AND TP_ID = " & CStr(pTRID) & " ORDER BY SupplierACNo,DocCode,TITLE"
    Else
        cmd.CommandText = "SELECT * FROM vODPOReportCancel_XML WHERE PAID = " & CStr(pPAID) & " ORDER BY SupplierACNo,DocCode,TITLE"
    End If
    cmd.commandType = adCmdText
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic
    rs.ActiveConnection = Nothing
     If rs.eof Then
        rs.Close
        Set rs = Nothing
        PrintAReminderorCancellation = False
        Exit Function
    End If
   
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Do While Not rs.eof
        p 2, "topOfLoop"
        strXMLFilePath = oPC.LocalFolder & "Reminders\" & IIf(pDocType = "R", "POReminder_", "POCancellation_") & StripToAlphanumeric(FNS(rs.Fields("SupplierName"))) & "_" & Format(Date, "MMDDHHmm") & ".xml"
        strFOFilePath = oPC.LocalFolder & "Reminders\" & IIf(pDocType = "R", "POReminder_", "POCancellation_") & StripToAlphanumeric(FNS(rs.Fields("SupplierName"))) & "_" & Format(Date, "MMDDHHmm") & ".FO"
        strPDFFilePath = oPC.LocalFolder & "Reminders\" & IIf(pDocType = "R", "POReminder_", "POCancellation_") & StripToAlphanumeric(FNS(rs.Fields("SupplierName"))) & "_" & Format(Date, "MMDDHHmm") & ".PDF"
        Set xMLDoc = New ujXML
        With xMLDoc
            .docProgID = "MSXML2.DOMDocument"
            .docInit "doc_PO_Reminder"
                .chCreate "MessageType"
                    .elText = "PO_Reminder"
                .elCreateSibling "MessageCreationDate"
                    .elText = Format(Now(), "dd-mm-yyyy HH:NN")
                .elCreateSibling "DestinationAddress"
                    .elText = FNS(rs.Fields("Email"))
                .elCreateSibling "TemplateName"
                    .elText = "POReminder_DOC"
                .elCreateSibling "SendersEmail"
                    If oPC.EmailFrom > "" Then
                        .elText = oPC.EmailFrom
                    Else
                        .elText = FNS(rs.Fields("StaffEmail"))
                    End If
                .elCreateSibling "LogoPath"
                If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
                Else
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
                End If
'                .elCreateSibling "DestinationAddress"
'                    .elText = FNS(rs.Fields("Email"))
                .elCreateSibling "Sender"
                    .elText = oPC.Configuration.DefaultCompany.CompanyName
                .elCreateSibling "SenderAddress"
                    .elText = oPC.Configuration.DefaultCompany.StreetAddress
                .elCreateSibling "SupplierName"
                    .elText = FNS(rs.Fields("SupplierName"))
                .elCreateSibling "SupplierWithAddress"
                    .elText = FNS(rs.Fields("SupplierAddress"))
                .elCreateSibling "SupplierPhone"
                    .elText = FNS(rs.Fields("Phone"))
                .elCreateSibling "SupplierFax"
                    .elText = FNS(rs.Fields("Fax"))
                .elCreateSibling "ACNo"
                    .elText = FNS(rs.Fields("SupplierACNo"))
            '    .elCreateSibling "DetailLine", True
                strAcno = FNS(rs.Fields("SupplierACNo"))
                lngTPID = FNN(rs.Fields("TP_ID"))
                lngTmpTPID = lngTPID
                Do While lngTmpTPID = lngTPID And lngTPID <> 0
                             p 3, "topOfInnerLoop"
                            ' MsgBox "TopOfInnerLoop"
                        .elCreateSibling "DetailLine", True
                            .chCreate "DocCode"
                                .elText = FNS(rs.Fields("DocCode"))
                            .elCreateSibling "DocDate", True
                                .elText = FNS(rs.Fields("DocDate"))
                            .elCreateSibling "SKU", True
                                .elText = FNS(rs.Fields("CODEF"))
                            .elCreateSibling "Title", True
                                .elText = FNS(rs.Fields("Title"))
                            .elCreateSibling "Ordered", True
                                .elText = FNS(rs.Fields("Qty"))
                            .elCreateSibling "Outstanding", True
                                .elText = FNS(rs.Fields("QtyOS"))
                            .elCreateSibling "POLReference", True
                                .elText = FNS(rs.Fields("POLRef"))
                            .elCreateSibling "StaffMember", True
                                .elText = FNS(rs.Fields("StaffMember"))
                            .elCreateSibling "Cancelled", True
                                .elText = FNS(rs.Fields("POLA_IsCancelled"))
                        .navUP
                        rs.MoveNext
                        If rs.eof Then
                            lngTPID = 0
                        Else
                            lngTPID = FNN(rs.Fields("TP_ID"))
                        End If
                Loop
    
        'FINALLY PRODUCE THE .XML FILE
                If Not fs.FolderExists(fs.GetParentFolderName(strXMLFilePath)) Then
                    fs.CreateFolder (fs.GetParentFolderName(strXMLFilePath))
                End If
                If fs.FileExists(strXMLFilePath) Then
                    fs.DeleteFile strXMLFilePath
                End If
                .docWriteToFile (Replace(strXMLFilePath, "/", "")), False, "UNICODE", "" ''Had some cases with / in name
'  CREATING THE .PDF file
 'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file
             Set objXSL = Nothing
             Set objXSL = New MSXML2.DOMDocument60
             objXSL.async = False
             objXSL.ValidateOnParse = False
             objXSL.resolveExternals = False
            If pDocType = "R" Then
                strTemplatePath = oPC.SharedFolderRoot & "\Templates\POREMINDER_DOC.xsl"
            Else
                strTemplatePath = oPC.SharedFolderRoot & "\Templates\POCancellation_DOC_FO.xsl"
            End If
             Set fs = New FileSystemObject
             If fs.FileExists(strTemplatePath) Then
                 objXSL.Load strTemplatePath
             End If
             Set opXMLDOC = New MSXML2.DOMDocument60
             opXMLDOC.async = False
             opXMLDOC.ValidateOnParse = False
             opXMLDOC.resolveExternals = False
             
          '   GoTo ends:
             xMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC
             
             docWriteTostream strFOFilePath, opXMLDOC, "UNICODE"
                    
             'Stage 2 Convert the .FO file to .PDF and clean up
                 '    strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFilePath & " " & strPDFFilePath
             ChDir "\PBKS\Executables\FOP\"
        p 3, "GenerateRemindersorCancellations Pos 3"
             If fs.FileExists(strPDFFilePath) Then
                  On Error Resume Next
                fs.DeleteFile strPDFFilePath, True
                If Err Then
                    MsgBox "A file cannot be cleared before being replaced. Please check there is no reminder document open.", vbOKOnly + vbInformation, "Can't do this"
                       PrintAReminderorCancellation = False
                    Exit Function
                End If
             End If
             If fs.FileExists(strFOFilePath) Then
                     strCommand = GetFOPCommandstring(strFOFilePath, strPDFFilePath)
                     Res = F_7_AB_1_ShellAndWaitSimple(strCommand, vbHide, 10000)
        
                     p 4, "Pos 4" & CStr(Res) & strCommand
                     If oPC.TestMode = True Then
                         MsgBox "You are in TEST mode:" & vbCrLf & "Current directory = " & CurDir()
                         MsgBox "Command line = " & strCommand
                     End If
                     'IF IT FAILS HERE CHECK THAT JAVA IS INSTALLED
                     'We do the following because the batch file executes and is not waited for by the F_7_AB_1_ShellAndWaitSimple routine
                     'so the FO file is sometimes deleted before the .PDF file is produced.
                     cnt = 0
                     MsgWaitObj 5000
                     Do While Not fs.FileExists(strPDFFilePath) And cnt < 40
                         MsgWaitObj 3000
                         cnt = cnt + 1
                     Loop
                     If oPC.TestMode = True Then
                         MsgBox "You are in TEST mode:" & vbCrLf & "Wait here before deleting .FO and .XML files - Is PDF in position?"
                     End If
                     If fs.FileExists(strFOFilePath) Then
                         fs.DeleteFile strFOFilePath
                     End If
                     If fs.FileExists(strXMLFilePath) Then
                         fs.DeleteFile strXMLFilePath
                     End If
                 
                     Screen.MousePointer = vbDefault
                     If fs.FileExists(strPDFFilePath) Then
                         strExecutable = GetPDFExecutable(strPDFFilePath)
                          If strExecutable = "" Then
                              MsgBox "There is no application set on this computer to open the file: " & strPDFFilePath & ". The document cannot be displayed", vbOKOnly, "Can't do this"
                          Else
                              Shell strExecutable & " " & strPDFFilePath, vbNormalFocus
                          End If
                     End If
            Else
                MsgBox "File: " & strFOFilePath & "Does not exist."
            End If
        End With

    Loop
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_StockManager.PrintAReminder(pTRID)", pTRID
    Exit Function
errHandler:
    ErrorIn "z_StockManager.PrintAReminderorCancellation(pTRID,pPAID,pDocType)", Array(pTRID, pPAID, _
         pDocType), , , "Error line", Array(Erl())
End Function


Public Function GenerateCancellationReports(pPAID As Long) As Boolean
    On Error GoTo errHandler
    Dim Res As Boolean
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String
Dim objXSL As New MSXML2.DOMDocument60
Dim opXMLDOC As New MSXML2.DOMDocument60
Dim objXMLDOC  As New MSXML2.DOMDocument60
Dim xMLDoc As ujXML
Dim XMLArgs As String
Dim strXMLFilePath As String
Dim strFOFilePath As String
Dim strPDFFilePath As String
Dim strAcno As String
Dim strTmpAcno As String
Dim fs As New FileSystemObject
Dim strCommand As String
Dim cnt As Long
Dim strExecutable As String
Dim strTemplatePath As String
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    GenerateCancellationReports = True
    
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "SELECT * FROM vODPOReportCancel_XML WHERE PAID = " & CStr(pPAID) & " ORDER BY SupplierACNo,DocCode,TITLE"
    cmd.commandType = adCmdText
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic
    rs.ActiveConnection = Nothing
     If rs.eof Then
        rs.Close
        Set rs = Nothing
        GenerateCancellationReports = False
        Exit Function
    End If
   
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Do While Not rs.eof
        strXMLFilePath = oPC.LocalFolder & "Reminders\" & "POCancellationReport_" & Left(Replace(FNS(rs.Fields("SupplierName")), " ", "_"), 20) & "_" & Format(Date, "MMDDHHmm") & ".xml"
        strFOFilePath = oPC.LocalFolder & "Reminders\" & "POCancellationReport_" & Left(Replace(FNS(rs.Fields("SupplierName")), " ", "_"), 20) & "_" & Format(Date, "MMDDHHmm") & ".FO"
        strPDFFilePath = oPC.LocalFolder & "Reminders\" & "POCancellationReport_" & Left(Replace(FNS(rs.Fields("SupplierName")), " ", "_"), 20) & "_" & Format(Date, "MMDDHHmm") & ".PDF"
        Set xMLDoc = New ujXML
        With xMLDoc
            .docProgID = "MSXML2.DOMDocument"
            .docInit "doc_PO_Reminder"
                .chCreate "MessageType"
                    .elText = "PO_Reminder"
                .elCreateSibling "MessageCreationDate"
                    .elText = Format(Now(), "dd-mm-yyyy HH:NN")
                .elCreateSibling "DestinationAddress"
                    .elText = FNS(rs.Fields("Email"))
                .elCreateSibling "TemplateName"
                    .elText = "POReminder_DOC"
                .elCreateSibling "SendersEmail"
                    If oPC.EmailFrom > "" Then
                        .elText = oPC.EmailFrom
                    Else
                        .elText = FNS(rs.Fields("StaffEmail"))
                    End If
                .elCreateSibling "LogoPath"
                If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
                Else
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
                End If
'                .elCreateSibling "DestinationAddress"
'                    .elText = FNS(rs.Fields("Email"))
                .elCreateSibling "Sender"
                    .elText = oPC.Configuration.DefaultCompany.CompanyName
                .elCreateSibling "SenderAddress"
                    .elText = oPC.Configuration.DefaultCompany.StreetAddress
                .elCreateSibling "SupplierName"
                    .elText = FNS(rs.Fields("SupplierName"))
                .elCreateSibling "SupplierWithAddress"
                    .elText = FNS(rs.Fields("SupplierAddress"))
                .elCreateSibling "SupplierPhone"
                    .elText = FNS(rs.Fields("Phone"))
                .elCreateSibling "SupplierFax"
                    .elText = FNS(rs.Fields("Fax"))
                .elCreateSibling "ACNo"
                    .elText = FNS(rs.Fields("SupplierACNo"))
            '    .elCreateSibling "DetailLine", True
                strAcno = FNS(rs.Fields("SupplierACNo"))
                strTmpAcno = strAcno
                Do While strTmpAcno = strAcno
                        .elCreateSibling "DetailLine", True
                            .chCreate "DocCode"
                                .elText = FNS(rs.Fields("DocCode"))
                            .elCreateSibling "DocDate", True
                                .elText = FNS(rs.Fields("DocDate"))
                            .elCreateSibling "SKU", True
                                .elText = FNS(rs.Fields("CODEF"))
                            .elCreateSibling "Title", True
                                .elText = FNS(rs.Fields("Title"))
                            .elCreateSibling "Ordered", True
                                .elText = FNS(rs.Fields("Qty"))
                            .elCreateSibling "Outstanding", True
                                .elText = FNS(rs.Fields("QtyOS"))
                            .elCreateSibling "POLReference", True
                                .elText = FNS(rs.Fields("POLRef"))
                            .elCreateSibling "StaffMember", True
                                .elText = FNS(rs.Fields("StaffMember"))
                            .elCreateSibling "Cancelled", True
                                .elText = FNS(rs.Fields("POLA_IsCancelled"))
                        .navUP
                        rs.MoveNext
                        If rs.eof Then
                            strTmpAcno = ""
                        Else
                            strTmpAcno = FNS(rs.Fields("SupplierACNo"))
                        End If
                Loop
    
        'FINALLY PRODUCE THE .XML FILE
                If Not fs.FolderExists(fs.GetParentFolderName(strXMLFilePath)) Then
                    fs.CreateFolder (fs.GetParentFolderName(strXMLFilePath))
                End If
                If fs.FileExists(strXMLFilePath) Then
                    fs.DeleteFile strXMLFilePath
                End If
                .docWriteToFile (strXMLFilePath), False, "UNICODE", ""
'  CREATING THE .PDF file
 'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file
        Set objXSL = Nothing
        Set objXSL = New MSXML2.DOMDocument60
        objXSL.async = False
        objXSL.ValidateOnParse = False
        objXSL.resolveExternals = False
        strTemplatePath = oPC.SharedFolderRoot & "\Templates\POCancellation_DOC_FO.xsl"
        Set fs = New FileSystemObject
        If fs.FileExists(strTemplatePath) Then
            objXSL.Load strTemplatePath
        End If
        Set opXMLDOC = New MSXML2.DOMDocument60
        opXMLDOC.async = False
        opXMLDOC.ValidateOnParse = False
        opXMLDOC.resolveExternals = False
        xMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC
        
        docWriteTostream strFOFilePath, opXMLDOC, "UNICODE"
               
        'Stage 2 Convert the .FO file to .PDF and clean up
            '    strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFilePath & " " & strPDFFilePath
        ChDir "\PBKS\Executables\FOP\"
        p 3, "GenerateCancellations Pos 3"
                 If fs.FileExists(strPDFFilePath) Then
                     fs.DeleteFile strPDFFilePath, True
                 End If
                 If fs.FileExists(strFOFilePath) Then
                     strCommand = GetFOPCommandstring(strFOFilePath, strPDFFilePath)
                     Res = F_7_AB_1_ShellAndWaitSimple(strCommand, vbHide, 10000)
'                   ' strCommand = "FOP.BAT" & " " & strFOFilePath & " " & strPDFFilePath
'                    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 10000
'                         res = F_7_AB_1_ShellAndWaitSimple(strCommand, vbHide, 50000)
            
                         p 4, "GenerateCancellations:Pos 4" & CStr(Res) & strCommand
                    If oPC.TestMode = True Then
                        MsgBox "You are in TEST mode:" & vbCrLf & "Current directory = " & CurDir()
                        MsgBox "Command line = " & strCommand
                    End If
                    'IF IT FAILS HERE CHECK THAT JAVA IS INSTALLED
                    'We do the following because the batch file executes and is not waited for by the F_7_AB_1_ShellAndWaitSimple routine
                    'so the FO file is sometimes deleted before the .PDF file is produced.
                    cnt = 0
                    Do While Not fs.FileExists(strPDFFilePath) And cnt < 40
                        MsgWaitObj 1000
                        cnt = cnt + 1
                    Loop
                    If oPC.TestMode = True Then
                        MsgBox "You are in TEST mode:" & vbCrLf & "Wait here before deleting .FO and .XML files - Is PDF in position?"
                    End If
                    If fs.FileExists(strFOFilePath) Then
                        fs.DeleteFile strFOFilePath
                    End If
                    If fs.FileExists(strXMLFilePath) Then
                        fs.DeleteFile strXMLFilePath
                    End If
                
                Screen.MousePointer = vbDefault
                If fs.FileExists(strPDFFilePath) Then
                    strExecutable = GetPDFExecutable(strPDFFilePath)
                    If strExecutable = "" Then
                        MsgBox "There is no application set on this computer to open the file: " & strPDFFilePath & ". The document cannot be displayed", vbOKOnly, "Can't do this"
                    Else
                        Shell strExecutable & " " & strPDFFilePath, vbNormalFocus
                    End If
                End If
            Else
                MsgBox "File: " & strFOFilePath & "Does not exist."
            End If
                
        End With

                    

    Loop
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.GenerateCancellationReports(pPAID)", pPAID
End Function

Public Sub docWriteTostream(ByVal FilePath As String, obj As MSXML2.DOMDocument60, _
                Optional ByVal CharSet As String = "UNICODE")
    On Error GoTo errHandler
    Dim s As Object
    Set s = CreateObject("ADODB.Stream")
    With s
        If CharSet <> "" Then .CharSet = CharSet
        .Open
        .WriteText obj.xml
        .SaveToFile FilePath, 2 'adSaveCreateOverWrite
        .Close
    End With
    Exit Sub
errHandler:
    ErrorIn "a_PO.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
End Sub


Public Function GenerateBOStatusReports(PAID As Long, Optional pReturnEmailAddress As String) As Boolean
    On Error GoTo errHandler
'    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim strPath As String
Dim strBillto As String
Dim strDelto As String
Dim strFOFile As String
Dim strPDFFile As String
Dim strXML As String
Dim strCommand As String
Dim i As Integer
Dim strHTML As String
Dim fs As New FileSystemObject
Dim objXSL As New MSXML2.DOMDocument60
Dim opXMLDOC As New MSXML2.DOMDocument60
Dim objXMLDOC  As New MSXML2.DOMDocument60
Dim strWorkingFolder As String
Dim strDispatchMethod As String
Dim strExecutable As String
Dim cnt As Integer
Dim rs As New ADODB.Recordset
Dim oSM As New z_StockManager
Dim xMLDoc As ujXML
Dim lngTPID As Long
Dim bEODoc As Boolean
Dim oSQL As New z_SQL
Dim strSM As String
Dim strCM As String
Dim strSignature As String

    GenerateBOStatusReports = True

    Set rs = GetBOReportData(PAID, strSM, strCM, strSignature)
'
    strWorkingFolder = oPC.SharedFolderRoot & "\PDF\"
    If rs.eof Then
        rs.Close
        Set rs = Nothing
        GenerateBOStatusReports = False
        Exit Function
    End If
'    If FNS(rs.Fields("DISPATCHMETHOD")) > "" Then
'        strDispatchMethod = FNS(rs.Fields("DISPATCHMETHOD"))
'    End If
    Do While Not rs.eof
      '  If FNS(rs.Fields("ActionType")) = "COL" Then
        Set xMLDoc = Nothing
        Set xMLDoc = New ujXML
        With xMLDoc
            'code removes spaces and commas from name
            strXML = strWorkingFolder & "CustomerReport" & Replace(oPC.WorkstationName, "-", "") & StripToAlphanumeric(FNS(rs.Fields("TP_NAME"))) & ".xml"
            .docProgID = "MSXML2.DOMDocument"
            .docInit "CustomerReport_DOC"
                .chCreate "MessageType"
                    .elText = "CustomerReport"
                .elCreateSibling "MessageCreationDate"
                    .elText = Format(Now(), "yyyymmddHHNN")
                .elCreateSibling "TemplateName"
                    .elText = "CustomerReport_DOC"
                .elCreateSibling "SendersEmail"
                    If pReturnEmailAddress > "" Then
                        .elText = pReturnEmailAddress
                    Else
                        .elText = oPC.EmailFrom
                    End If
                .elCreateSibling "LogoPath"
                If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
                Else
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
                End If
                .elCreateSibling "DocDate", True
                    .elText = Format(Date, "dd/mm/yyyy")
                .elCreateSibling "OurCompany", True
                    .elText = oPC.Configuration.DefaultCompany.CompanyName
                .elCreateSibling "OurAddress", True
                    .elText = Replace(oPC.Configuration.DefaultCompany.StreetAddress, Chr(13) & Chr(10), Chr(10))
                .elCreateSibling "CustomerName", True
                    .elText = FNS(rs.Fields("FULLNAME"))
                .elCreateSibling "CustomerPhone", True
                    .elText = FNS(rs.Fields("ADD_Phone"))
                .elCreateSibling "CustomerFax", True
                    .elText = FNS(rs.Fields("ADD_Fax"))
                .elCreateSibling "CustomerAddress", True
                    .elText = FNS(rs.Fields("CustomerAddress"))
                .elCreateSibling "ACNo", True
                    .elText = FNS(rs.Fields("TP_ACNO"))
                .elCreateSibling "Message", True
                    .elText = strCM
                .elCreateSibling "ReturnEmail", True
                    .elText = ""
                .elCreateSibling "CustomersEmail", True
                    .elText = FNS(rs.Fields("ADD_EMAIL"))
                .elCreateSibling "Signature", True
                    .elText = strSignature
                lngTPID = FNN(rs.Fields("TP_ID"))
                bEODoc = False
                Do While bEODoc = False
                        .elCreateSibling "DetailLine", True
                        .chCreate "DocumentRef"
                            .elText = FNS(rs.Fields("TR_CODE"))
                        .elCreateSibling "DocDate", True
                            .elText = Format(FND(rs.Fields("TR_DATE")), "dd/mm/yyyy")
                        .elCreateSibling "Title", True
                            .elText = FNS(rs.Fields("P_Title"))
                        .elCreateSibling "ISBN", True
                            .elText = FNS(rs.Fields("CODEF"))
                        .elCreateSibling "Qty", True
                            .elText = FNN(rs.Fields("COL_QTY"))
                        .elCreateSibling "QtyDispatched", True
                            .elText = FNN(rs.Fields("COL_QTYDISPATCHED"))
                        .elCreateSibling "Reference", True
                            .elText = FNS(rs.Fields("COL_Ref"))
                        .elCreateSibling "BookStatus", True
                            .elText = FNS(rs.Fields("P_STATUS"))
                        .elCreateSibling "DetailedMessage", True
                            .elText = FNS(rs.Fields("COLA_DetailedMessage"))
                        .elCreateSibling "COL_ETA", True
                            .elText = Format(FND(rs.Fields("COL_ETA")), "dd/mm/yyyy")
                        .navUP
                    rs.MoveNext
                    If rs.eof Then
                        bEODoc = True
                    Else
                        If lngTPID <> FNN(rs.Fields("TP_ID")) Then bEODoc = True
                    End If
                Loop
                ''FINALLY PRODUCE THE .XML FILE
                With xMLDoc
                    If fs.FileExists(strXML) Then
                        fs.DeleteFile strXML
                    End If
                    .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
                End With
                'Produce the PDF
                Set objXSL = Nothing
                Set objXSL = New MSXML2.DOMDocument60
                objXSL.async = False
                objXSL.ValidateOnParse = False
                objXSL.resolveExternals = False
                strPath = oPC.SharedFolderRoot & "\Templates\CustomerReport_FO.xsl"
                Set fs = New FileSystemObject
                If fs.FileExists(strPath) Then
                    objXSL.Load strPath
                End If
                Set opXMLDOC = New MSXML2.DOMDocument60
                opXMLDOC.async = False
                opXMLDOC.ValidateOnParse = False
                opXMLDOC.resolveExternals = False
                xMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC
                strFOFile = strXML & ".FO"
                strPDFFile = strXML & ".PDF"
                docWriteTostream strFOFile, opXMLDOC, "UNICODE"
                
                ''Stage 2 Convert the .FO file to .PDF and clean up
                '       ''' strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFile & " " & strPDFFile
            '    strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFile & " " & strPDFFile
            
                ChDir "\PBKS\Executables\FOP\"
               ' strCommand = "FOP.BAT" & " " & strFOFile & " " & strPDFFile
                strCommand = GetFOPCommandstring(strFOFile, strPDFFile)
                F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 600000
                'IF IT FAILS HERE CHECK THAT JAVA IS INSTALLED
                'We do the following because the batch file executes and is not waited for by the F_7_AB_1_ShellAndWaitSimple routine
                'so the FO file is sometimes deleted before the .PDF file is produced.
                cnt = 0
                Do While Not fs.FileExists(strPDFFile) And cnt < 10
                    MsgWaitObj 1000
                    cnt = cnt + 1
                Loop
                If fs.FileExists(strFOFile) Then
                    fs.DeleteFile strFOFile
                End If
                strExecutable = GetPDFExecutable(strPDFFile)
                If strExecutable = "" Then
                    MsgBox "There is no application set on this computer to open the file: " & strPDFFile & ". The document cannot be displayed", vbOKOnly, "Can't do this"
                Else
                    Shell strExecutable & " " & strPDFFile, vbNormalFocus
                End If
        End With
'        Else
'            rs.MoveNext
'        End If
    Loop

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.GenerateBOStatusReports(PAID,pReturnEmailAddress)", Array(PAID, _
         pReturnEmailAddress)
End Function

Public Function GeneratOrderStatusTransmission(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim sEDIMsg As String
Dim i As Integer
Dim sDocCOde As String

        rs.MoveFirst
        sDocCOde = FNS(rs.Fields("TRCODE"))
        sEDIMsg = "UNA:+.? '"
        sEDIMsg = sEDIMsg & "UNB+UNOC:3+1556150:31B+123HORI:31B+060302:1011+643'"
        sEDIMsg = sEDIMsg & "UNH++ORDRSP'"
        sEDIMsg = sEDIMsg & "BGM+231+" & sDocCOde & "-" & CStr(Format(Now(), "YYYYMMDDHHNN")) & "'"
        sEDIMsg = sEDIMsg & "DTM+137:" & Format(Date, "YYYYMMDD") & ":102'"
        i = 0
        Do While Not rs.eof
            i = i + 1
            sEDIMsg = sEDIMsg & "LIN+" & CStr(i) & "+24'"
            sEDIMsg = sEDIMsg & "QTY+21:" & CStr(FNN(rs.Fields("QtyOrdered"))) & "'"
            sEDIMsg = sEDIMsg & "QTY+12:" & CStr(FNN(rs.Fields("QtyDispatched"))) & "'"
           ' sEDIMsg = sEDIMsg & "DTM+11:" & Format(FNS(COSRL_ETA), "YYYYMMDD") & ":610'"
            If IsDate(rs.Fields("COL_ETA")) Then
                If CDate(rs.Fields("COL_ETA")) > CDate("2008-01-01") Then
                    sEDIMsg = sEDIMsg & "DTM+44:" & Format(FNS(rs.Fields("COL_ETA")), "YYYYMMDD") & ":610'"
                End If
            End If
            sEDIMsg = sEDIMsg & "FTX+LIN++" & FNS(rs.Fields("COSRL_AvailabilityCode")) & ":8B:28'"
            sEDIMsg = sEDIMsg & "FTX+LIN++" & FNS(rs.Fields("COSRL_ActionCode")) & ":12B:28'"
            sEDIMsg = sEDIMsg & "RFF+LI:" & FNS(rs.Fields("COSRL_CustomerOrderRef")) & "'"
            rs.MoveNext
        Loop
        sEDIMsg = sEDIMsg & "UNS+S'"
        sEDIMsg = sEDIMsg & "UNT+51+61'"
        sEDIMsg = sEDIMsg & "UNZ+1+643'"
        oTF.OpenTextFile oPC.SharedFolderRoot & "\EDI_OUT\COSR\" & sDocCOde & ".TXT"
        oTF.WriteToTextFile_NoLineTerminator sEDIMsg
        oTF.CloseTextFile

MsgBox sEDIMsg
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_StockManager.GeneratOrderStatusTransmission(PAID)", PAID
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.GeneratOrderStatusTransmission(rs)", rs
End Function


Public Function LoadForPreDelAdvice(x As String, POLS As ADODB.Recordset, COLS As ADODB.Recordset, PROS As ADODB.Recordset, PID As String, XMLType As String)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String
Dim rs As ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.commandType = adCmdText
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "LoadForPreDelAdvice_POLS"
    cmd.commandType = adCmdStoredProc
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@POLS", adVarChar, adParamInput, 10000, x)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 40, PID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@XMLType", adChar, adParamInput, 1, XMLType)
    cmd.Parameters.Append par

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockOptimistic
    Set rs.ActiveConnection = Nothing
    Set POLS = rs
    Set rs = Nothing
    Set cmd = Nothing
    
'================
    Set cmd = New ADODB.Command
    cmd.commandType = adCmdText
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "LoadForPreDelAdvice_COLS"
    cmd.commandType = adCmdStoredProc
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@POLS", adVarChar, adParamInput, 10000, x)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 40, PID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@XMLType", adChar, adParamInput, 1, XMLType)
    cmd.Parameters.Append par
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockOptimistic
    Set rs.ActiveConnection = Nothing
    Set COLS = rs

    Set rs = Nothing
    Set cmd = Nothing
'=================================
    Set cmd = New ADODB.Command
    cmd.commandType = adCmdText
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "LoadForPreDelAdvice_Products"
    cmd.commandType = adCmdStoredProc
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@XMLArgs", adVarChar, adParamInput, 10000, IIf(XMLType = "P", "", x))
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@POLS", adVarChar, adParamInput, 10000, IIf(XMLType = "P", x, ""))
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 40, PID)
    cmd.Parameters.Append par
  
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockOptimistic
    Set rs.ActiveConnection = Nothing
    Set PROS = rs

    Set rs = Nothing
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.LoadForPreDelAdvice(X,POLS,COLS)", Array(x, POLS, COLS)
End Function
Public Sub Action_POLSet(POLS As String, COLS As String, WSName As String, NewDiaryDate As Date, SupplierMessage As String, _
        CustomerMessage As String, NEWPRODUCTSTATUS As String, lngPaid As Long, _
        strSignature As String)
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
Dim i As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "Action_POLSet"
    cmd.commandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.COShort
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@POLS", adVarChar, adParamInput, 10000, POLS)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@COLS", adVarChar, adParamInput, 10000, COLS)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@WSName", adVarChar, adParamInput, 50, WSName)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@NewDiaryDate", adDate, adParamInput, , NewDiaryDate)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@SupplierMessage", adVarChar, adParamInput, 150, SupplierMessage)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@CustomerMessage", adVarChar, adParamInput, 300, CustomerMessage)
    cmd.Parameters.Append par
       
    Set par = Nothing
    Set par = cmd.CreateParameter("@NEWPRODUCTSTATUS", adVarChar, adParamInput, 1, NEWPRODUCTSTATUS)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@Signature", adVarChar, adParamInput, 100, strSignature)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@PAID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.execute

    lngPaid = FNN(cmd.Parameters("@PAID"))
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.Action_POLSet(WSName,NewDiaryDate,SupplierMessage,CustomerMessage," & _
        "NEWPRODUCTSTATUS,lngPAID,strSignature)", Array(WSName, NewDiaryDate, SupplierMessage, CustomerMessage, _
         NEWPRODUCTSTATUS, lngPaid, strSignature)
End Sub

Public Sub ProductStatusChange(PROS As String, POLS As String, COLS As String, WSName As String, SupplierMessage As String, _
        NewStatusID As Long, NewETA As Date, strSignature As String, lngPSCID As Long, COLActionID As Long)
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
Dim i As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "Action_ProductStatusChange"
    cmd.commandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.COShort
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@PROS", adVarChar, adParamInput, 10000, PROS)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@POLS", adVarChar, adParamInput, 10000, POLS)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@COLS", adVarChar, adParamInput, 10000, COLS)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@WSName", adVarChar, adParamInput, 50, WSName)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@SupplierMessage", adVarChar, adParamInput, 150, SupplierMessage)
    cmd.Parameters.Append par
       
    Set par = Nothing
    Set par = cmd.CreateParameter("@NEWPRODUCTSTATUS", adInteger, adParamInput, , NewStatusID)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@COLActionID", adInteger, adParamInput, , COLActionID)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@NewETA", adDate, adParamInput, , NewETA)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@Signature", adVarChar, adParamInput, 100, strSignature)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@PSCID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@pPAID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.execute

    lngPSCID = FNN(cmd.Parameters("@PSCID"))
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.ProductStatusChange(PROS,WSName,SupplierMessage,NewStatusID,strSignature," & _
        "lngPSCID)", Array(PROS, WSName, SupplierMessage, NewStatusID, strSignature, lngPSCID)
End Sub

Public Sub CreateTransferFromXML(pStoreID As Long, pFilename As String, pNewID As Long)
Dim oTF As New z_TextFile
Dim s As String
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim lngTFID As Long


Dim fs As New FileSystemObject
Dim TsStream As TextStream

Dim strAlltxt As String, strLine As String
Dim varAsplit As Variant, varBsplit As Variant

    Set TsStream = fs.OpenTextFile(pFilename, ForReading, , TristateTrue) 'read the file content in memory with the unicode flag "TristateTrue"
    strAlltxt = TsStream.ReadAll ' atach the stream to a string var
    TsStream.Close ' cleanup and dispose unnecesary objects
    Set TsStream = Nothing
    Set fs = Nothing
    

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "CreateTransfer"
    cmd.commandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.COShort
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@StoreID", adInteger, adParamInput, , pStoreID)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@XML", adVarChar, adParamInput, 10000, strAlltxt)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@TFID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.execute

    lngTFID = FNN(cmd.Parameters("@TFID"))
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    pNewID = lngTFID
    
End Sub

Public Sub CreateGRNFromFile(pSupplierID As Long, pFilename As String, pNewID As Long)
Dim pErrorFilePath As String
Dim pMsg As String
Dim oSQL As New z_SQL

    pErrorFilePath = oPC.SharedFolderRoot & "\Logs\GRNImport" & Format(Now(), "YYYYMMDDHHNN") & ".txt"
    oSQL.ImportFromFile "GRN2", oPC.SharedFolderRoot & "\Templates\ImportGRN2Format.xml", pFilename, pMsg, pErrorFilePath

''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    Set cmd = New ADODB.Command
'    cmd.CommandText = "CreateGRN"
'    cmd.CommandType = adCmdStoredProc
'    cmd.ActiveConnection = oPC.COShort
'
'    Set par = Nothing
'    Set par = cmd.CreateParameter("@SupplierID", adInteger, adParamInput, , pSupplierID)
'    cmd.Parameters.Append par
'
'    Set par = Nothing
'    Set par = cmd.CreateParameter("@XML", adVarChar, adParamInput, 10000, strAlltxt)
'    cmd.Parameters.Append par
'
'    Set par = Nothing
'    Set par = cmd.CreateParameter("@TRID", adInteger, adParamOutput)
'    cmd.Parameters.Append par
'
'    cmd.execute
'
'    pNewID = FNN(cmd.Parameters("@TRID"))
'    Set cmd = Nothing
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'    pNewID = lngTFID
    
End Sub


Public Function GetOverInvoicedIems(pTRID As Long) As String
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim sOverinvoiced As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "GetOverInvoicedIems"
    cmd.commandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.COShort
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@TRID", adInteger, adParamInput, , pTRID)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@Res", adVarChar, adParamOutput, 1100)
    cmd.Parameters.Append par
    
    cmd.execute

    GetOverInvoicedIems = FNS(cmd.Parameters("@Res"))
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Function

'Public Function GenerateReturnsPerPub(pTPID As Long, pMinDate As Date, pMaxDate As Date, pStaffID As Long, Optional bSupplementary As Boolean) As Integer
Public Function GenerateReturnsPerPub(pTPID As Long, pSpecifiedMonth As Date, pSTAFFID As Long, Optional bSupplementary As Boolean) As Integer
    On Error GoTo errHandler
'NOTE Min date is the most recent date and maxdate is the least recent date
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset

Dim rsR As ADODB.Recordset
Dim rsRL As ADODB.Recordset
Dim rsTR As ADODB.Recordset
Dim RRP As ADODB.Recordset
Dim oCode As New z_Code
Dim lngTRID  As Long
Dim i As Long
Dim lngCurrentRLID As Long
Dim strRef As String
Dim lngPrice As Long
Dim dblDiscount As Double
Dim iCount As Integer
Dim strDefaultCo As String
Dim OpenResult As Integer
Dim oSQL As z_SQL
Dim lngTPID As Long


    'Get titles that have been received SeeSafe on supplier invoices prior to or on the pDATE parameter.
    'The Titles should have QtyOnHand > 0
    If bSupplementary Then
    strSQL = "SELECT SUM(dbo.tRL.RL_QtyApproved - dbo.tRL.RL_QtyReturned) AS QTYSS, dbo.tProduct.P_ID, dbo.tR.R_ApprovalRef, dbo.tProduct.P_QtyOnHand, " _
            & " dbo.tProduct.P_RRP , dbo.tProduct.P_Title " _
            & " FROM dbo.tTR INNER JOIN " _
            & " dbo.tR INNER JOIN " _
            & " dbo.tRL ON dbo.tR.R_ID = dbo.tRL.RL_TR_ID INNER JOIN " _
            & " dbo.tProduct ON dbo.tRL.RL_P_ID = dbo.tProduct.P_ID ON dbo.tTR.TR_ID = dbo.tR.R_ID " _
            & " Where (dbo.tRL.RL_QtyApproved > dbo.tRL.RL_QtyReturned) AND (isnull(tR.R_Type,'S') <> 'E') AND tTR.TR_TP_ID = " & pTPID _
            & " GROUP BY  dbo.tProduct.P_ID, " _
            & " dbo.tR.R_ApprovalRef, dbo.tProduct.P_QtyOnHand, dbo.tProduct.P_RRP, dbo.tProduct.P_Title" _
            & " Having (dbo.tProduct.P_QtyOnHand > 0)"
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
        
        '-------------------------------
            OpenResult = oPC.OpenDBSHort
        '-------------------------------
            rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockOptimistic
    Else
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            Set oSQL = New z_SQL
            oSQL.FetchPotentialReturns pTPID, pSpecifiedMonth, "", rs
    End If
    
    If rs.eof Then
        GenerateReturnsPerPub = 0
        Exit Function
    Else
        GenerateReturnsPerPub = 1
    End If
    
    RaiseEvent eMax(rs.RecordCount)
    i = 0
    oPC.COShort.BeginTrans
        lngTPID = CLng(rs!TPID)
        Do While Not rs.eof
                        Set rsTR = New ADODB.Recordset
                        rsTR.Open "Select * FROM tTR WHERE TR_ID = 0", oPC.COShort, adOpenKeyset, adLockOptimistic
                        rsTR.AddNew
                            rsTR!TR_DATE = ReverseDate(Date)
                            rsTR!TR_TP_ID = rs!TPID
                            rsTR!TR_CaptureDate = ReverseDateTime(Now)
                            strDefaultCo = oPC.Configuration.DefaultCompany.CompanyCode
                            rsTR!TR_CODE = strDefaultCo & "R" & oCode.GetNextCode(enReturn, "GETANDWRITE")
                            rsTR!TR_Type = 11
                            If bSupplementary Then
                                rsTR!TR_Status = 3
                            Else
                                rsTR!TR_Status = 2
                            End If
                            rsTR!TR_StaffID = pSTAFFID
                            rsTR!TR_COMP_ID = oPC.Configuration.DefaultCOMPID
                        rsTR.Update
                        rsTR.Bookmark = rsTR.Bookmark
                        lngTRID = rsTR("TR_ID")
                        
                        Set rsR = New ADODB.Recordset
                        rsR.Open "Select * FROM tR WHERE R_ID = 0", oPC.COShort, adOpenKeyset, adLockOptimistic
                        rsR.AddNew
                            rsR!R_ID = lngTRID
                            rsR!R_Status = 0
                            If bSupplementary Then
                                rsR!R_Type = "E"
                            Else
                                rsR!R_Type = "P"
                            End If
                        rsR.Update
                        
                        Set rsRL = New ADODB.Recordset
                        rsRL.Open "Select * FROM tRL WHERE RL_ID = 0", oPC.COShort, adOpenKeyset, adLockOptimistic
                        i = 0
                        Do While Not rs.eof
                                If rs!TPID <> lngTPID Then Exit Do
                            rsRL.AddNew
                                rsRL!RL_P_ID = rs!P_ID
                                rsRL!RL_TR_ID = lngTRID
                                rsRL!RL_Price = rs!P_RRP
                            '    rsRL!RL_DELL_ID = rs!DELL_ID
                                rsRL!RL_QtySystem = rs!P_QtyOnHand
                                rsRL!RL_QtyRequested = GetMin(CLng(rs!P_QtyOnHand), CLng(rs!QtySS))
                                If bSupplementary Then
                                    rsRL!RL_QtyApproved = GetMin(CLng(rs!P_QtyOnHand), CLng(rs!QtySS))
                                    rsRL!RL_SiREF = FNS(rs!R_ApprovalRef)
                               End If
                            rsRL.Update
                            rs.MoveNext
                            If rs.eof Then Exit Do
                            i = i + 1
                            RaiseEvent eProgress(i)
                        Loop
                        
                        If Not bSupplementary Then
                             Set RRP = Nothing
                             Set RRP = New ADODB.Recordset
                             RRP.CursorLocation = adUseClient
                            oSQL.ReturnRecordset RRP, lngTRID, pSpecifiedMonth
                            Do While Not RRP.eof
                             '   If rRP!TPID <> lngTPID Then Exit Do
                                lngCurrentRLID = RRP!RL_ID
                                strRef = ""
                                iCount = 0
                                Do While lngCurrentRLID = RRP!RL_ID And iCount < 4
                                    If RRP!DELL_QTYSS > 0 Then
                                        strRef = strRef & IIf(Len(strRef) > 0, ",", "") & FNS(RRP!DEL_SUpplierInvoiceRef) & ": " & Format(RRP!DEL_SUpplierInvoiceDate, "dd/mm/yyyy") & "(" & RRP!DELL_QTYSS & ")"
                                        lngPrice = RRP!DELL_Price
                                        dblDiscount = RRP!DELL_Discount
                                    End If
                                    oPC.COShort.execute "UPDATE tDELL SET DELL_RL_ID = " & lngCurrentRLID & " WHERE DELL_ID = " & RRP!DELL_ID
                                    RRP.MoveNext
                                    iCount = 1 + 1
                                    If RRP.eof Then Exit Do
                                Loop
                                If RRP.eof Then Exit Do
                                oPC.COShort.execute "UPDATE tRL SET RL_SIRef = '" & Right(Replace(strRef, "'", ""), 300) & "' ,RL_Discount = " & dblDiscount & " , RL_Price = " & lngPrice & " WHERE RL_ID = " & lngCurrentRLID
                            Loop
                        End If
                        RRP.Close
                        Set RRP = Nothing
                        rsTR.Close
                        Set rsTR = Nothing
                        rsR.Close
                        Set rsR = Nothing
                        rsRL.Close
                        Set rsRL = Nothing
                If Not rs.eof Then lngTPID = rs!TPID
        Loop
        rs.Close
        Set rs = Nothing
        
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
Exit Function
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.GenerateReturnsPerPub(pTPID,pSpecifiedMonth,pStaffID,bSupplementary)", _
         Array(pTPID, pSpecifiedMonth, pSTAFFID, bSupplementary)
End Function

Public Sub ImportIBTFromFile(pFilePath As String, pBadRecords As Long)
    On Error GoTo errHandler
Dim lngLineCount As Long
Dim strCode As String
Dim strQty As String
Dim lngQty As Long
Dim oBatch As z_SQL
Dim strCodeOut As String
Dim strEANOut As String
Dim strLine As String
Dim fs As New FileSystemObject
Dim arTF() As String
Dim arLine() As String
Dim Res As Long
Dim txtStream As Scripting.TextStream
Dim DestCode As String
Dim lngTFRID As Long

    Set oBatch = New z_SQL
    Set txtStream = fs.OpenTextFile(pFilePath)

    lngLineCount = 0
    Do While Not txtStream.AtEndOfStream
        txtStream.ReadLine
        lngLineCount = lngLineCount + 1
    Loop
    lngLineCount = 0
    txtStream.Close
    Set txtStream = Nothing
    Set txtStream = fs.OpenTextFile(pFilePath)
    If Not txtStream.AtEndOfStream Then
        DestCode = txtStream.ReadLine
    End If
    lngLineCount = 0
    Do While Not txtStream.AtEndOfStream
        strLine = txtStream.ReadLine
        If strLine = "" Then GoTo skip
        If Right(strLine, 1) = "," Then strLine = Left(strLine, Len(strLine) - 1)
        arLine = Split(strLine, ",")
        strCode = Trim(arLine(0))
        If UBound(arLine) = 0 Then
            lngQty = 1
        Else
            lngQty = CLng(Trim(arLine(1)))
        End If
        ReDim arTF(lngLineCount, 2)
        arTF(lngLineCount, 0) = arLine(0)
        If UBound(arLine) > 1 Then
            arTF(lngLineCount, 2) = arLine(1)
        End If
        lngLineCount = lngLineCount + 1

skip:
    Loop


Dim xMLDoc As ujXML
Dim XMLArgs As String
    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "doc_IBT"
            .chCreate "MessageType"
                .elText = "IBT"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "WORKSTATION"
                .elText = oPC.WorkstationName
            .elCreateSibling "StoreCode"
                .elText = DestCode
            .elCreateSibling "INOUT"
                .elText = "OUT"
            .elCreateSibling "DetailLine", True
            For i = 0 To UBound(arTF, 1)
                If arTF(i, 0) > "" Then
                    .chCreate "ITEM"
                    .chCreate "EAN"
                        .elText = arTF(i, 0)
                    .elCreateSibling "Qty", True
                        .elText = "1"
                    .navUP
                    .navUP
                End If
            Next i
         XMLArgs = .docXML
    End With


    lngTFRID = CreateNewTransferFromFile(XMLArgs, "OUT", "Transfer from scanned file")
    If lngTFRID = 0 Then
        MsgBox "There is a problem generating a transfer"
    End If

    Set oBatch = Nothing
EXIT_Handler:
    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.ImportIBTFromFile(pFilePath,pBadRecords)", Array(pFilePath, pBadRecords)
End Sub

Public Sub CreateNewCOStatusReport(TRID As Long, COSTRID As Long, rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
Dim i As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "CreateNewCOStatusReport"
    cmd.commandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.COShort
    
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@TRID", adInteger, adParamInput, , TRID)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@COSRID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    Set rs = cmd.execute

    COSTRID = FNN(cmd.Parameters("@COSRID"))
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CreateNewCOStatusReport(TRID,COSTRID,rs)", Array(TRID, COSTRID, rs)
End Sub

Public Sub GetCOSRLines(TRID As Long, rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
Dim i As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "[GetCOSRLines]"
    cmd.commandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.COShort
    
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@TRID", adInteger, adParamInput, , TRID)
    cmd.Parameters.Append par
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd

    Set cmd = Nothing
'---------------------------------------------------
    'If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.GetCOSRLines(TRID,rs)", Array(TRID, rs)
End Sub

Public Function GenerateCategoryCheck(CategoryID As Long, StaffID As Long, bFilter As Boolean, MBID As Long) As Long
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
Dim i As Integer
Dim Res As Long
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "GenerateCategoryCheck"
    cmd.commandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.COShort
    
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@CATEGID", adInteger, adParamInput, , CategoryID)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@STAFFID", adInteger, adParamInput, , StaffID)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@Filter", adBoolean, adParamInput, , bFilter)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@MB", adInteger, adParamInput, , MBID)
    cmd.Parameters.Append par
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@CATCHKID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.execute
    Res = FNN(cmd.Parameters(4))
    Set cmd = Nothing
    GenerateCategoryCheck = Res
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.GenerateCategoryCheck(CategoryID)", CategoryID
End Function

Public Function UpdateCategoryCheck(Strguid As String) As Boolean
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'    If Len(PXMLArgs) > 20000 Then
'        UpdateCategoryCheck = False
'        MsgBox "section check too big - refer to Papyrus support.", vbInformation, "Can't do this"
'        Exit Function
'    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "Action_CategoryCheck"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Strguid", adVarChar, adParamInput, 50, Strguid)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.UpdateCategoryCheck(Strguid)", Strguid
End Function

Public Function InsertScript(guid As String, Optional XMLScript As String, Optional FreeScript As String)
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.Open "SELECT * FROM tScripts", oPC.COShort, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs.Fields(1) = XMLScript
    rs.Update
    rs.Bookmark = rs.Bookmark
    guid = rs.Fields(0)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.InsertScript(guid,XMLScript,FreeScript)", Array(guid, XMLScript, _
         FreeScript)
End Function

Public Sub CreateSpecialOrder(PXMLArgs As String)
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "[CreateorAddtoCOFromXML]"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@ARGXML", adVarChar, adParamInput, Len(PXMLArgs), PXMLArgs)
    cmd.Parameters.Append par
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
    

End Sub

