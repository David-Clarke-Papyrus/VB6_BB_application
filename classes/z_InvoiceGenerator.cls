VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_OrderFulfilmentDocGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GenerateInvoicesFromCOLALLOCs(pUseDiscountFromOrder As String)
    On Error GoTo errHandler
Dim strSQL As String
Dim strMsg As String
Dim rs As ADODB.Recordset
Dim oI As a_Invoice
Dim oIL As a_InvoiceLine
Dim curTPID As Long
Dim curCOID As Long
Dim curForAttn As String
Dim OpenResult As Integer
Dim bEOINV As Boolean
Dim strParentCode As String

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
Dim strWSName As String

    strWSName = oPC.NameOfPC
    If oPC.IncludeSupplierFeatures Then
        strSQL = "SELECT * FROM vInvoiceGen_FROMGDNs ORDER BY TR_TP_ID,COID"
    Else
        strSQL = "SELECT * FROM vInvoiceGen WHERE ISNULL(COLALLOC_DEL_ID,0) = 0 AND COLALLOC_Action = 1 AND COLALLOC_LOCKEDBY = '" & oPC.NameOfPC & "'  AND COLALLOC_WSNAME = '" & strWSName & "' ORDER BY TR_TP_ID,COID"
    End If
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockReadOnly
    If rs.eof Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
 '   bEOINV = Not rs.EOF
    Do While Not rs.eof
        curTPID = FNN(rs!TR_TP_ID)
        curCOID = FNN(rs!COID)
        curForAttn = FNS(rs!CO_ForAttn)
        'Create Invoice
        Set oI = Nothing
        Set oI = New a_Invoice
        oI.BeginEdit
        oI.SetCustomer FNN(rs!TR_TP_ID)
        oI.VATable = oI.Customer.VATable
        oI.SetStatus stInProcess
        oI.BillToAddressID = FNN(rs!TR_AddressDOCID)
        oI.DelToAddressID = FNN(rs!TR_AddressGoodsID)
        oI.ShowVAT = True
        oI.SetForAttn FNS(rs!CO_ForAttn)
        oI.DOCDate = Date
        oI.CaptureDate = Now()
        bEOINV = False
        Do While bEOINV = False
            Set oIL = oI.InvoiceLines.Add
            oIL.BeginEdit
            If FNS(rs!COLALLOC_SUBSYN) = "Y" Then
                oIL.SubstitutesAvailable = True    'IIf(Trim(rs!COLALLOC_SUBSYN) = "Y", True, False)
                oIL.SetNote "Substitute"
            Else
                oIL.SubstitutesAvailable = False
            End If
            If oI.Customer.oLineToInvoice Then
                oIL.SetNote oIL.Note & " " & FNS(rs!COL_Note)
            End If
            oIL.SetDeposit FNN(rs!COL_Deposit)
            If pUseDiscountFromOrder = "CO" Then
                oIL.SetDiscountPercent FNDBL(rs!COLDIsc)
            Else
                If pUseDiscountFromOrder = "C" Then
                    oIL.SetDiscountPercent FNDBL(rs!CustDisc)
                End If
            End If
            oIL.PID = FNS(rs!P_ID)
            oIL.ServiceItem = False
            oIL.COLID = FNN(rs!COL_ID)
            If IsNull(rs!P_VATRATE) Then
                oIL.VATRate = oPC.Configuration.VATRate
            Else
                oIL.VATRate = CDbl(rs!P_VATRATE)
            End If
                If FNS(rs!COL_Ref) > " " Then
                    oIL.Ref = FNS(rs!COL_Ref)
                End If
            oIL.code = FormatProductCode(FNS(rs!P_Code), False)
            If pUseDiscountFromOrder = "CO" Then
                If oI.Customer.VATable = False And oI.Customer.ShowVAT = False Then
                    oIL.Price = Round(FNN(rs!COL_PRICE) * (100 / (100 + FNN(rs!P_VATRATE))))
                Else
                    oIL.Price = FNN(rs!COL_PRICE)
                End If
            Else
                If FNN(rs!P_SP) > 0 Then
                    If oI.Customer.VATable = False And oI.Customer.ShowVAT = False Then
                        oIL.Price = Round(FNN(rs!P_SP) * (100 / (100 + FNN(rs!P_VATRATE))))
                    Else
                        oIL.Price = FNN(rs!P_SP)
                    End If
                Else
                    If oI.Customer.VATable = False And oI.ShowVAT = False Then
                        oIL.Price = Round(FNN(rs!P_RRP) * (100 / (100 + FNN(rs!P_VATRATE))))
                    Else
                        oIL.Price = FNN(rs!P_RRP)
                    End If
                End If
            End If
            If oPC.AllowsSSInvoicing Then
                oIL.SetQtyFirm FNN(rs!COLALLOC_Qty)
                oIL.SetQtySS FNN(rs!COLALLOC_QtySS)
            Else
                oIL.SetQty FNN(rs!COLALLOC_Qty)
            End If
            If FNDBL(rs!COL_FCFACTOR) > 0 Then
                oIL.SetFCFactor FNDBL(rs!COL_FCFACTOR)
            Else
                oIL.SetFCFactor FNDBL(rs!P_LastFCFactor)
            End If
            If FNN(rs!COL_FCID) > 0 Then
                oIL.FCID = FNN(rs!COL_FCID)
            Else
                oIL.FCID = FNN(rs!P_LastFCID)
            End If
            If FNN(rs!COL_FCPrice) > 0 Then
                oIL.ForeignPrice = FNN(rs!COL_FCPrice)
            Else
                oIL.ForeignPrice = FNN(rs!P_LastForeignPrice)
            End If
            oIL.ApplyEdit
            
            If FNS(rs.fields("COL_AdditionalProductPID")) > "" Then
                strParentCode = formatSKU(rs!P_EAN, rs!P_Code, False)
                Set oIL = oI.InvoiceLines.Add
                oIL.BeginEdit
                oIL.PID = FNS(rs!COL_AdditionalProductPID)
                oIL.ServiceItem = True
                oIL.COLID = FNN(rs!COL_ID)
                oIL.SetQty "1"
                oIL.Price = FNN(rs!COL_AdditionalProductValue)
                oIL.Note = "additional charge for " & strParentCode
                If IsNull(rs!ExtraChargeVATRate) Then
                    oIL.VATRate = oPC.Configuration.VATRate
                Else
                    oIL.VATRate = CDbl(rs!ExtraChargeVATRate)
                End If
                oIL.ApplyEdit
            End If
            
            rs.MoveNext
            If rs.eof Then Exit Do
            If FNB(rs!TP_GenerateSeparateInvoicesForSeparateOrders) = True Then
                If (FNN(rs!TR_TP_ID) <> curTPID) Or (FNN(rs!COID) <> curCOID) Then
                    bEOINV = True
                End If
            Else
                If FNN(rs!TR_TP_ID) <> curTPID Then
                    bEOINV = True
                End If
            End If
            If FNS(rs!CO_ForAttn) <> curForAttn Then
                bEOINV = True
            End If
            If rs!TP_OneLinePerInvoice = True Then
                bEOINV = True
            End If
        Loop
        oI.Post stInProcess
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_OrderFulfilmentDocGenerator.GenerateInvoicesFromCOLALLOCs"
End Function

Public Function GenerateGDNsFromCOLALLOCs(pUseDiscountFromOrder As String)
    On Error GoTo errHandler
Dim strSQL As String
Dim strMsg As String
Dim rs As ADODB.Recordset
Dim oGDN As a_GDN
Dim oGDNL As a_InvoiceLine
Dim curTPID As Long
Dim curCOID As Long
Dim curInvoiceID As Long
Dim curForAttn As String
Dim OpenResult As Integer
Dim bEOINV As Boolean
Dim strParentCode As String

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
Dim strWSName As String

    strWSName = oPC.NameOfPC
    
    strSQL = "SELECT * FROM vInvoiceGen WHERE ISNULL(COLALLOC_DEL_ID,0) = 0 AND COLALLOC_Action = 1 AND COLALLOC_LOCKEDBY = '" & oPC.NameOfPC & "'  AND COLALLOC_WSNAME = '" & strWSName & "' ORDER BY TR_TP_ID,COID"
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockReadOnly
    If rs.eof Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
 '   bEOINV = Not rs.EOF
    Do While Not rs.eof
        curTPID = FNN(rs!TR_TP_ID)
        curCOID = FNN(rs!COID)
        curInvoiceID = FNN(rs!COLALLOC_IID)
        curForAttn = FNS(rs!CO_ForAttn)
        'Create Invoice
        Set oGDN = Nothing
        Set oGDN = New a_GDN
        oGDN.BeginEdit
        oGDN.SetCustomer FNN(rs!TR_TP_ID)
        oGDN.VATable = oGDN.Customer.VATable
        oGDN.SetStatus stInProcess    'IIf(oPC.GetProperty("GenerateInvoicesForGDNsAuto") = "TRUE", stCOMPLETE, stInProcess)
        oGDN.BillToAddressID = FNN(rs!TR_AddressDOCID)
        oGDN.DelToAddressID = FNN(rs!TR_AddressGoodsID)
        oGDN.ShowVAT = True
        oGDN.SetForAttn FNS(rs!CO_ForAttn)
        oGDN.DOCDate = Date
        oGDN.CaptureDate = Now()
        bEOINV = False
        Do While bEOINV = False
            Set oGDNL = oGDN.GDNLines.Add
            'If there is already an invoice generated for this item we should not create a new invoice line but use it in the GDN
            'If however the invoice is a predelivered invoice then we must ask if this GDN matches the invoice lines exactly (qtys) or not.
            'If not then we should create extra lines on the tILine table and link them to the original tILine of the invoice. Any further gdnlines must do the same
            If FNN(rs!COLALLOC_ILID) > 0 And (FNDBL(rs!INVQty) = FNDBL(rs!COLALLOC_Qty)) Then
                oGDNL.InvoiceLineID = FNN(rs!COLALLOC_ILID)
            Else
                oGDNL.InvoiceLineID = FNN(rs!COLALLOC_ILID) * -1
            End If
            oGDNL.BeginEdit
            If FNS(rs!COLALLOC_SUBSYN) = "Y" Then
                oGDNL.SubstitutesAvailable = True    'IIf(Trim(rs!COLALLOC_SUBSYN) = "Y", True, False)
                oGDNL.SetNote "Substitute"
            Else
                oGDNL.SubstitutesAvailable = False
            End If
            If oGDN.Customer.oLineToInvoice Then
                oGDNL.SetNote oGDNL.Note & " " & FNS(rs!COL_Note)
            End If
            oGDNL.SetDeposit FNN(rs!COL_Deposit)
            If pUseDiscountFromOrder = "CO" Then
                oGDNL.SetDiscountPercent FNDBL(rs!COLDIsc)
            Else
                If pUseDiscountFromOrder = "C" Then
                    oGDNL.SetDiscountPercent FNDBL(rs!CustDisc)
                End If
            End If
            oGDNL.PID = FNS(rs!P_ID)
            oGDNL.ServiceItem = False
            oGDNL.COLID = FNN(rs!COL_ID)
            If IsNull(rs!P_VATRATE) Then
                oGDNL.VATRate = oPC.Configuration.VATRate
            Else
                oGDNL.VATRate = CDbl(rs!P_VATRATE)
            End If
                If FNS(rs!COL_Ref) > " " Then
                    oGDNL.Ref = FNS(rs!COL_Ref)
                End If
            oGDNL.code = FormatProductCode(FNS(rs!P_Code), False)
            If pUseDiscountFromOrder = "CO" Then
                If oGDN.Customer.VATable = False And oGDN.Customer.ShowVAT = False Then
                    oGDNL.Price = Round(FNN(rs!COL_PRICE) * (100 / (100 + FNN(rs!P_VATRATE))))
                Else
                    oGDNL.Price = FNN(rs!COL_PRICE)
                End If
            Else
                If FNN(rs!P_SP) > 0 Then
                    If oGDN.Customer.VATable = False And oGDN.Customer.ShowVAT = False Then
                        oGDNL.Price = Round(FNN(rs!P_SP) * (100 / (100 + FNN(rs!P_VATRATE))))
                    Else
                        oGDNL.Price = FNN(rs!P_SP)
                    End If
                Else
                    If oGDN.Customer.VATable = False And oGDN.ShowVAT = False Then
                        oGDNL.Price = Round(FNN(rs!P_RRP) * (100 / (100 + FNN(rs!P_VATRATE))))
                    Else
                        oGDNL.Price = FNN(rs!P_RRP)
                    End If
                End If
            End If
            If oPC.AllowsSSInvoicing Then
                oGDNL.SetQtyFirm FNN(rs!COLALLOC_Qty)
                oGDNL.SetQtySS FNN(rs!COLALLOC_QtySS)
            Else
                oGDNL.SetQty FNN(rs!COLALLOC_Qty)
            End If
            If FNDBL(rs!COL_FCFACTOR) > 0 Then
                oGDNL.SetFCFactor FNDBL(rs!COL_FCFACTOR)
            Else
                oGDNL.SetFCFactor FNDBL(rs!P_LastFCFactor)
            End If
            If FNN(rs!COL_FCID) > 0 Then
                oGDNL.FCID = FNN(rs!COL_FCID)
            Else
                oGDNL.FCID = FNN(rs!P_LastFCID)
            End If
            If FNN(rs!COL_FCPrice) > 0 Then
                oGDNL.ForeignPrice = FNN(rs!COL_FCPrice)
            Else
                oGDNL.ForeignPrice = FNN(rs!P_LastForeignPrice)
            End If
            oGDNL.ApplyEdit
            
            If FNS(rs.fields("COL_AdditionalProductPID")) > "" Then
                strParentCode = formatSKU(rs!P_EAN, rs!P_Code, False)
                Set oGDNL = oGDN.GDNLines.Add
                oGDNL.BeginEdit
                oGDNL.PID = FNS(rs!COL_AdditionalProductPID)
                oGDNL.ServiceItem = True
                oGDNL.COLID = FNN(rs!COL_ID)
                oGDNL.SetQty "1"
                oGDNL.Price = FNN(rs!COL_AdditionalProductValue)
                oGDNL.Note = "additional charge for " & strParentCode
                If IsNull(rs!ExtraChargeVATRate) Then
                    oGDNL.VATRate = oPC.Configuration.VATRate
                Else
                    oGDNL.VATRate = CDbl(rs!ExtraChargeVATRate)
                End If
                oGDNL.ApplyEdit
            End If
            
            rs.MoveNext
            If rs.eof Then Exit Do
            If FNB(rs!TP_GenerateSeparateInvoicesForSeparateOrders) = True Then
                If (FNN(rs!TR_TP_ID) <> curTPID) Or (FNN(rs!COID) <> curCOID) Then
                    bEOINV = True
                End If
            Else
                If FNN(rs!TR_TP_ID) <> curTPID Then
                    bEOINV = True
                End If
            End If
            If FNS(rs!CO_ForAttn) <> curForAttn Then
                bEOINV = True
            End If
            If rs!TP_OneLinePerInvoice = True Then
                bEOINV = True
            End If
            If curInvoiceID <> FNN(rs!COLALLOC_IID) Then
                bEOINV = True
            End If
            
        Loop
        oGDN.Post stInProcess
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_OrderFulfilmentDocGenerator.GenerateInvoicesFromCOLALLOCs"
End Function



Public Function GenerateApprosFromCOLALLOCs(pUseDiscountFromOrder As String)
    On Error GoTo errHandler
Dim strSQL As String
Dim strMsg As String
Dim rs As ADODB.Recordset
Dim oAPP As a_APP
Dim oAPPL As a_APPL
Dim curTPID As Long
Dim curCOID As Long
Dim curForAttn As String
Dim OpenResult As Integer
Dim bEOINV As Boolean
Dim strParentCode As String

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
Dim strWSName As String

    strWSName = oPC.NameOfPC
    
    strSQL = "SELECT * FROM vInvoiceGen WHERE ISNULL(COLALLOC_DEL_ID,0) = 0 AND COLALLOC_Action = 1 AND COLALLOC_LOCKEDBY = '" & oPC.NameOfPC & "'  AND COLALLOC_WSNAME = '" & strWSName & "' ORDER BY TR_TP_ID,COID"
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockReadOnly
    If rs.eof Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
 '   bEOINV = Not rs.EOF
    Do While Not rs.eof
        curTPID = FNN(rs!TR_TP_ID)
        curCOID = FNN(rs!COID)
        curForAttn = FNS(rs!CO_ForAttn)
        'Create Invoice
        Set oAPP = Nothing
        Set oAPP = New a_APP
        oAPP.BeginEdit
        oAPP.SetCustomer curTPID
        oAPP.VATable = oAPP.Customer.VATable
        oAPP.SetStatus stInProcess
'        oApp. = FNN(rs!TR_AddressDOCID)
'        oApp.SetApproToAddress FNN(rs!TR_AddressGoodsID)
        oAPP.ShowVAT = True
     '  oApp.SetForAttn FNS(rs!CO_ForAttn)
        oAPP.DOCDate = Date
      '  oApp.capturedate Now()
        bEOINV = False
        Do While bEOINV = False
            Set oAPPL = oAPP.ApproLines.Add
            oAPPL.BeginEdit
            If FNS(rs!COLALLOC_SUBSYN) = "Y" Then
                oAPPL.SubstitutesAvailable = True    'IIf(Trim(rs!COLALLOC_SUBSYN) = "Y", True, False)
                oAPPL.Note = "Substitute"
            Else
                oAPPL.SubstitutesAvailable = False
            End If
            If oAPP.Customer.oLineToInvoice Then
                oAPPL.Note = oAPPL.Note & " " & FNS(rs!COL_Note)
            End If
         '  oAppL FNN(rs!COL_Deposit)
            If pUseDiscountFromOrder = "CO" Then
                oAPPL.SetDiscount FNDBL(rs!COLDIsc)
            Else
                If pUseDiscountFromOrder = "C" Then
                    oAPPL.SetDiscount FNDBL(rs!CustDisc)
                End If
            End If
            oAPPL.PID = FNS(rs!P_ID)
   '         oAppL.ServiceItem = False
            oAPPL.COLID = FNN(rs!COL_ID)
            If IsNull(rs!P_VATRATE) Then
                oAPPL.VATRate = oPC.Configuration.VATRate
            Else
                oAPPL.VATRate = CDbl(rs!P_VATRATE)
            End If
                If FNS(rs!COL_Ref) > " " Then
                    oAPPL.Ref = FNS(rs!COL_Ref)
                End If
            oAPPL.code = FormatProductCode(FNS(rs!P_Code), False)
            If pUseDiscountFromOrder = "CO" Then
                If oAPP.Customer.VATable = False And oAPP.Customer.ShowVAT = False Then
                    oAPPL.Price = Round(FNN(rs!COL_PRICE) * (100 / (100 + FNN(rs!P_VATRATE))))
                Else
                    oAPPL.Price = FNN(rs!COL_PRICE)
                End If
            Else
                If FNN(rs!P_SP) > 0 Then
                    If oAPP.Customer.VATable = False And oAPP.Customer.ShowVAT = False Then
                        oAPPL.Price = Round(FNN(rs!P_SP) * (100 / (100 + FNN(rs!P_VATRATE))))
                    Else
                        oAPPL.Price = FNN(rs!P_SP)
                    End If
                Else
                    If oAPP.Customer.VATable = False And oAPP.ShowVAT = False Then
                        oAPPL.Price = Round(FNN(rs!P_RRP) * (100 / (100 + FNN(rs!P_VATRATE))))
                    Else
                        oAPPL.Price = FNN(rs!P_RRP)
                    End If
                End If
            End If
            If oPC.AllowsSSInvoicing Then
                oAPPL.SetQty FNN(rs!COLALLOC_Qty) + FNN(rs!COLALLOC_QtySS)
              '  oAppL.SetQtySS FNN(rs!COLALLOC_QtySS)
            Else
                oAPPL.SetQty FNN(rs!COLALLOC_Qty)
            End If
'            If FNDBL(rs!COL_FCFACTOR) > 0 Then
'                oAppL.SetFCFactor FNDBL(rs!COL_FCFACTOR)
'            Else
'                oAppL.SetFCFactor FNDBL(rs!P_LastFCFactor)
'            End If
'            If FNN(rs!COL_FCID) > 0 Then
'                oAppL.FCID = FNN(rs!COL_FCID)
'            Else
'                oAppL.FCID = FNN(rs!P_LastFCID)
'            End If
'            If FNN(rs!COL_FCPrice) > 0 Then
'                oAppL.ForeignPrice = FNN(rs!COL_FCPrice)
'            Else
'                oAppL.ForeignPrice = FNN(rs!P_LastForeignPrice)
'            End If
            oAPPL.ApplyEdit
            
'            If FNS(rs.Fields("COL_AdditionalProductPID")) > "" Then
'                strParentCode = formatSKU(rs!P_EAN, rs!P_Code, False)
'                Set oAppL = oApp.ApproLines.Add
'                oAppL.BeginEdit
'                oAppL.PID = FNS(rs!COL_AdditionalProductPID)
'                oAppL.ServiceItem = True
'                oAppL.COLID = FNN(rs!COL_ID)
'                oAppL.SetQty "1"
'                oAppL.Price = FNN(rs!COL_AdditionalProductValue)
'                oAppL.Note = "additional charge for " & strParentCode
'                If IsNull(rs!ExtraChargeVATRate) Then
'                    oAppL.VATRate = oPC.Configuration.VATRate
'                Else
'                    oAppL.VATRate = CDbl(rs!ExtraChargeVATRate)
'                End If
'                oAppL.ApplyEdit
'            End If
            
            rs.MoveNext
            If rs.eof Then Exit Do
'            If FNB(rs!TP_GenerateSeparateInvoicesForSeparateOrders) = True Then
'                If (FNN(rs!TR_TP_ID) <> curTPID) Or (FNN(rs!COID) <> curCOID) Then
'                    bEOINV = True
'                End If
'            Else
                If FNN(rs!TR_TP_ID) <> curTPID Then
                    bEOINV = True
                End If
'            End If
            If FNS(rs!CO_ForAttn) <> curForAttn Then
                bEOINV = True
            End If
'            If rs!TP_OneLinePerInvoice = True Then
'                bEOINV = True
'            End If
        Loop
        oAPP.Post
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_OrderFulfilmentDocGenerator.GenerateApprosFromCOLALLOCs"
End Function

Public Function FetchOrderLinesToFulfil(WithCompleteRequirement As Boolean) As ADODB.Recordset
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim rs As ADODB.Recordset

Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetAllUInissuedGDNs"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    
    Set par = cmd.CreateParameter("@WithCompleteRequirement", adBoolean, adParamInput)
    cmd.Parameters.Append par
    par.Value = WithCompleteRequirement
   
    cmd.ActiveConnection = oPC.COShort
    
    Set rs = New ADODB.Recordset
    rs.open cmd
    Set FetchOrderLinesToFulfil = rs
    
    Set cmd = Nothing
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_OrderFulfilmentDocGen.FetchOrderLinesToFulfil(WithCompleteRequirement)", _
         WithCompleteRequirement, EA_NORERAISE
End Function

Public Function CustomerDocsToDispatch() As ADODB.Recordset
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim rs As ADODB.Recordset

Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "SELECT * FROM vCustomerDocsToDispatch"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    
   
    cmd.ActiveConnection = oPC.COShort
    
    Set rs = New ADODB.Recordset
    rs.open cmd
    Set CustomerDocsToDispatch = rs
    
    Set cmd = Nothing
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_OrderFulfilmentDocGen.CustomerDocsToDispatch", , EA_NORERAISE
End Function
