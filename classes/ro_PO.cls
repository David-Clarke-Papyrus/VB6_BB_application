VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ro_PO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lngLPLessDiscExt_tot As Long
Dim lngFPLessDiscExt_tot As Long
Dim lngLPExt_tot As Long
Dim lngFPExt_tot As Long
Dim lngVAT_tot As Long
Dim lngDisc_tot As Long
Dim mOriginalStatus As enStatus
Dim oPCL As z_PrintClient
Event Valid(pMsg As String)
Event TotalChange(lngTotal As String, lngTotalForeign As String)
Event reloadlist()
Event CurrRowStatus(pMsg As String)
Event Dirty(pVal As Boolean)
Private mudtProps As POProps
Private mcolStack As Collection
Private mPOL As ro_POLs
Attribute mPOL.VB_VarHelpID = -1
Private WithEvents mSupplier As a_Supplier
Attribute mSupplier.VB_VarHelpID = -1
Private colClassors As Collection
Private oBillToAddress As a_Address
Private oDelToAddress As a_Address
Private oBillingCompany As a_Company
'Private CaptureCurrency As a_Currency
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private moCaptureInCurrency As a_Currency 'usually this will be the system default currency but it could be a foreign currency
' if we are ordering from a foreign company and wish to reflect our understanding of their prices in their currency.
' The currency used and the conversionrate as at time of issuing PO are stored with the PO.
' We can thus record differnces between the ordered and delivered prices which reflect the currency flustuations.
Const CDOCCODE = "PO"

Public Property Get CaptureCurrency() As a_Currency
    Set CaptureCurrency = moCaptureInCurrency
End Property

Public Function CalculateTotals()
    On Error GoTo errHandler
    CalculateTotal_Local
    CalculateTotal_Foreign
    RaiseEvent TotalChange(Me.TotalLessDiscExtF(False), Me.TotalLessDiscExtF(True))
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.CalculateTotals"
End Function
Public Function CalculateTotal_Local()
    On Error GoTo errHandler
Dim ol As ro_POL
Dim bLineDiscountApplied As Boolean
    bLineDiscountApplied = False
    lngLPLessDiscExt_tot = 0
    lngLPExt_tot = 0
    lngVAT_tot = 0
    lngDisc_tot = 0
    'NOTE Some items on the invoice might be ServiceItem e.g. Insurance, Postage and so on. These items are non subject to any discount
    'and they may mut probably will not carry VAT. e.g. a handling charge would include VAT
    'VAT is calculated on the discounted price, so it is necessary to apply the discount before calculating VAT
    'Both forms of discount (by line or overall(mutually exclusive)) will need to be applied line by line, and the VAT for that line
    'calculated afterwards. Remember VAT rate can vary per line.
    'Any general discount will merely be distributed as a line by line discount on all non ServiceItem items
    'Any deposits are ignored until the last when the total deposit is subtracted from the total payable figure
    For Each ol In mPOL
        If ol.Fulfilled <> "CAN" Then
            ol.RecalculateLine
            lngLPExt_tot = lngLPExt_tot + ol.PExt(False)
            lngVAT_tot = lngVAT_tot + ol.PLessDiscExtVAT(False)
            lngDisc_tot = lngDisc_tot + ol.PDiscExt(False)
            lngLPLessDiscExt_tot = lngLPLessDiscExt_tot + ol.PLessDiscExt(False)
        End If
    Next
    mudtProps.TotalExtensionSimple = lngLPExt_tot
    mudtProps.TotalDiscount = lngDisc_tot
    mudtProps.TotalVAT = lngVAT_tot
    mudtProps.TotalExtension = lngLPLessDiscExt_tot
    mudtProps.TotalPayable = lngLPLessDiscExt_tot ' - lngLDeposit_tot
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.CalculateTotal_Local"
End Function
Public Function CalculateTotal_Foreign()
    On Error GoTo errHandler
Dim ol As ro_POL
Dim bLineDiscountApplied As Boolean
    bLineDiscountApplied = False
    lngFPLessDiscExt_tot = 0
    lngFPExt_tot = 0
    lngVAT_tot = 0
    lngDisc_tot = 0
    For Each ol In mPOL
        If ol.Fulfilled <> "CAN" Then
            lngFPExt_tot = lngFPExt_tot + ol.PExt(True)
            lngVAT_tot = lngVAT_tot + ol.PLessDiscExtVAT(True)
            lngDisc_tot = lngDisc_tot + ol.PDiscExt(True)
            lngFPLessDiscExt_tot = lngFPLessDiscExt_tot + ol.PLessDiscExt(True)
        End If
    Next
    mudtProps.TotalExtensionSimple_Foreign = lngFPExt_tot
    mudtProps.TotalDiscount_Foreign = lngDisc_tot
    mudtProps.TotalVAT_Foreign = lngVAT_tot
    mudtProps.TotalExtension_Foreign = lngFPLessDiscExt_tot
    mudtProps.TotalPayable_Foreign = lngFPLessDiscExt_tot ' - lngFDeposit_tot
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.CalculateTotal_Foreign"
End Function
Public Property Get TotalLessDiscExtF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalLessDiscExtF = Format(mudtProps.TotalExtension_Foreign / CaptureCurrency.Divisor, CaptureCurrency.FormatString)
    Else
        TotalLessDiscExtF = Format(mudtProps.TotalExtension / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.TotalLessDiscExtF(bForeign)", bForeign
End Property
Public Property Get TotalExtF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalExtF = Format(mudtProps.TotalExtensionSimple_Foreign / CaptureCurrency.Divisor, CaptureCurrency.FormatString)
    Else
        TotalExtF = Format(mudtProps.TotalExtensionSimple / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.TotalExtF(bForeign)", bForeign
End Property

Public Property Get TotalLessDiscExt(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalLessDiscExt = mudtProps.TotalExtension_Foreign
    Else
        TotalLessDiscExt = mudtProps.TotalExtension
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.TotalLessDiscExt(bForeign)", bForeign
End Property
Public Property Get TotalVATF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalVATF = Format(mudtProps.TotalVAT_Foreign / CaptureCurrency.Divisor, CaptureCurrency.FormatString)
    Else
        TotalVATF = Format(mudtProps.TotalVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.TotalVATF(bForeign)", bForeign
End Property
Public Property Get TotalPayable(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalPayable = mudtProps.TotalPayable_Foreign / CaptureCurrency.Divisor
    Else
        TotalPayable = mudtProps.TotalPayable / oPC.Configuration.DefaultCurrency.Divisor
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.TotalPayable(bForeign)", bForeign
End Property
Public Property Get TotalPayableF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalPayableF = Format(TotalPayable(bForeign), CaptureCurrency.FormatString)
    Else
        TotalPayableF = Format(TotalPayable(bForeign), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.TotalPayableF(bForeign)", bForeign
End Property
Public Property Get TotalDiscount(bForeign As Boolean) As Long
    On Error GoTo errHandler
    If bForeign Then
        TotalDiscount = mudtProps.TotalDiscount_Foreign
    Else
        TotalDiscount = mudtProps.TotalDiscount
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.TotalDiscount(bForeign)", bForeign
End Property
Public Property Get TotalDiscountF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalDiscountF = Format(mudtProps.TotalDiscount_Foreign / CaptureCurrency.Divisor, CaptureCurrency.FormatString)
    Else
        TotalDiscountF = Format(mudtProps.TotalDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.TotalDiscountF(bForeign)", bForeign
End Property
'Public Property Get TotalDepositF(bForeign As Boolean) As String
'    If bForeign Then
'        TotalDepositF = Format(TotalDeposit(bForeign), CaptureCurrency.FormatString)
'    Else
'        TotalDepositF = Format(TotalDeposit(bForeign), oPC.Configuration.DefaultCurrency.FormatString)
'    End If
'End Property
'Public Property Get TotalDeposit(bForeign As Boolean) As Double
'    If bForeign Then
'        TotalDeposit = lngFDeposit_tot / CaptureCurrency.Divisor
'    Else
'        TotalDeposit = lngLDeposit_tot / oPC.Configuration.DefaultCurrency.Divisor
'    End If
'End Property
Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As POData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.GetState"
End Function
Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As POData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.SetState(buffer)", buffer
End Sub
Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    mudtProps.CurrencyID = moCaptureInCurrency.ID
    .WriteProperty "State", GetState
    .WriteProperty "POLs", mPOL.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.GetSuperState"
End Function
Public Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  
  Set objPB = New PropertyBag
  arBuffer = buffer
  With objPB
    .Contents = arBuffer
    SetState .ReadProperty("State")
    mPOL.SetSuperState .ReadProperty("POLS")
    For Each moCaptureInCurrency In oPC.Configuration.Currencies
        If moCaptureInCurrency.ID = mudtProps.CurrencyID Then
            Exit For
        End If
    Next
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.SetSuperState(buffer)", buffer
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mPOL = New ro_POLs
    mPOL.component Me
    Set mSupplier = New a_Supplier
    Set oPCL = New z_PrintClient
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set colClassors = Nothing
    Set mcolStack = Nothing
    Set mobjValid = Nothing
    Set mPOL = Nothing
    Set mSupplier = Nothing
    Set oPCL = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.Class_Terminate"
End Sub


Public Sub Load(TRID As Long, ReadOnly As Boolean)
    On Error GoTo errHandler
Dim objPersist As ro_PO_P
    Set objPersist = New ro_PO_P
    SetSuperState objPersist.Fetch(TRID)
    mSupplier.Load mudtProps.TPID
    Set objPersist = Nothing
    CalculateTotals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.Load(TRID,ReadOnly)", Array(TRID, ReadOnly)
End Sub

Public Sub Reload()
    On Error GoTo errHandler
Dim lngTRID As Long
    lngTRID = Me.TRID
    Class_Terminate
    Class_Initialize
    Load lngTRID, True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.Reload"
End Sub
Public Property Get Supplier() As a_Supplier
    On Error GoTo errHandler
    Set Supplier = mSupplier
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.Supplier"
End Property
Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.Delete"
End Sub



Public Property Get SOID() As Long
    On Error GoTo errHandler
    SOID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.SOID"
End Property
Public Property Let POID(val As Long)
    On Error GoTo errHandler
    mudtProps.TRID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.POID(val)", val
End Property

Public Property Get BillToAddressID() As Long
    On Error GoTo errHandler
    BillToAddressID = mudtProps.DocStoreID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.BillToAddressID"
End Property
Public Property Get DELTOStoreID() As Long
    On Error GoTo errHandler
    DELTOStoreID = mudtProps.DELTOStoreID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.DELTOStoreID"
End Property

Public Property Get TRID() As Long
    On Error GoTo errHandler
    TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.TRID"
End Property
Public Property Get ContainsCO() As Boolean
    On Error GoTo errHandler
    ContainsCO = mudtProps.ContainsCO
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.ContainsCO"
End Property
Public Function SetCaptureCurrency(val As a_Currency) As Boolean
    On Error GoTo errHandler
    Set moCaptureInCurrency = val
    mudtProps.CurrencyID = val.ID
    mudtProps.CurrencyRate = val.Factor
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.SetCaptureCurrency(val)", val
End Function
Public Property Get StaffID() As Long
    On Error GoTo errHandler
    StaffID = mudtProps.StaffID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.StaffID"
End Property
Public Property Get StaffName() As String
    On Error GoTo errHandler
    If StaffID < 1 Then
        StaffName = ""
        Exit Property
    End If
    StaffName = oPC.Configuration.Staff.FindStaffByID(StaffID).Shortname
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.StaffName"
End Property
Public Property Get StaffNameB() As String
    On Error GoTo errHandler
Dim strT As String
    If StaffID < 1 Then
        StaffNameB = ""
        Exit Property
    End If
    strT = oPC.Configuration.Staff.FindStaffByID(StaffID).Shortname
    If strT > "" Then
        StaffNameB = "   (" & strT & ")"
    Else
        StaffNameB = ""
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.StaffNameB"
End Property

Public Property Get BillingCompany() As a_Company
    Set BillingCompany = oBillingCompany
End Property
Public Property Get TPNAME() As String
    On Error GoTo errHandler
    TPNAME = FNS(mudtProps.TPNAME)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.TPName"
End Property

Public Property Get TPPhone() As String
    On Error GoTo errHandler
    TPPhone = FNS(mudtProps.TPPhone)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.TPPhone"
End Property

Public Property Get TPFax() As String
    On Error GoTo errHandler
    TPFax = FNS(mudtProps.TPFax)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.TPFax"
End Property
Public Property Get Memo() As String
    On Error GoTo errHandler
    Memo = FNS(mudtProps.Memo)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.Memo"
End Property
Public Property Let Memo(val As String)
    On Error GoTo errHandler
    mudtProps.Memo = FNS(val)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.Memo(val)", val
End Property

Public Property Get OrderType() As String
    On Error GoTo errHandler
    OrderType = FNS(mudtProps.OrderType)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.OrderType"
End Property
Public Property Let OrderType(val As String)
    On Error GoTo errHandler
    mudtProps.OrderType = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.OrderType(val)", val
End Property

Public Property Get DOCCode() As String
    On Error GoTo errHandler
    DOCCode = FNS(mudtProps.DOCCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.DocCode"
End Property
Public Property Get TPACCNum() As String
    On Error GoTo errHandler
    TPACCNum = FNS(mudtProps.TPACCNum)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.TPACCNum"
End Property
Public Property Get DeliverToAddress() As String
    On Error GoTo errHandler
    DeliverToAddress = mudtProps.DelToAddress
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.DeliverToAddress"
End Property

Public Property Get DOCDate() As Date
    On Error GoTo errHandler
    DOCDate = mudtProps.DOCDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.DocDate"
End Property
Public Property Get DocDateF() As String
    On Error GoTo errHandler
    DocDateF = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.DocDateF"
End Property

Public Property Get IssDate() As Date
    On Error GoTo errHandler
    IssDate = mudtProps.CaptureDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.IssDate"
End Property
Public Property Get IssDateF() As Date
    On Error GoTo errHandler
    IssDateF = Format(mudtProps.CaptureDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.IssDateF"
End Property

Public Property Get OrderDate() As Date
    On Error GoTo errHandler
    OrderDate = mudtProps.DOCDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.OrderDate"
End Property
Public Property Get OrderDateFormatted() As String
    On Error GoTo errHandler
    OrderDateFormatted = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.OrderDateFormatted"
End Property

Public Property Get POLines() As ro_POLs
  Set POLines = mPOL
End Property
Public Property Get DisplayPO() As d_PO
    On Error GoTo errHandler
    Set DisplayPO = New d_PO
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.DisplayPO"
End Property
Public Property Get StatusF() As String
    On Error GoTo errHandler
    If mudtProps.Status = stVOID Then
        StatusF = "VOID"
    ElseIf mudtProps.Status = stInProcess Then
        StatusF = "IN PROCESS"
    ElseIf mudtProps.Status = stISSUED Then
        StatusF = "ISSUED"
    ElseIf mudtProps.Status = stCOMPLETE Then
        StatusF = "COMPLETE"
    ElseIf mudtProps.Status = stPROFORMA Then
        StatusF = "PROFORMA"
    ElseIf mudtProps.Status = stCANCELLED Then
        StatusF = "CANCELLED"
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.StatusF"
End Property
Public Property Get Status() As enStatus
    On Error GoTo errHandler
    Status = mudtProps.Status
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.Status"
End Property
Public Function FetchPONumber(strTmp As String) As String
    On Error GoTo errHandler
Dim oCode As z_Code
Dim strDefaultCo As String
Dim strPONumber As String
    Set oCode = New z_Code
    
    strPONumber = oCode.GetNextCode(enSupplierOrder, strTmp)
    
    strDefaultCo = oPC.Configuration.DefaultCompany.CompanyCode
    FetchPONumber = strDefaultCo & "P" & strPONumber
    
    Set oCode = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.FetchPONumber(strTmp)", strTmp
End Function

Public Function PrintPO_Display(bForeign As Boolean) As String
    On Error GoTo errHandler
Dim bDiscountExists As Boolean
Dim oPO As ro_POL
Dim strPO As String
Dim i As Integer

    
    strPO = "Purchase order: " & Me.DOCCode & vbCrLf
    
    For i = 1 To Me.POLines.Count
        With Me.POLines(i)
            strPO = strPO & .ProductCodeF & "  " & .Ref & "  " & .QtyFirmF & "  " & .QtySS & "  " & Left(.TitleAuthor, 30) & "  " & .PriceF(bForeign) & "  " & IIf(.Discount = 0, "", .DiscountF) & "  " & .PLessDiscExtF(bForeign)
        End With
        strPO = strPO & vbCrLf
    Next i
    
    strPO = strPO & "Total   " & TotalPayableF(bForeign) & vbCrLf
    
    strPO = strPO & vbCrLf
   
    PrintPO_Display = strPO
EXIT_Handler:
'Err_Handler:
'    oPCL.FinishRequest
'    Select Case
'    Case 5941
'        MsgBox "Bookmark missing"
'        Resume Next
'    Case Else
'        MsgBox error
'        GoTo EXIT_Handler
'        Resume
'    End Select
'
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.PrintPO_Display(bForeign)", bForeign
End Function
Public Function PrintPO(pShowVAT As Boolean, bForeign As Boolean, Optional ByVal pRoundedUp As Boolean) As Boolean
    On Error GoTo errHandler
Dim bDiscountExists As Boolean
Dim tmp As String

    Set oPCL = New z_PrintClient
    If oPCL.StartRequestPrinting("PO_" & Me.DOCCode) = False Then
        PrintPO = False
        Exit Function
    End If
    
    With oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
        oPCL.PrinterSetup .PrinterName(), .PreviewPrint, .Style, .QtyCopies
    End With
    
    If Me.StatusF <> "In PROCESS" And Me.StatusF <> "COMPLETE" Then
        oPCL.WriteTag "STATUS", Me.StatusF
    End If
    oPCL.LineCOunt POLines.Count
    oPCL.LineTotalCount 1
    oPCL.WriteTag "TOTAL", "Total" & "|" & Me.TotalPayableF(bForeign)
    oPCL.WriteTag "DOCUMENTCODE", Me.DOCCode
    oPCL.WriteTag "SUPPLIER", Supplier.NameAndCode(35)
    oPCL.WriteTag "PHONE", IIf(Me.Supplier.BillTOAddress.Phone > "", "Phone: " & Supplier.BillTOAddress.Phone, "")
    oPCL.WriteTag "FAX", IIf(Supplier.BillTOAddress.Fax > "", "Fax: " & Supplier.BillTOAddress.Fax, "")
   ' tmp = Replace(oPC.Configuration.DefaultStore.BillAddress, vbCrLf, "%")
    oPCL.WriteTag "BILLTO", PackText(oPC.Configuration.DefaultStore.BillAddress)
   ' tmp = Replace(oPC.Configuration.Stores.FindStoreByID(Me.DELTOStoreID).DelAddress, vbCrLf, "%")
    oPCL.WriteTag "DELTO", PackText(oPC.Configuration.Stores.FindStoreByID(Me.DELTOStoreID).DelAddress)
    oPCL.WriteTag "DATE", Me.DOCDate
    oPCL.WriteTag "MAINMSG", oPC.Configuration.OrderText
    
    FillTableRows Me, bForeign, bDiscountExists
    
 '   If Me.Supplier.VATable Then ''''''''''(1)
 '       oPCL.WriteTag "TOTAL", "Total" & "|" & Me.TotalLessDiscExtF(bForeign)
 '       oPCL.WriteTag "TOTAL", "Includes VAT of " & "|" & Me.TotalVATF(bForeign)
 '   Else  'overseas clients
 '           oPCL.WriteTag "TOTAL", "subtotal" & "|" & Me.TotalLessDiscExtF(bForeign)
 '           oPCL.WriteTag "TOTAL2", "Less VAT of " & "|" & Me.TotalVATF(bForeign)
 '           oPCL.WriteTag "TOTAL3", "Payable " & "|" & Me.TotalPayableF(bForeign)
 '       End If
 '   End If
    oPCL.FinishRequest
    
EXIT_Handler:
'Err_Handler:
'    oPCL.FinishRequest
'    Select Case
'    Case 5941
'        MsgBox "Bookmark missing"
'        Resume Next
'    Case Else
'        MsgBox error
'        GoTo EXIT_Handler
'        Resume
'    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.PrintPO(pShowVAT,bForeign,pRoundedUp)", Array(pShowVAT, bForeign, pRoundedUp)
End Function
Public Function DisplayTotals(pLabels As String, pValues As String, bForeign As Boolean) 'This must be the same as PRINTINVOICE
    On Error GoTo errHandler
Dim strLabels As String
Dim strValues As String
    strLabels = ""
    strValues = ""
    If bForeign Then
        If oPC.Configuration.IsVATRegion Then
                strLabels = strLabels & Chr(13) & "Total"
                strValues = strValues & Chr(13) & Me.TotalPayableF(bForeign)
        Else
                strLabels = strLabels & Chr(13) & "subtotal"
                strValues = strValues & Chr(13) & Me.TotalExtF(bForeign)
                strLabels = strLabels & Chr(13) & "Less discount"
                strValues = strValues & Chr(13) & Me.TotalDiscountF((bForeign))
                strLabels = strLabels & Chr(13) & "Total"
                strValues = strValues & Chr(13) & Me.TotalPayableF(bForeign)
        End If
    Else
        If oPC.Configuration.IsVATRegion Then
                strLabels = strLabels & Chr(13) & "Total"
                strValues = strValues & Chr(13) & Me.TotalExtF(bForeign)
                strLabels = strLabels & Chr(13) & "Less discount"
                strValues = strValues & Chr(13) & Me.TotalDiscountF((bForeign))
                strLabels = strLabels & Chr(13) & "Includes VAT of "
                strValues = strValues & Chr(13) & Me.TotalVATF(bForeign)
        Else
                strLabels = strLabels & Chr(13) & "subtotal"
                strValues = strValues & Chr(13) & Me.TotalExtF(bForeign)
                strLabels = strLabels & Chr(13) & "Less discount"
                strValues = strValues & Chr(13) & Me.TotalDiscountF((bForeign))
                strLabels = strLabels & Chr(13) & "Total"
                strValues = strValues & Chr(13) & Me.TotalPayableF(bForeign)
        End If
    End If
    pLabels = strLabels
    pValues = strValues
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.DisplayTotals(pLabels,pValues,bForeign)", Array(pLabels, pValues, bForeign)
End Function
Private Sub FillTableRows(pPO As a_PO, bForeign As Boolean, pDiscountExists As Boolean)
    On Error GoTo errHandler
Dim i As Integer

    For i = 1 To pPO.POLines.Count
        With pPO.POLines(i)
            If .Discount > 0 Then pDiscountExists = True
            oPCL.POLSend .ProductCodeF, .QtyFirmF, .QtySSF, .TitleAuthor, .PriceF(bForeign), .PLessDiscExtF(bForeign), .DiscountF, .Note, .Ref '"InvDate", oInv.TransDateFormatted
        End With
    Next i
EXIT_Handler:
'Err_Handler:
'    MsgBox error
'    GoTo EXIT_Handler
   ' Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.FillTableRows(pPO,bForeign,pDiscountExists)", Array(pPO, bForeign, _
         pDiscountExists)
End Sub


Public Property Get CanCancel() As Boolean
    On Error GoTo errHandler
Dim bCanCancel As Boolean
Dim oPOL As ro_POL
    bCanCancel = True
    For Each oPOL In POLines
        If oPOL.Fulfilled <> "CAN" Then
            If oPOL.QtyReceivedSoFar > 0 Then
                bCanCancel = False
                Exit For
            End If
        End If
    Next
    CanCancel = bCanCancel
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.CanCancel"
End Property

Public Function VoidDocument()
    On Error GoTo errHandler
Dim oPO As a_PO
    Set oPO = New a_PO
    oPO.Load Me.TRID, False
    oPO.BeginEdit
    oPO.SetStatus stVOID
    oPO.ApplyEdit
    Set oPO = Nothing
    Me.Load Me.TRID, True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.VoidDocument"
End Function

Public Property Get ISForeignCurrency() As Boolean
    On Error GoTo errHandler
    ISForeignCurrency = (Not oPC.Configuration.DefaultCurrency Is Me.CaptureCurrency)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.ISForeignCurrency"
End Property

Public Property Get CurrencyConversionAsText() As String
    On Error GoTo errHandler
    CurrencyConversionAsText = mudtProps.CurrencyRate & " " & Me.CaptureCurrency.Description & " per " & oPC.Configuration.DefaultCurrency.Description
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO.CurrencyConversionAsText"
End Property

Public Property Get Log() As String
    On Error GoTo errHandler
    Log = Trim(mudtProps.Log)
    Exit Property
errHandler:
    ErrorIn "a_PO.Log"
End Property

Public Property Let Log(val As String)
    On Error GoTo errHandler
    mudtProps.Log = Trim(val)
    Exit Property
errHandler:
    ErrorIn "a_PO.Log(val)", val
End Property
Public Sub setDelToStoreIDImmediate(val As Long)
    On Error GoTo errHandler
Dim oSM As New z_StockManager

    mudtProps.DELTOStoreID = val
    mudtProps.DelToAddress = oPC.Configuration.Stores.FindStoreByID(val).DelAddress
    'DUE TO A historic mixup the following line uses BILL rather than DEL
    oSM.SetAddressImmediate "BILL", Me.TRID, mudtProps.DELTOStoreID
    oSM.SetAddressImmediate "DEL", Me.TRID, mudtProps.DELTOStoreID
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.setDelToStoreIDImmediate(val)", val
End Sub

