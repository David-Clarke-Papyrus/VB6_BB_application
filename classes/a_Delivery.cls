VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Delivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lngLPLessDiscExt_tot As Long
Dim lngLPShortLessDiscExt_tot As Long
Dim lngFPLessDiscExt_tot As Long
Dim lngFPShortLessDiscExt_tot As Long
Dim lngLPExt_tot As Long
Dim lngFPExt_tot As Long
Dim lngVAT_tot As Long
Dim lngDisc_tot As Long
Dim lngTotalQty As Long
Dim lngTotalQtyShort As Long
Dim lngvalueShort As Long
Private xMLDoc As ujXML

Event TotalChange(lngTotal As String, lngTotalForeign As String, strQtyTotal As String)
Dim oC As New c_COLsperDEL

Event Valid(pMsg As String)
Event ValidToSave(pOK As Boolean)
Private cPolsOSPerSupp As c_POLsOS
Private mudtProps As DELProps
Private mcolStack As Collection
Private colClassors As Collection
Private WithEvents mSupplier As a_Supplier
Attribute mSupplier.VB_VarHelpID = -1
Private mColDeliveryLines As ch_DeliveryLine
Dim oPCL As z_PrintClient
Private moCaptureInCurrency As a_Currency
Attribute moCaptureInCurrency.VB_VarHelpID = -1
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Event Dirty(pVal As Boolean)
Const CDOCCODE = "DE"
Public Property Get constDOCCODE() As String
    constDOCCODE = CDOCCODE
End Property
Public Property Get ProcessingDate() As Date
    On Error GoTo errHandler
    ProcessingDate = mudtProps.ProcessingDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ProcessingDate"
End Property
Public Property Get ProcessingDateF() As String
    On Error GoTo errHandler
    ProcessingDateF = Format(mudtProps.ProcessingDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ProcessingDateF"
End Property

Public Function CalculateTotals()
    On Error GoTo errHandler
    CalculateTotal_Local
    CalculateTotal_Foreign
    RaiseEvent TotalChange(Me.TotalLessDiscExtF(False), Me.TotalLessDiscExtF(True), Me.TotalQuantityNetF)
    ValidateObject "TOTAL"
   ' GetSTatus
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.CalculateTotals"
End Function
Public Function SetMemo(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Memo) Then
    
        Err.Raise 384
    End If
    If val <> mudtProps.Memo Then
        mudtProps.Memo = val
        SetDirty True
    End If
    SetMemo = True
    Exit Function
End Function
Public Property Get Memo() As String
    On Error GoTo errHandler
    Memo = FNS(mudtProps.Memo)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.Memo"
End Property

Public Function CalculateTotal_Local()
    On Error GoTo errHandler
Dim ol As a_DeliveryLine
Dim bLineDiscountApplied As Boolean
    bLineDiscountApplied = False
    lngLPLessDiscExt_tot = 0
    lngLPShortLessDiscExt_tot = 0
    lngLPExt_tot = 0
    lngVAT_tot = 0
    lngDisc_tot = 0
    lngTotalQty = 0
    lngTotalQtyShort = 0
    For Each ol In mColDeliveryLines
       ' If oL.Fulfilled <> "CAN" Then
        If Not ol.IsDeleted Then
            ol.RecalculateLine
            lngLPExt_tot = lngLPExt_tot + ol.PExt(False)
            lngVAT_tot = lngVAT_tot + ol.PLessDiscExtVAT(False)
            lngDisc_tot = lngDisc_tot + ol.PDiscExt(False)
            lngLPLessDiscExt_tot = lngLPLessDiscExt_tot + ol.PLessDiscExt(False)
            lngLPShortLessDiscExt_tot = lngLPShortLessDiscExt_tot + ol.PShortLessDiscExt(False)
            lngTotalQty = lngTotalQty + ol.QtyFirm + ol.QtySS
            lngTotalQtyShort = lngTotalQtyShort + ol.QtyShort
        End If
      '  End If
    Next
    mudtProps.TotalQtyItems = lngTotalQty
    mudtProps.TotalQtyShort = lngTotalQtyShort
    mudtProps.TotalExtensionSimple = lngLPExt_tot
    mudtProps.TotalDiscount = lngDisc_tot
    mudtProps.TotalVAT = lngVAT_tot
    mudtProps.TotalExtension = lngLPLessDiscExt_tot
    mudtProps.TotalExtensionShort = lngLPShortLessDiscExt_tot
    
    mudtProps.TotalPayable = lngLPLessDiscExt_tot ' - lngLDeposit_tot
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'overflow
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.CalculateTotal_Local"
End Function
Public Function CalculateTotal_Foreign()
    On Error GoTo errHandler
Dim ol As a_DeliveryLine
Dim bLineDiscountApplied As Boolean
    mobjValid.RuleBroken "OVERFLOW", False
    bLineDiscountApplied = False
    lngFPLessDiscExt_tot = 0
    lngFPShortLessDiscExt_tot = 0
    lngFPExt_tot = 0
    lngVAT_tot = 0
    lngDisc_tot = 0
    For Each ol In mColDeliveryLines
      '  If oL.Fulfilled <> "CAN" Then
            lngFPExt_tot = lngFPExt_tot + ol.PExt(True)
            lngVAT_tot = lngVAT_tot + ol.PLessDiscExtVAT(True)
            lngDisc_tot = lngDisc_tot + ol.PDiscExt(True)
            lngFPLessDiscExt_tot = lngFPLessDiscExt_tot + ol.PLessDiscExt(True)
            lngFPShortLessDiscExt_tot = lngFPShortLessDiscExt_tot + ol.PShortLessDiscExt(True)
      '  End If
    Next
    mudtProps.TotalExtensionSimple_Foreign = lngFPExt_tot
    mudtProps.TotalDiscount_Foreign = lngDisc_tot
    mudtProps.TotalVAT_Foreign = lngVAT_tot
    mudtProps.TotalExtension_Foreign = lngFPLessDiscExt_tot
    mudtProps.TotalExtensionShort_Foreign = lngFPShortLessDiscExt_tot
    mudtProps.TotalPayable_Foreign = lngFPLessDiscExt_tot ' - lngFDeposit_tot
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'ov erflow
        mobjValid.RuleBroken "OVERFLOW", True
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.CalculateTotal_Foreign"
End Function
'Public Property Get TotalInvoiceDiscountAmountF() As Long
'    TotalInvoiceDiscountAmountF = Format(lngTotalInvoiceDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
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
    ErrorIn "a_Delivery.TotalLessDiscExtF(bForeign)", bForeign
End Property
Public Property Get TotalShortLessDiscExtF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalShortLessDiscExtF = Format(mudtProps.TotalExtensionShort_Foreign / CaptureCurrency.Divisor, CaptureCurrency.FormatString)
    Else
        TotalShortLessDiscExtF = Format(mudtProps.TotalExtensionShort / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.TotalShortLessDiscExtF(bForeign)", bForeign
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
    ErrorIn "a_Delivery.TotalExtF(bForeign)", bForeign
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
    ErrorIn "a_Delivery.TotalLessDiscExt(bForeign)", bForeign
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
    ErrorIn "a_Delivery.TotalVATF(bForeign)", bForeign
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
    ErrorIn "a_Delivery.TotalPayable(bForeign)", bForeign
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
    ErrorIn "a_Delivery.TotalPayableF(bForeign)", bForeign
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
    ErrorIn "a_Delivery.TotalDiscount(bForeign)", bForeign
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
    ErrorIn "a_Delivery.TotalDiscountF(bForeign)", bForeign
End Property




Public Sub BeginEdit()
    On Error GoTo errHandler
  mColDeliveryLines.BeginEdit
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  mColDeliveryLines.CancelEdit
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.CancelEdit"
End Sub
Public Sub ApplyEdit()
    On Error GoTo errHandler
Dim objPersist As a_Delivery_P
 '   CalculateTotal
    mudtProps.TPID = Me.Supplier.ID
        
    If DOCCode = "" Then
        DOCCode = GetRef("GETANDWRITE")
        DOCDate = Now()
    End If
    If IssDate < #1/1/1995# And mudtProps.IsNew Then IssDate = Now
    mudtProps.COMPID = oPC.Configuration.DefaultCOMPID
    If mcolStack.Count = 0 Then Err.Raise 445
    mColDeliveryLines.ApplyEdit
    Set objPersist = New a_Delivery_P
    If mudtProps.IsDeleted Then
        objPersist.DeleteObject mudtProps.TRID
        mcolStack.Remove mcolStack.Count
        mudtProps.IsNew = True
        mudtProps.IsDeleted = False
    ElseIf IsDirty Or mudtProps.IsNew Then
        If mudtProps.IsNew Then mudtProps.CaptureDate = Now()
        If Me.StatusF = "ISSUED" Then
            If Not IsValid Then Err.Raise 445
        End If
        mcolStack.Remove mcolStack.Count
        SetSuperState objPersist.Save(GetSuperState)
        mudtProps.IsNew = False
    Else
        mcolStack.Remove mcolStack.Count
    End If
    Set objPersist = Nothing
    mudtProps.IsDirty = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.ApplyEdit"
End Sub

Public Function GetRef(strTmp As String) As String
    On Error GoTo errHandler
Dim oCode As z_Code
Dim strDefaultRef As String
Dim strRef As String
    Set oCode = New z_Code
    
    strRef = oCode.GetNextCode(enDelivery, strTmp)
    
    strDefaultRef = oPC.Configuration.DefaultCompany.CompanyCode
    GetRef = strDefaultRef & "D" & strRef
    
    Set oCode = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.GetRef(strTmp)", strTmp
End Function
Public Property Get IssDate() As Date
    On Error GoTo errHandler
    IssDate = mudtProps.CaptureDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.IssDate"
End Property
Public Property Get IssDateF() As String
    On Error GoTo errHandler
    IssDateF = Format(mudtProps.CaptureDate, "dd/mm/yyyy hh:nn AM/PM")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.IssDateF"
End Property
Public Property Let IssDate(val As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CaptureDate = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.IssDate(val)", val
End Property

Public Property Get Supplier() As a_Supplier
    Set Supplier = mSupplier
End Property
Public Function SetSupplier(pTPID As Long) As Boolean
    On Error GoTo errHandler
Dim bSuccess As Boolean
    bSuccess = Supplier.Load(pTPID)
    SetSupplier = bSuccess
    If bSuccess Then
        mudtProps.TPID = pTPID
        VATable = Supplier.VATable
        If oPC.SupplierBasedCurrencyConversion Then
            SetDELCurrency Supplier.DefaultCurrency, Supplier.ConversionToLocalFactor
        Else
            If Me.CurrencyID > 0 Then
                SetDELCurrency oPC.Configuration.Currencies.FindCurrencyByID(Me.CurrencyID), 0
            Else
                SetDELCurrency Supplier.DefaultCurrency, 0
            End If
        End If
        mobjValid.RuleBroken "TP", False
        SetDirty True
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.SetSupplier(pTPID)", pTPID
End Function
Public Function ReloadMatches(pPID As String)
    On Error GoTo errHandler
        Set cPolsOSPerSupp = Nothing
        Set cPolsOSPerSupp = New c_POLsOS
        cPolsOSPerSupp.Load , , , pPID
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.ReloadMatches(pPID)", pPID
End Function
Public Function ReloadMatches_forSubstitutions(pPID As String)
    On Error GoTo errHandler
        Set cPolsOSPerSupp = Nothing
        Set cPolsOSPerSupp = New c_POLsOS
        cPolsOSPerSupp.LoadSubstitutions , , , pPID
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.ReloadMatches(pPID)", pPID
End Function
Public Function SetDELCurrency(val As a_Currency, Optional ConversionRate As Double) As Boolean
    On Error GoTo errHandler
    If val Is Nothing Then
        Set val = oPC.Configuration.DefaultCurrency
'        SetDELCurrency = False
'        Exit Function
    End If
    Set moCaptureInCurrency = val
    mudtProps.CurrencyID = val.ID
    If oPC.SupplierBasedCurrencyConversion Then
        If mudtProps.CurrencyRate = 0 Then mudtProps.CurrencyRate = ConversionRate
    Else
        If mudtProps.CurrencyRate = 0 Then mudtProps.CurrencyRate = val.Factor
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.SetDELCurrency(val)", val
End Function

Public Property Get VATable() As Boolean
    On Error GoTo errHandler
  '  Vatable = mudtProps.Vatable
  VATable = mudtProps.VATable
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.VATable"
End Property
Public Property Let VATable(val As Boolean)
    On Error GoTo errHandler
    mudtProps.VATable = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.VATable(val)", val
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = mcolStack.Count > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.IsEditing"
End Property

Private Sub Class_Initialize()
    On Error GoTo errHandler
    LoadClassorsCollection
    Set mobjValid = New z_BrokenRules
    mudtProps.IsNew = True
    Set mColDeliveryLines = New ch_DeliveryLine
    Set mcolStack = New Collection
    mColDeliveryLines.component Me
    Set mSupplier = New a_Supplier
    mudtProps.IsNew = True
    mudtProps.VATRate = oPC.Configuration.VATRate
    mudtProps.VATable = True
    mudtProps.BatchTotalExtras = 0
    mudtProps.BatchTotal = 0
    mudtProps.BatchQtyTotal = 0
    mobjValid.BreakRule "TP", True
  '  Set oCurrency_Foreign = oPC.Configuration.DefaultCurrency
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set mobjValid = Nothing
    Set mColDeliveryLines = Nothing
    Set mcolStack = Nothing
    Set mSupplier = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.Class_Terminate"
End Sub

Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As COrderData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.GetState"
End Function
Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As COrderData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.SetState(buffer)", buffer
End Sub
Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "DELLS", mColDeliveryLines.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.GetSuperState"
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
    mColDeliveryLines.SetSuperState .ReadProperty("DELLS")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.SetSuperState(buffer)", buffer
End Sub

Public Sub Load(DeliveryID As Long)
    On Error GoTo errHandler
Dim oDeliveryPersist As a_Delivery_P

    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    Set oDeliveryPersist = New a_Delivery_P
    SetSuperState oDeliveryPersist.Fetch(DeliveryID)
    Set oDeliveryPersist = Nothing
    If TPID > 0 Then
      SetSupplier Me.TPID
    End If
    mudtProps.IsNew = False
    CalculateTotals
    mudtProps.IsDirty = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.Load(DeliveryID)", DeliveryID
End Sub

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.Delete"
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.IsDeleted"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty Or mColDeliveryLines.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.IsDirty"
End Property

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.IsValid"
End Property
    
    Public Property Get TPACCNum() As String
    On Error GoTo errHandler
        TPACCNum = FNS(mudtProps.TPACCNum)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.TPACCNum"
    End Property
    
    Public Property Let TPACCNum(Value As String)
    On Error GoTo errHandler
        mudtProps.TPACCNum = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.TPACCNum(Value)", Value
    End Property
    
Public Property Get TPNAME() As String
    On Error GoTo errHandler
    TPNAME = FNS(mudtProps.TPNAME)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.TPName"
End Property

Public Property Let TPNAME(Value As String)
    On Error GoTo errHandler
    mudtProps.TPNAME = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.TPName(Value)", Value
End Property
Public Property Get SupplierInvoiceRef() As String
    On Error GoTo errHandler
    SupplierInvoiceRef = FNS(mudtProps.SupplierInvoiceRef)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.SupplierInvoiceRef"
End Property

Public Property Let SupplierInvoiceRef(Value As String)
    On Error GoTo errHandler
    mudtProps.SupplierInvoiceRef = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.SupplierInvoiceRef(Value)", Value
End Property
Public Property Get SupplierInvoiceDate() As Date
    On Error GoTo errHandler
    SupplierInvoiceDate = mudtProps.SupplierInvoiceDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.SupplierInvoiceDate"
End Property
Public Property Get SupplierInvoiceDateF() As String
    On Error GoTo errHandler
    If CLng(mudtProps.SupplierInvoiceDate) = 0 Then
        SupplierInvoiceDateF = ""
    Else
        SupplierInvoiceDateF = Format(mudtProps.SupplierInvoiceDate, "dd/mm/yyyy")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.SupplierInvoiceDateF"
End Property

Public Property Let SupplierInvoiceDate(Value As Date)
    On Error GoTo errHandler
    mudtProps.SupplierInvoiceDate = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.SupplierInvoiceDate(Value)", Value
End Property
    
Public Function SetSupplierInvoiceDate(val As String) As Boolean
Dim lngMonths As Long

    If Not IsDate(val) Then
        SetSupplierInvoiceDate = False
        Exit Function
    Else
        lngMonths = DateDiff("m", Date, CDate(val))
        If lngMonths > 0 Then
            RaiseEvent Valid("Date is in the future - cannot accept.")
            SetSupplierInvoiceDate = False
            Exit Function
        ElseIf lngMonths < -2 Then
            Select Case lngMonths
            Case Is > -13
                RaiseEvent Valid("The invoice date is more than two months ago. Is it correct?")
                SetSupplierInvoiceDate = False
                Exit Function
            Case Else
                RaiseEvent Valid("The invoice date is too old. You must correct it.")
                SetSupplierInvoiceDate = True
                Exit Function
            End Select
        End If
    End If
    SetSupplierInvoiceDate = True
    SupplierInvoiceDate = CDate(val)

End Function


    
Public Property Get DOCCode() As String
    On Error GoTo errHandler
    DOCCode = FNS(mudtProps.DOCCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.DocCode"
End Property

Public Property Let DOCCode(Value As String)
    On Error GoTo errHandler
    mudtProps.DOCCode = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.DocCode(Value)", Value
End Property

Public Property Get DOCDate() As Date
    On Error GoTo errHandler
    DOCDate = mudtProps.DOCDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.DocDate"
End Property
Public Property Get DocDateF() As Date
    On Error GoTo errHandler
    DocDateF = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.DocDateF"
End Property

Public Property Let DOCDate(Value As Date)
    On Error GoTo errHandler
    mudtProps.DOCDate = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.DocDate(Value)", Value
End Property
'
'
'    Public Property Get POLID() As Long
'        POLID = mudtProps.POLID
'    End Property
'
'    Public Property Let POLID(Value As Long)
'        mudtProps.POLID = Value
'    End Property
    
    Public Property Get TRID() As Long
    On Error GoTo errHandler
        TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.TRID"
    End Property
    
    Public Property Let TRID(Value As Long)
    On Error GoTo errHandler
        mudtProps.TRID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.TRID(Value)", Value
    End Property
Public Property Get CurrencyID() As Long
    On Error GoTo errHandler
    CurrencyID = FNN(mudtProps.CurrencyID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.CurrencyID"
End Property
Public Property Get CaptureCurrency() As a_Currency
    Set CaptureCurrency = moCaptureInCurrency
End Property
Public Property Let CurrencyID(Value As Long)
    On Error GoTo errHandler
    mudtProps.CurrencyID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.CurrencyID(Value)", Value
End Property
Public Property Get ISForeignCurrency() As Boolean
    On Error GoTo errHandler
    ISForeignCurrency = Not (mudtProps.CurrencyID = oPC.Configuration.DefaultCurrencyID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.ISForeignCurrency"
End Property
Public Property Get TPID() As Long
    On Error GoTo errHandler
    TPID = mudtProps.TPID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.TPID"
End Property

Public Property Let TPID(Value As Long)
    On Error GoTo errHandler
    mudtProps.TPID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.TPID(Value)", Value
End Property
Public Property Get StaffID() As Long
    On Error GoTo errHandler
    StaffID = mudtProps.StaffID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.StaffID"
End Property
Public Property Get StaffName() As String
    On Error GoTo errHandler
    If mudtProps.StaffID < 1 Then
        StaffName = ""
        Exit Property
    End If
    StaffName = oPC.Configuration.Staff.FindStaffByID(StaffID).Shortname
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.StaffName"
End Property
Public Property Get StaffNameB() As String
    On Error GoTo errHandler
Dim strT As String
    If StaffID = 0 Then
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
    ErrorIn "a_Delivery.StaffNameB"
End Property
Public Property Let StaffID(Value As Long)
    On Error GoTo errHandler
    mudtProps.StaffID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.StaffID(Value)", Value
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
    ErrorIn "a_Delivery.StatusF"
End Property
Public Property Get GetRepricingRequirement() As ADODB.Recordset
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "Select * FROM vLabels_D where TRID = " & Me.TRID & " AND RepriceCount > 0", oPC.COShort, adOpenKeyset, adLockOptimistic
    Set GetRepricingRequirement = rs

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.GetRepricingRequirement"
End Property
Public Property Get Status() As enStatus
    On Error GoTo errHandler
    Status = mudtProps.Status
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.Status"
End Property
    
Public Property Get BatchTotalExtras() As Long
    On Error GoTo errHandler
    BatchTotalExtras = mudtProps.BatchTotalExtras
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.BatchTotalExtras"
End Property
Public Property Get BatchTotalExtrasF() As String
    On Error GoTo errHandler
    BatchTotalExtrasF = Format(mudtProps.BatchTotalExtras / Me.CaptureCurrency.Divisor, CaptureCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.BatchTotalExtrasF"
End Property
Public Function SetBatchTotalExtras(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long

        lngTemp = mudtProps.BatchTotalExtras
        bOK = SetField_LONG(mudtProps.BatchTotalExtras, val, "BATCHTOTALEXTRAS", mcolStack.Count)
        bOK = ValidateObject("BATCHTOTALEXTRAS")
        SetBatchTotalExtras = bOK

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.SetBatchTotalExtras(val)", val
End Function

Public Property Get BatchTotal() As Long
    On Error GoTo errHandler
    BatchTotal = mudtProps.BatchTotal
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.BatchTotal"
End Property
Public Property Get BatchTotalF() As String
    On Error GoTo errHandler
    BatchTotalF = Format(mudtProps.BatchTotal / Me.CaptureCurrency.Divisor, CaptureCurrency.FormatString) & " " & Me.CaptureCurrency.SYSNAME
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.BatchTotalF"
End Property
Public Function SetBatchTotal(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long

        lngTemp = mudtProps.BatchTotal
        bOK = SetField_LONG(mudtProps.BatchTotal, val, "BATCHTOTAL", mcolStack.Count)
        bOK = ValidateObject("BATCHTOTAL")
        SetBatchTotal = bOK

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.SetBatchTotal(val)", val
End Function
Public Function SetBatchQtyTotal(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long

        lngTemp = mudtProps.BatchQtyTotal
        bOK = SetField_LONG(mudtProps.BatchQtyTotal, val, "BATCHQTYTOTAL", mcolStack.Count)
        bOK = ValidateObject("BATCHQTYTOTAL")
        SetBatchQtyTotal = bOK

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.SetBatchQtyTotal(val)", val
End Function
Public Property Get BatchQtyTotal() As Long
    On Error GoTo errHandler
    BatchQtyTotal = mudtProps.BatchQtyTotal
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.BatchQtyTotal"
End Property
Public Property Get BatchQtyTotalF() As String
    On Error GoTo errHandler
    BatchQtyTotalF = mudtProps.BatchQtyTotal
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.BatchQtyTotalF"
End Property

Public Property Get DeliveryLines() As ch_DeliveryLine
  Set DeliveryLines = mColDeliveryLines
End Property
Public Property Get POLsOSPersSUPP() As c_POLsOS
  Set POLsOSPersSUPP = cPolsOSPerSupp
End Property
Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    mudtProps.IsDirty = pVal
    RaiseEvent Dirty(pVal)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.SetDirty(pVal)", pVal
End Sub
Public Function GetStatus()
    On Error GoTo errHandler
    If IsValid Then
        RaiseEvent Valid("")
    Else
        RaiseEvent Valid(TranslateErrors(mobjValid.AllBrokenRules))
        RaiseEvent ValidToSave((InStr(1, mobjValid.AllBrokenRules, ",TP") = 0) And (InStr(1, mobjValid.AllBrokenRules, ",NOTE") = 0))
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.GetSTatus"
End Function
Private Function TranslateErrors(pRawors As String) As String
    On Error GoTo errHandler
Dim strOut As String
Dim strRule, strAllRules As String
Dim NoMoreRules As Boolean
Dim iMarker, iStart As Integer
    iMarker = 1
    strAllRules = ""
    If Len(pRawors) > 0 Then
        iMarker = InStr(iMarker + 1, pRawors, ",")
        If iMarker > 0 Then
            strAllRules = colClassors(Left(pRawors, iMarker - 1))
        Else
            strAllRules = colClassors(pRawors)
        End If
        Do Until iMarker = 0
            iStart = iMarker + 1
            iMarker = InStr(iStart, pRawors, ",")
            If iMarker > 0 Then
                strRule = colClassors(Mid(pRawors, iStart, iMarker - iStart))
            Else
                strRule = colClassors(Mid(pRawors, iStart))
            End If
                
            strAllRules = strAllRules & vbCrLf & strRule
        Loop
    End If
    TranslateErrors = strAllRules
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Missing supplier", "TP"
    colClassors.Add "Note too short", "NOTE"
    colClassors.Add "Batch control value mismatch", "BATCHTOTAL"
    colClassors.Add "Batch control quantity mismatch", "BATCHQTYTOTAL"
    colClassors.Add "Extras Out of balance", "BATCHTOTALEXTRAS"
    colClassors.Add "A numeric value is too great for calculation", "OVERFLOW"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.LoadClassorsCollection"
End Sub
Friend Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean

    bValid = True
    If mudtProps.TPID = 0 Then
        mobjValid.BreakRule "TP", True
        If pFld = "TP" Then bValid = False
    Else
        mobjValid.BreakRule "TP", False
    End If
'    If (mudtProps.BatchTotal > Me.TotalLessDiscExt(Me.CurrencyID <> oPC.Configuration.DefaultCurrencyID) * 1.005) _
'        Or (mudtProps.BatchTotal < Me.TotalLessDiscExt(Me.CurrencyID <> oPC.Configuration.DefaultCurrencyID) * 0.995) Then
    If (mudtProps.BatchTotal - mudtProps.BatchTotalExtras > (Me.TotalLessDiscExt(Me.CurrencyID <> oPC.Configuration.DefaultCurrencyID) * (1 + oPC.SupplierInvoiceTolerance))) _
        Or (mudtProps.BatchTotal - mudtProps.BatchTotalExtras < (Me.TotalLessDiscExt(Me.CurrencyID <> oPC.Configuration.DefaultCurrencyID) * (1 - oPC.SupplierInvoiceTolerance))) Then
        mobjValid.BreakRule "BATCHTOTAL", True
    Else
        mobjValid.BreakRule "BATCHTOTAL", False
    End If
    If mudtProps.BatchQtyTotal <> (mudtProps.TotalQtyItems + mudtProps.TotalQtyShort) Then
        mobjValid.BreakRule "BATCHQTYTOTAL", True
    Else
        mobjValid.BreakRule "BATCHQTYTOTAL", False
    End If

    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.ValidateObject(pFld)", pFld
End Function
Public Property Get TotalQuantityNet() As Long
    TotalQuantityNet = (mudtProps.TotalQtyItems + mudtProps.TotalQtyShort)
End Property
Public Property Get TotalQuantityNetF() As String
    TotalQuantityNetF = Format(TotalQuantityNet, "###,##0")
End Property

Public Sub SetStatus(val As enStatus)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Status = val
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.SetStatus(val)", val
End Sub
Public Function Post() As String
    On Error GoTo errHandler
Dim strResult As String
Dim oSM As z_StockManager
Dim dCOLine As d_COLine
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
 '   oPC.COShort.BeginTrans
    Me.ApplyEdit
    Me.CalculateTotals
    Set oSM = New z_StockManager
    
    Post = ""
    Select Case Status
    Case stVOID
    Case stInProcess
        Post = "In Process"
    Case stCOMPLETE
        oSM.IssueDEL Me
    Case stISSUED
        oSM.IssueDEL Me
    Case stPROFORMA
        oSM.IssueDEL Me
    End Select
    
'    'put certain books into reserve
'    Set oC = Nothing
'    Set oC = New c_COLsperDEL
'    oC.Load Me.TRID
'    For Each dCOLine In oC
'        oPC.COShort.execute "UPDATE tPRoduct SET P_QtyReserved = dbo.MaxNumber(0,P_QtyReserved) + " & dCOLine.Qty & " WHERE P_ID = '" & dCOLine.PID & "'"
'    Next
'    Set oC = Nothing
    Set oSM = Nothing
 '   oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.Post", , EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Function
Public Function Print_Display(bForeign As Boolean) As String
    On Error GoTo errHandler
Dim str As String
Dim i As Integer

    
    str = "Delivery: " & Me.DOCCode & vbCrLf
    
    For i = 1 To DeliveryLines.Count
        With Me.DeliveryLines(i)
            str = str & .CodeF & "  " & .Ref & "  " & .QtyFirmF & "  " & .QtySSF & "  " & Left(.Title, 30) & "  " & .PriceF(bForeign) & "  " & IIf(.Discount = 0, "", .DiscountF) & "  " & .PLessDiscExtF(bForeign)
        End With
        str = str & vbCrLf
    Next i
    
    
    str = str & vbCrLf
   
    Print_Display = str
    
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
    ErrorIn "a_Delivery.Print_Display(bForeign)", bForeign
End Function


Public Function VoidDocument()
    On Error GoTo errHandler
Dim strResult As String
    Me.BeginEdit
    Me.SetStatus stVOID
    Me.ApplyEdit
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.VoidDocument"
End Function

Public Sub Reload()
    On Error GoTo errHandler
Dim lngTRID As Long
    lngTRID = Me.TRID
    Class_Terminate
    Class_Initialize
    Load lngTRID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.Reload"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent ValidToSave(oRS = "BATCHTOTAL")
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent ValidToSave(oRS = "BATCHTOTAL")
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(errors As String)
    On Error GoTo errHandler
    RaiseEvent ValidToSave(errors = "BATCHTOTAL")
    RaiseEvent Valid(TranslateErrors(errors))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.mobjValid_Status(Errors)", errors
End Sub

Public Property Get CurrencyConversionAsText() As String
    On Error GoTo errHandler
    CurrencyConversionAsText = Format(FNDBL(mudtProps.CurrencyRate), "##0.000") & " " & Me.CaptureCurrency.Description & " per " & oPC.Configuration.DefaultCurrency.Description
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.CurrencyConversionAsText"
End Property
Public Property Get CurrencyConversionInverseAsText() As String
    On Error GoTo errHandler
    CurrencyConversionInverseAsText = Format(FNDBL(1 / mudtProps.CurrencyRate), "##0.000") & " " & oPC.Configuration.DefaultCurrency.Description & " per " & Me.CaptureCurrency.Description
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.CurrencyConversionInverseAsText"
End Property
Public Property Get CurrencyConversionInverseRate() As String
    On Error GoTo errHandler
    CurrencyConversionInverseRate = Format(FNDBL(1 / IIf(mudtProps.CurrencyRate = 0, 1, mudtProps.CurrencyRate)), "##0.000")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.CurrencyConversionInverseRate"
End Property

Public Function HasSupplierClaim() As Boolean
Dim oDELL As a_DeliveryLine


'    lngTotalQtyShort = 0
'    lngvalueShort = 0
'    For Each oDELL In Me.DeliveryLines
'        lngTotalQtyShort = lngTotalQtyShort + oDELL.QtyShort
'        lngvalueShort = lngvalueShort + oDELL.PShortLessDiscExt(Me.ISForeignCurrency)
'    Next
    HasSupplierClaim = (Me.QtyShort > 0)
    
End Function
Public Property Get QtyShort() As Long
    QtyShort = lngTotalQtyShort
End Property
Public Property Get ValueShort() As Long
    ValueShort = lngvalueShort
End Property
Public Property Get ValueShortF() As String
    ValueShortF = Format(CDbl(lngvalueShort) / oPC.Configuration.Currencies.FindCurrencyByID(Me.CurrencyID).Divisor, oPC.Configuration.Currencies.FindCurrencyByID(Me.CurrencyID).FormatString)
End Property

'Public Sub PrintSupplierClaim()
'Dim rpt As New arSupplierClaim
'
'    Me.CalculateTotals
'    rpt.Component Me
'    rpt.Show vbModal
'
'End Sub


Public Function ExportToXML(DispatchMode As enTransmitType, _
                        Optional pDestinationEmailAddress As String, _
                        Optional pWholeMessage As String, Optional pQtyCopies As Integer, Optional QuickPrint As Boolean) As Boolean
    On Error GoTo errHandler
Dim bForeign As Boolean
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
Dim bReferenceExists As Boolean
Dim bDiscountExists As Boolean
Dim bDepositExists As Boolean
Dim bQtyColumnNeeded As Boolean
Dim oLine As a_DeliveryLine
Dim oDC As a_DocumentControl
Dim strWorkingFolder As String
Dim strDispatchMethod As String
Dim strTemplateName As String
Dim strExecutable As String
Dim bMultipleCOLReps As Boolean
Dim strTmpRep As String
Dim flgFirst As Boolean
Dim cnt As Integer

    If DispatchMode = enMail Then
        If oPC.UsesOutlookForINVEmail Then
            strWorkingFolder = oPC.LocalFolder & "TEMP\"
            strDispatchMethod = "Outlook"
        Else
            strWorkingFolder = oPC.SharedFolderRoot & "\Emails\"
            strDispatchMethod = "Emailing"
        End If
    ElseIf DispatchMode = enPrint Then
        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
        strDispatchMethod = "Printing"
    ElseIf DispatchMode = enView Then
        strWorkingFolder = oPC.SharedFolderRoot & "\PDF\"
        strDispatchMethod = ""
    Else
        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
        strDispatchMethod = "Printing"
    End If

                        p 1
    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing And strDispatchMethod <> "Emailing" And strDispatchMethod <> "Outlook" Then
        ExportToXML = False
        Exit Function
    End If
                        p 4
    If pQtyCopies < 1 Then pQtyCopies = 1

    Set xMLDoc = New ujXML

    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "GRN_DOC"
            .chCreate "MessageType"
                .elText = "GRN"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "TemplateName"
                    .elText = "GRN_DOC"
            strTemplateName = .elText
                        p 5
            .elCreateSibling "CopyCount"
                .elText = pQtyCopies
            .elCreateSibling "Printer"
                If Not oDC Is Nothing Then .elText = oDC.PrinterName
            .elCreateSibling "Status"
                .elText = Me.StatusF
            .elCreateSibling "LogoPath"
                If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
                Else
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
                End If
            .elCreateSibling "Sender", True
                .elText = oPC.Configuration.DefaultCompany.CompanyName
            .elCreateSibling "SenderAddress", True
                .elText = Replace(oPC.Configuration.DefaultCompany.StreetAddress, Chr(13) & Chr(10), Chr(10))
           
            .elCreateSibling "DocCode"
                .elText = Me.DOCCode
            .elCreateSibling "DocDate", True
                .elText = DocDateF
                
            .elCreateSibling "SInvRef"
                .elText = SupplierInvoiceRef
            .elCreateSibling "SInvDate", True
                .elText = SupplierInvoiceDateF
                
            .elCreateSibling "StoreCode", True
                .elText = oPC.Configuration.DefaultStore.code
            
                                p 6
            .elCreateSibling "SupplierName", True
                .elText = Supplier.Name
                
            .elCreateSibling "SupplierPhone", True
                If Supplier.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = IIf(Me.Supplier.BillTOAddress.Phone > "", "Phone: " & Supplier.BillTOAddress.Phone, "")
                End If
            .elCreateSibling "SupplierFax", True
                If Supplier.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = IIf(Supplier.BillTOAddress.Fax > "", "Fax: " & Supplier.BillTOAddress.Fax, "")
                End If
            
            .elCreateSibling "ACNO"
                .elText = IIf(Me.Supplier.AcNo > "", "Ac/no. " & Me.Supplier.AcNo, "")
                            p 8
                               p 9
            For i = 1 To DeliveryLines.Count
                    .elCreateSibling "DetailLine", True
                    .chCreate "CodeF"
                        .elText = DeliveryLines(i).CodeF
                    .elCreateSibling "Title", True
                        .elText = DeliveryLines(i).Title
                    .elCreateSibling "QtyFirm", True
                        .elText = DeliveryLines(i).QtyFirmF
                    .elCreateSibling "QtySS", True
                        .elText = DeliveryLines(i).QtySSF
                    .elCreateSibling "QtyShort", True
                        .elText = DeliveryLines(i).QtyShortF
                    .elCreateSibling "PriceF", True
                        .elText = DeliveryLines(i).PriceF(ISForeignCurrency)
                    .elCreateSibling "SPF", True
                        .elText = DeliveryLines(i).PriceSellF
                    .elCreateSibling "MBCode", True
                        .elText = oPC.Configuration.Multibuys.ItemByF4(DeliveryLines(i).MBCode)
                    .elCreateSibling "DiscountRateF", True
                        .elText = DeliveryLines(i).DiscountF
                    .elCreateSibling "Reference", True
                        .elText = DeliveryLines(i).Ref
                    .elCreateSibling "Extension", True
                        .elText = DeliveryLines(i).PLessDiscExtF(ISForeignCurrency)
                    .elCreateSibling "Note", True
                        .elText = DeliveryLines(i).Note
                    .elCreateSibling "POCode", True
                        .elText = DeliveryLines(i).POCode
                    .elCreateSibling "VATRate", True
                        .elText = DeliveryLines(i).VATRateF
                    .navUP
            Next i
                                p 10
            .elCreateSibling "TotalQtyItems", True
                .elText = Me.BatchQtyTotalF
                                p 11
            .elCreateSibling "Memo", True
                .elText = Memo
                
            .elCreateSibling "TotalValue", True
                .elText = Me.TotalPayableF(False)
                
            .elCreateSibling "TotalValueOnInvoice", True
                .elText = Me.BatchTotalF
                
            .elCreateSibling "StaffMember", True
                .elText = IIf(StaffName > "", StaffName, "")
                
    End With
                                p 12

'FINALLY PRODUCE THE .XML FILE
    strXML = strWorkingFolder & "GRN_" & Me.DOCCode & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

                                p 13

    If DispatchMode = enMail Then
        'WRITE THE .HTML FILE
            objXSL.async = False
            objXSL.ValidateOnParse = False
            objXSL.resolveExternals = False
            strPath = oPC.SharedFolderRoot & "\Templates\" & fs.GetBaseName(strTemplateName) & "_HTML.xslt"
            Set fs = New FileSystemObject
            If fs.FileExists(strPath) Then
                objXSL.Load strPath
            End If
            If fs.FileExists(strWorkingFolder & "GRN_" & Me.DOCCode & ".HTML") Then
                fs.DeleteFile strWorkingFolder & "GRN_" & Me.DOCCode & ".HTML", True
            End If
            oTF.OpenTextFileToAppend strWorkingFolder & "GRN_" & Me.DOCCode & ".HTML"
            pWholeMessage = xMLDoc.docObject.transformNode(objXSL)
            oTF.WriteToTextFile pWholeMessage
            oTF.CloseTextFile
    End If


'WRITE THE .PDF FILE IF NECESSARY
'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file
    If DispatchMode = enView Or (DispatchMode = enPrint And QuickPrint = True) Then
        Set objXSL = Nothing
        Set objXSL = New MSXML2.DOMDocument60
        objXSL.async = False
        objXSL.ValidateOnParse = False
        objXSL.resolveExternals = False
        strPath = oPC.SharedFolderRoot & "\Templates\" & strTemplateName & "_FO.XSL"
        Set fs = New FileSystemObject
        If fs.FileExists(strPath) Then
            objXSL.Load strPath
        End If
                        p 6
        If fs.FileExists(strFOFile) Then
            fs.DeleteFile strFOFile
        End If
                        
        Set opXMLDOC = New MSXML2.DOMDocument60
        opXMLDOC.async = False
        opXMLDOC.ValidateOnParse = False
        opXMLDOC.resolveExternals = False
        xMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC

        strFOFile = strWorkingFolder & "GRN_" & Me.DOCCode & ".FO"
        strPDFFile = strWorkingFolder & "GRN_" & Me.DOCCode & ".PDF"
                        p 7
        If docWriteTostream(strFOFile, opXMLDOC, "UNICODE") = False Then
              MsgBox "Cannot write to .FO file. Please contact support to check the permissions for writing " & strFOFile & " are correctly set.", vbInformation + vbOKOnly, "Can't do this"
              oTF.CloseTextFile
              Exit Function
        End If

'Stage 2 Convert the .FO file to .PDF and clean up
        If fs.FileExists(strPDFFile) Then
            On Error Resume Next
            fs.DeleteFile strPDFFile
            If Err Then
                MsgBox "It looks like you have this document open already in your PDF reader. Close the document before re-processing it.", vbInformation + vbOKOnly, "Warning"
            End If
        End If
        ChDir "\PBKS\Executables\FOP\"
        strCommand = GetFOPCommandstring(strFOFile, strPDFFile)
        
        Dim fl As New z_TextFile
        If fs.FileExists("c:\PBKS\TEMP\CREATEPDF.BAT") Then
          fs.DeleteFile ("c:\PBKS\TEMP\CREATEPDF.BAT")
        End If
        fl.OpenTextFile ("c:\PBKS\TEMP\CREATEPDF.BAT")
        fl.WriteToTextFile (strCommand)
        fl.CloseTextFile
        Shell "c:\PBKS\TEMP\CREATEPDF.BAT", vbHide
        Set fl = Nothing
        
        
        'IF IT FAILS HERE CHECK THAT JAVA IS INSTALLED
        'We do the following because the batch file executes and is not waited for by the F_7_AB_1_ShellAndWaitSimple routine
        'so the FO file is sometimes deleted before the .PDF file is produced.
    End If
    
    
    ExportToXML = True
    Dim cntr As Integer
    cntr = 0
    Do While ((fs.FileExists(strPDFFile) = False) And cntr < 15)
        MsgWaitObj 3000
        cntr = cntr + 1
    Loop
        MsgWaitObj 1000
   
    If (Not fs.FileExists(strPDFFile)) And (DispatchMode = enView Or DispatchMode = enMail) Then
        MsgBox "The PDF file has not been produced: " & strPDFFile & ". Please contact support.", vbOKOnly, "Can't do this"
    Else
        If DispatchMode = enView Or DispatchMode = enPrint Then
            OpenFileWithApplication strPDFFile, enPDF, QuickPrint
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.ExportToXML(DispatchMode,pDestinationEmailAddress,pWholeMessage,pQtyCopies)", _
         Array(DispatchMode, pDestinationEmailAddress, pWholeMessage, pQtyCopies)
End Function

Public Function ExportToSpreadsheet(bForeign As Boolean, _
        pFilename As String) As Boolean
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim s As String
Dim s2 As String
Dim lngNumberOfLines As Long

    ExportToSpreadsheet = False
    pFilename = oPC.LocalFolder & "Temp\" & DOCCode & "_" & Format(Now(), "yyyymmddHHnn") & ".xls"
    
    oTF.OpenTextFile pFilename
    oTF.WriteToTextFile "Status" & vbTab & Me.Status
    oTF.WriteToTextFile "DocCode" & vbTab & DOCCode
    oTF.WriteToTextFile "DocDate" & vbTab & DOCDate
    oTF.WriteToTextFile "SInvRef" & vbTab & SupplierInvoiceRef
    oTF.WriteToTextFile "SInvDate" & vbTab & SupplierInvoiceDateF
    oTF.WriteToTextFile "StoreCode" & vbTab & oPC.Configuration.DefaultStore.code
    oTF.WriteToTextFile "Sender" & vbTab & oPC.Configuration.DefaultCompany.CompanyName
    oTF.WriteToTextFile "SupplierName" & vbTab & Supplier.Name
    oTF.WriteToTextFile "ACNO" & vbTab & IIf(Me.Supplier.AcNo > "", "Ac/no. " & Me.Supplier.AcNo, "")
    
'    oTF.WriteToTextFile_NoLineTerminator "SupplierWithAddress" & vbTab
'    If Not Supplier.OrderToAddress Is Nothing Then
'        oTF.WriteToTextFile Replace(Supplier.OrderToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
'    End If
    
    oTF.WriteToTextFile_NoLineTerminator "SupplierPhone" & vbTab
    If Not Supplier.BillTOAddress Is Nothing Then
        oTF.WriteToTextFile IIf(Me.Supplier.BillTOAddress.Phone > "", "Phone: " & Supplier.BillTOAddress.Phone, "")
    End If
    oTF.WriteToTextFile_NoLineTerminator "SupplierFax" & vbTab
    If Not Supplier.BillTOAddress Is Nothing Then
        oTF.WriteToTextFile IIf(Supplier.BillTOAddress.Fax > "", "Fax: " & Supplier.BillTOAddress.Fax, "")
    End If


    s = "SKU" & vbTab & "Title" & vbTab & "QtyFirm" & vbTab & "QtySS" & vbTab & "QtyShort" & vbTab & "Price" & vbTab & "DiscountRate" & vbTab & "Ref" & vbTab _
    & "Extension" & vbTab & "Note" & vbTab & "POCode" & vbTab & "VATRATE"
    
    oTF.WriteToTextFile s
               
                
  lngNumberOfLines = 0
  For i = 1 To DeliveryLines.Count
    lngNumberOfLines = lngNumberOfLines + 1
    s = DeliveryLines(i).CodeF & vbTab
    s = s & vbTab & DeliveryLines(i).Title
    s = s & vbTab & DeliveryLines(i).QtyFirmF
    s = s & vbTab & DeliveryLines(i).QtySSF
    s = s & vbTab & DeliveryLines(i).QtyShortF
    s = s & vbTab & DeliveryLines(i).PriceF(ISForeignCurrency)
    s = s & vbTab & DeliveryLines(i).DiscountF
    s = s & vbTab & DeliveryLines(i).Ref
    s = s & vbTab & DeliveryLines(i).PLessDiscExtF(ISForeignCurrency)
    s = s & vbTab & DeliveryLines(i).Note
    s = s & vbTab & DeliveryLines(i).POCode
    s = s & vbTab & DeliveryLines(i).VATRateF
    oTF.WriteToTextFile s
Next

    oTF.WriteToTextFile "TotalNumberOfLines" & vbTab & CStr(Me.DeliveryLines.Count)
    oTF.WriteToTextFile "TotalQtyItems" & vbTab & BatchQtyTotalF
    oTF.WriteToTextFile "TotalValue" & vbTab & TotalPayableF(False)
    oTF.WriteToTextFile "TotalValueOnInvoice" & vbTab & BatchTotalF
    oTF.WriteToTextFile "Memo" & vbTab & Memo
    oTF.WriteToTextFile "StaffMember" & vbTab & IIf(Me.StaffName > "", Me.StaffName, "")
    oTF.CloseTextFile
    ExportToSpreadsheet = True
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.ExportToSpreadsheet(bForeign,pFilename)", Array(bForeign, pFilename)
End Function


Public Function docWriteTostream(ByVal FilePath As String, obj As MSXML2.DOMDocument60, _
                Optional ByVal CharSet As String = "UNICODE") As Boolean
    On Error GoTo errHandler
    Dim s As Object
    docWriteTostream = False
    Set s = CreateObject("ADODB.Stream")
    With s
        If CharSet <> "" Then .CharSet = CharSet
        .Open
        .WriteText obj.xml
        .SaveToFile FilePath, 2 'adSaveCreateOverWrite
        .Close
    End With
    docWriteTostream = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.docWriteTostream(FilePath,obj,CharSet)", Array(FilePath, obj, CharSet)
End Function

Public Function PrintDEL(bForeign As Boolean) As Boolean
    On Error GoTo errHandler
Dim strLocalPrinter As String
Dim oDC As a_DocumentControl

    
    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing Then
        PrintDEL = False
        Exit Function
    Else
        Set oPCL = New z_PrintClient
        If oPCL.StartRequestPrinting("DEL_" & Me.DOCCode) = False Then
            PrintDEL = False
            Exit Function
        End If
        With oDC
            oPCL.PrinterSetup .PrinterName(), .PreviewPrint, .Style, .QtyCopies
        End With
        PrintDEL = True
    End If
    
    If Me.StatusF <> "In PROCESS" And Me.StatusF <> "COMPLETE" Then
        oPCL.WriteTag "STATUS", Me.StatusF
    End If
    oPCL.LineCOunt DeliveryLines.Count
    oPCL.LineTotalCount 1
    oPCL.WriteTag "TOTAL", "Total" & "|" & Me.TotalPayableF(bForeign)
    oPCL.WriteTag "DOCUMENTCODE", Me.DOCCode
    oPCL.WriteTag "SINVREF", Me.SupplierInvoiceRef
    oPCL.WriteTag "STORECODE", oPC.Configuration.DefaultStore.code
    oPCL.WriteTag "SINVDATE", Me.SupplierInvoiceDateF
    oPCL.WriteTag "PHONE", IIf(Me.Supplier.BillTOAddress.Phone > "", Supplier.BillTOAddress.Phone, "")
    oPCL.WriteTag "SUPPLIER", Supplier.NameAndCode(35)
    oPCL.WriteTag "FAX", IIf(Supplier.BillTOAddress.Fax > "", "Fax: " & Supplier.BillTOAddress.Fax, "")
    oPCL.WriteTag "DATE", Me.IssDateF
    oPCL.WriteTag "MEMO", PackText(Memo)
    oPCL.WriteTag "COMPANYNAME", oPC.Configuration.DefaultCompany.CompanyName
    oPCL.WriteTag "SM", IIf(Me.StaffName > "", Me.StaffName, "")
    oPCL.WriteTag "OURADDRESS", PackText(oPC.Configuration.DefaultCompany.StreetandPostalAddress)
    FillTableRows Me, bForeign
    
    
    oPCL.FinishRequest
    
EXIT_Handler:
    Exit Function
errHandler:
    ErrPreserve
    oPCL.FinishRequest
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.PrintDEL(bForeign)", bForeign
End Function
Private Sub FillTableRows(pDel As a_Delivery, bForeign As Boolean)
    On Error GoTo errHandler
Dim i As Integer

    For i = 1 To pDel.DeliveryLines.Count
        With pDel.DeliveryLines(i)
            oPCL.DELLSend .CodeF, .QtyFirmF, .QtySSF, .QtyShort, .Ref, .Title, .PriceF(bForeign), .PLessDiscExtF(bForeign), .DiscountF, PackText(.Note), .POCode
        End With
    Next i
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery.FillTableRows(pDEL,bForeign)", Array(pDel, bForeign)
End Sub
