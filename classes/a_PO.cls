VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_PO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private oSM As z_StockManager
Private oFSO As FileSystemObject
Private xMLDoc As ujXML

Dim lngLPLessDiscExt_tot As Long
Dim lngFPLessDiscExt_tot As Long
Dim lngLPExt_tot As Long
Dim lngFPExt_tot As Long
Dim lngVAT_tot As Long
Dim lngDisc_tot As Long
Dim lngQty_tot As Long
Dim mOriginalStatus As enStatus
Dim oPCL As z_PrintClient
Event Valid(pMsg As String)
Event TotalChange(lngTotal As String, lngTotalForeign As String)
Event reloadlist()
Event CurrRowStatus(pMsg As String)
Event Dirty(pVal As Boolean)

Private mudtProps As POProps
Private mcolStack As Collection

Private WithEvents mPOL As ch_POL
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
Public Property Get constDOCCODE() As String
    constDOCCODE = CDOCCODE
End Property
Private Sub mPOL_CurrRowStatus(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent CurrRowStatus(pMsg)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.mPOL_CurrRowStatus(pMsg)", pMsg
End Sub
Public Property Get CaptureCurrency() As a_Currency
    On Error GoTo errHandler
    Set CaptureCurrency = moCaptureInCurrency
    Exit Property
errHandler:
    ErrorIn "a_PO.CaptureCurrency"
End Property
'Private Sub mPOL_RowsChange()
''occurs for new rows and deletion of rows
'
'    CalculateTotal
'    RaiseEvent TotalChange(lngTotalExtension, Format(lngTotalExtension / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString), _
'                        lngTotalDeposit, Format(lngTotalDeposit / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString), _
'                        lngTotalVAT, Format(lngTotalVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString))
'    SetDirty True
'End Sub
Public Function LevelNewValues(pPOL As a_POL)
    On Error GoTo errHandler
Dim oPOL As a_POL

    For Each oPOL In Me.POLines
        If (Not (oPOL Is pPOL)) And oPOL.PID = pPOL.PID Then
            oPOL.LocalPrice = pPOL.Price(False)
            oPOL.SetETA pPOL.ETA
            oPOL.SetSection pPOL.Section
            oPOL.ProductTypeID = pPOL.ProductTypeID
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.LevelNewValues(pPOL)", pPOL
End Function

Public Function CalculateTotals()
    On Error GoTo errHandler
    CalculateTotal_Local
    CalculateTotal_Foreign
    RaiseEvent TotalChange(Me.TotalLessDiscExtF(False), Me.TotalLessDiscExtF(True))
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CalculateTotals"
End Function
Public Function CalculateTotal_Local()
    On Error GoTo errHandler
Dim ol As a_POL
Dim bLineDiscountApplied As Boolean
    bLineDiscountApplied = False
    lngLPLessDiscExt_tot = 0
    lngLPExt_tot = 0
    lngVAT_tot = 0
    lngDisc_tot = 0
    lngQty_tot = 0
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
                lngQty_tot = lngQty_tot + ol.QtyFirm
            lngLPExt_tot = lngLPExt_tot + ol.PExt(False)
            lngVAT_tot = lngVAT_tot + ol.PLessDiscExtVAT(False)
            lngDisc_tot = lngDisc_tot + ol.PDiscExt(False)
            lngLPLessDiscExt_tot = lngLPLessDiscExt_tot + ol.PLessDiscExt(False)
        End If
    Next
    mudtProps.TotalQty = lngQty_tot
    mudtProps.TotalExtensionSimple = lngLPExt_tot
    mudtProps.TotalDiscount = lngDisc_tot
    mudtProps.TotalVAT = lngVAT_tot
    mudtProps.TotalExtension = lngLPLessDiscExt_tot
    mudtProps.TotalPayable = lngLPLessDiscExt_tot ' - lngLDeposit_tot
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'overflow
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CalculateTotal_Local"
End Function
Public Function CalculateTotal_Foreign()
    On Error GoTo errHandler
Dim ol As a_POL
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
    ErrPreserve
    If Err = 6 Then 'overflow
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CalculateTotal_Foreign"
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
    ErrorIn "a_PO.TotalLessDiscExtF(bForeign)", bForeign
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
    ErrorIn "a_PO.TotalExtF(bForeign)", bForeign
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
    ErrorIn "a_PO.TotalLessDiscExt(bForeign)", bForeign
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
    ErrorIn "a_PO.TotalVATF(bForeign)", bForeign
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
    ErrorIn "a_PO.TotalPayable(bForeign)", bForeign
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
    ErrorIn "a_PO.TotalPayableF(bForeign)", bForeign
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
    ErrorIn "a_PO.TotalDiscount(bForeign)", bForeign
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
    ErrorIn "a_PO.TotalDiscountF(bForeign)", bForeign
End Property
Public Property Get TotalQty() As Long
    On Error GoTo errHandler

        TotalQty = mudtProps.TotalQty

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TotalQty"
End Property
Public Property Get TotalQtyF() As String
    On Error GoTo errHandler

        TotalQtyF = CStr(TotalQty)

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TotalQtyF"
End Property

Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As POData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.GetState"
End Function
Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As POData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SetState(buffer)", buffer
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
    ErrorIn "a_PO.GetSuperState"
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
    ErrorIn "a_PO.SetSuperState(buffer)", buffer
End Sub


Public Sub BeginEdit()
    On Error GoTo errHandler
    mPOL.BeginEdit
    mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mPOL.CancelEdit
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
Dim objPersist As a_PO_P
Dim strSuffix As String
Dim strPos As String

'    CalculateTotal
strPos = "pos 1"
    mudtProps.TPID = Me.Supplier.ID
    'Set these values fromthe objects before saving -  the objects cant be saved
    If mudtProps.DocStoreID = 0 Then mudtProps.DocStoreID = oPC.Configuration.DefaultStoreID
strPos = "pos 2"
    If DOCCode = "" Then
        strPos = "pos 2a"
        strSuffix = oPC.GetProperty("SubscriptionOrderSuffix")
        DOCCode = FetchSONumber("GETANDWRITE") & IIf(ContainsCO = True, "X", "") & IIf(OrderType = "NS", strSuffix, "")
        strPos = "pos 2b"
        DOCDate = Now()
    End If
    If CaptureDate < #1/1/1995# And mudtProps.IsNew Then CaptureDate = Now
    
        strPos = "pos 3"
    If mcolStack.Count = 0 Then Err.Raise 445
    mPOL.ApplyEdit
    Set objPersist = New a_PO_P
    If mudtProps.IsDeleted Then
        strPos = "pos 4"
        objPersist.DeletePO mudtProps.TRID
        mcolStack.Remove mcolStack.Count
        mudtProps.IsNew = True
        mudtProps.IsDeleted = False
    ElseIf IsDirty Or mudtProps.IsNew Then
        strPos = "pos 5"
        If mudtProps.IsNew Then mudtProps.CaptureDate = Now()
        If Not IsValid Then Err.Raise 445
        mcolStack.Remove mcolStack.Count
        SetSuperState objPersist.Save(GetSuperState)
        mudtProps.IsNew = False
        strPos = "pos 5b"
    Else
        mcolStack.Remove mcolStack.Count
        strPos = "pos 6"
    End If
    strPos = "pos 7"
    Set objPersist = Nothing
    mudtProps.IsDirty = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.ApplyEdit", , , , "strPOS", Array(strPos)
End Sub
Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = mcolStack.Count > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty Or mPOL.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.IsDirty"
End Property

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.IsValid"
End Property
Public Function GetStatus()
    On Error GoTo errHandler
    If IsValid Then
        RaiseEvent Valid("")
    Else
        RaiseEvent Valid(TranslateErrors(mobjValid.AllBrokenRules))
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.GetSTatus"
End Function
Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    Set mPOL = New ch_POL
    mPOL.component Me
    Set mSupplier = New a_Supplier
    Set oPCL = New z_PrintClient
    mudtProps.IsNew = True
    mudtProps.Status = 2
    mOriginalStatus = stInProcess
    
    mudtProps.DocStoreID = oPC.Configuration.BillToStoreID
    mudtProps.DELTOStoreID = oPC.Configuration.DefaultStoreID
    mobjValid.BreakRule "TP", True
    LoadClassorsCollection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
    Set colClassors = Nothing
    Set mcolStack = Nothing
    Set mobjValid = Nothing
    Set mPOL = Nothing
    Set mSupplier = Nothing
    Set oPCL = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.mobjValid_Status(pMsg)", pMsg
End Sub

Public Sub Load(TRID As Long, ReadOnly As Boolean)
    On Error GoTo errHandler
Dim objPersist As a_PO_P
Dim oAdd As a_Address
Dim oComp As a_Company
Dim oCurr As a_Currency
    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    mudtProps.IsNew = False
    
    Set objPersist = New a_PO_P
    SetSuperState objPersist.Fetch(TRID)
    mOriginalStatus = Me.Status
    mSupplier.Load mudtProps.TPID
    Set objPersist = Nothing
    mobjValid.RuleBroken "TP", False
    CalculateTotals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Load(TRID,ReadOnly)", Array(TRID, ReadOnly)
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
    ErrorIn "a_PO.Reload"
End Sub
Public Property Get Supplier() As a_Supplier
    On Error GoTo errHandler
    Set Supplier = mSupplier
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Supplier"
End Property
Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Delete"
End Sub



Public Property Get SOID() As Long
    On Error GoTo errHandler
    SOID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SOID"
End Property
Public Property Let SOID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TRID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SOID(val)", val
End Property

Public Property Get BillToAddressID() As Long
    On Error GoTo errHandler
    BillToAddressID = mudtProps.DocStoreID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.BillToAddressID"
End Property
'Public Sub SetBillToAddress(val As a_Address)
'    If mcolStack.Count = 0 Then Err.Raise 383
'    Set oBillToAddress = val
' '   mudtProps.DocStoreID = o
'    SetDirty True
'End Sub
Public Property Get DELTOStoreID() As Long
    On Error GoTo errHandler
    DELTOStoreID = mudtProps.DELTOStoreID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DELTOStoreID"
End Property
Public Sub setDelToStoreID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DELTOStoreID = val
    mudtProps.DelToAddress = oPC.Configuration.Stores.FindStoreByID(val).DelAddress
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.setDelToStoreID(val)", val
End Sub
Public Sub setDelToStoreIDImmediate(val As Long)
    On Error GoTo errHandler
Dim oSM As New z_StockManager

    mudtProps.DELTOStoreID = val
    mudtProps.DelToAddress = oPC.Configuration.Stores.FindStoreByID(val).DelAddress
    oSM.SetAddressImmediate "DEL", Me.TRID, mudtProps.DELTOStoreID
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.setDelToStoreIDImmediate(val)", val
End Sub
Public Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    mudtProps.IsDirty = pVal
    RaiseEvent Dirty(pVal)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SetDirty(pVal)", pVal
End Sub

Public Property Get TRID() As Long
    On Error GoTo errHandler
    TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TRID"
End Property
Public Property Let TRID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TRID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TRID(val)", val
End Property
Public Property Let ContainsCO(p As Boolean)
    On Error GoTo errHandler
    mudtProps.ContainsCO = p
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.ContainsCO(p)", p
End Property
Public Property Get ContainsCO() As Boolean
    On Error GoTo errHandler
    ContainsCO = mudtProps.ContainsCO
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.ContainsCO"
End Property
'Public Function SetGeneralDiscount(val As String) As Boolean
'Dim dblDiscount As Double
'Dim oPO As a_POL
'
'    If mcolStack.Count = 0 Then Err.Raise 383
'    SetGeneralDiscount = True
'    If Trim$(val) = "" Then
'        dblDiscount = 0
'    ElseIf Not ConvertToDBL(val, dblDiscount) Then
'        SetGeneralDiscount = False
'        Exit Function
'    End If
'    If (dblDiscount <= 0) Then
'        mobjValid.RuleBroken "Discount", True
'    Else
'        mobjValid.RuleBroken "Discount", False
'    End If
'    mudtProps.DiscountRate = dblDiscount
'    SetDirty True
'    For Each oPO In InvoiceLines  'disctribute discount rate over all lines where non nonstaock items exists
'        If Not oPO.ServiceItem Then
'            oPO.Discount = dblDiscount
'        End If
'    Next
'    mPOL_RowsChange  'recalulates andd raises event
'    RaiseEvent reloadlist
'End Function
Public Function SetTP(pTPID As Long) As Boolean
    On Error GoTo errHandler
Dim bSuccess As Boolean
    bSuccess = mSupplier.Load(pTPID)
    SetCaptureCurrency Supplier.DefaultCurrency
    SetTP = bSuccess
    If bSuccess Then
        mobjValid.RuleBroken "TP", False
    End If
    SetDirty True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SetTP(pTPID)", pTPID
End Function
Public Function SetSupplierFromAccNum(pAccnum As String) As Boolean
    On Error GoTo errHandler
Dim bSuccess As Boolean

    bSuccess = Me.Supplier.Load(, pAccnum)
    SetSupplierFromAccNum = bSuccess
    Set moCaptureInCurrency = Supplier.DefaultCurrency

    If bSuccess Then
        mobjValid.RuleBroken "TP", False
'        If Me.BillToAddressID = 0 Then
'            Set oBillToAddress = Me.Supplier.DefaultAddress
'        End If
'        If Me.DELTOStoreID = 0 Then
'            Set oDelToAddress = Me.Supplier.DefaultAddress
'        End If
    End If
    mudtProps.DispatchMethodID = Supplier.DispatchModeID
    SetDirty True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SetSupplierFromAccNum(pAccnum)", pAccnum
End Function
Public Function SetCaptureCurrency(val As a_Currency) As Boolean
    On Error GoTo errHandler
    Set moCaptureInCurrency = val
    mudtProps.CurrencyID = val.ID
    mudtProps.CurrencyRate = val.Factor
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SetCaptureCurrency(val)", val
End Function
Public Property Get StaffID() As Long
    On Error GoTo errHandler
    StaffID = mudtProps.StaffID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.StaffID"
End Property
Public Property Let StaffID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.StaffID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.StaffID(val)", val
End Property
Public Property Get StaffName() As String
    On Error GoTo errHandler
    If mudtProps.StaffID < 1 Then
        StaffName = ""
        Exit Property
    End If
    StaffName = oPC.Configuration.Staff.FindStaffByID(StaffID).StaffName
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.StaffName"
End Property
Public Property Get Signature()
Signature = FNS(mudtProps.Signature)
End Property
Public Property Get SendersEmail()
    SendersEmail = FNS(mudtProps.StaffEmail)
End Property
Public Property Get StaffNameB() As String
    On Error GoTo errHandler
Dim strT As String
    If mudtProps.StaffID < 1 Then
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
    ErrorIn "a_PO.StaffNameB"
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
    ErrorIn "a_PO.TPName"
End Property
Public Property Let TPNAME(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPNAME = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TPName(val)", val
End Property

Public Property Get TPPhone() As String
    On Error GoTo errHandler
    TPPhone = FNS(mudtProps.TPPhone)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TPPhone"
End Property
Public Property Let TPPhone(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPPhone = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TPPhone(val)", val
End Property

Public Property Get TPFax() As String
    On Error GoTo errHandler
    TPFax = FNS(mudtProps.TPFax)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TPFax"
End Property
Public Property Let TPFax(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPFax = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TPFax(val)", val
End Property
Public Property Let tmpETA(val As Date)
    mudtProps.TempETA = val
End Property
Public Property Get tmpETA() As Date
    tmpETA = mudtProps.TempETA
End Property
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
    ErrorIn "a_PO.Memo"
End Property
Public Property Let Memo(val As String)
    On Error GoTo errHandler
    mudtProps.Memo = FNS(val)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Memo(val)", val
End Property


'Public Property Get TPAccNum() As String
'    TPAccNum = FNS(mudtProps.TPAccNum)
'End Property
'Public Property Let TPAccNum(val As String)
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.TPAccNum = val
'    SetDirty True
'End Property
'
'Public Property Get TPMemo() As String
'    TPMemo = FNS(mudtProps.TPMemo)
'End Property
'Public Property Let TPMemo(val As String)
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.TPMemo = val
'    SetDirty True
'End Property

'Public Property Get BusPhone() As String
'    BusPhone = FNS(mudtProps.BusPhone)
'End Property
'Public Property Let BusPhone(val As String)
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.BusPhone = val
'    SetDirty True
'End Property
Public Property Get OrderType() As String
    On Error GoTo errHandler
    OrderType = FNS(mudtProps.OrderType)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.OrderType"
End Property
Public Property Let OrderType(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.OrderType = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.OrderType(val)", val
End Property

Public Property Get DOCCode() As String
    On Error GoTo errHandler
    DOCCode = FNS(mudtProps.DOCCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DocCode"
End Property
Public Property Let DOCCode(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DOCCode = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DocCode(val)", val
End Property
Public Property Get TPACCNum() As String
    On Error GoTo errHandler
    TPACCNum = FNS(mudtProps.TPACCNum)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TPACCNum"
End Property
Public Property Let TPACCNum(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPACCNum = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TPACCNum(val)", val
End Property
Public Property Get DeliverToAddress() As String
    On Error GoTo errHandler
    DeliverToAddress = FNS(mudtProps.DelToAddress)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DeliverToAddress"
End Property
Public Property Let DeliverToAddress(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DelToAddress = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DeliverToAddress(val)", val
End Property
Public Sub SetBillToAddressImmediate(val As a_Address)
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    Set oBillToAddress = val
    mudtProps.BillToAddressID = oBillToAddress.ID
    oSM.SetAddressImmediate "BILL", Me.TRID, oBillToAddress.ID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SetBillToAddressImmediate(val)", val
End Sub
Public Sub setDelToAddressImmediate(val As a_Address)
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    Set oDelToAddress = val
    mudtProps.DelToAddressID = oDelToAddress.ID
    oSM.SetAddressImmediate "DEL", Me.TRID, oDelToAddress.ID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.setDelToAddressImmediate(val)", val
End Sub
Public Property Get IssDate() As Date
    On Error GoTo errHandler
    IssDate = mudtProps.CaptureDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.IssDate"
End Property
Public Property Get IssDateF() As Date
    On Error GoTo errHandler
    IssDateF = Format(mudtProps.CaptureDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.IssDateF"
End Property

'Public Property Get CurrencyFormat() As String
'    CurrencyFormat = FNS(mudtProps.CurrencyFormat)
'    If CurrencyFormat <= "" Then CurrencyFormat = oPC.Configuration.Currencies(oPC.Configuration.DefaultCurrencyID & "k").FormatString
'End Property
'Public Property Let CurrencyFormat(val As String)
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.CurrencyFormat = val
'    SetDirty True
'End Property

Public Property Get DOCDate() As Date
    On Error GoTo errHandler
    DOCDate = mudtProps.DOCDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DocDate"
End Property
Public Property Get DocDateF() As String
    On Error GoTo errHandler
    DocDateF = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DocDateF"
End Property
Public Property Get ProcessingDate() As Date
    On Error GoTo errHandler
    ProcessingDate = mudtProps.ProcessingDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.ProcessingDate"
End Property
Public Property Get ProcessingDateF() As String
    On Error GoTo errHandler
    ProcessingDateF = Format(mudtProps.ProcessingDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.ProcessingDateF"
End Property
Public Property Get ProcessingDateFF() As String
    On Error GoTo errHandler
    If mudtProps.ProcessingDate < "2010-01-01" Then
        ProcessingDateFF = ""
    Else
        ProcessingDateFF = Format(mudtProps.ProcessingDate, "dd/mm/yyyy Hh:Nn")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.ProcessingDateFF"
End Property

Public Property Get TransDateFormatted() As String
    On Error GoTo errHandler
    TransDateFormatted = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TransDateFormatted"
End Property
Public Property Let DOCDate(val As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    If IsDate(val) Then
        mudtProps.DOCDate = CDate(val)
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DocDate(val)", val
End Property

Public Property Get CaptureDate() As Date
    On Error GoTo errHandler
    CaptureDate = mudtProps.CaptureDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CaptureDate"
End Property
Public Property Get CaptureDateFormatted() As Date
    On Error GoTo errHandler
    CaptureDateFormatted = Format(mudtProps.CaptureDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CaptureDateFormatted"
End Property
Public Property Let CaptureDate(val As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CaptureDate = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CaptureDate(val)", val
End Property

'Public Property Get ForeignCurrency() As a_Currency
'    Set ForeignCurrency = CaptureCurrency
'End Property

'Public Property Get CurrencyFactor() As Double
'    If mudtProps.CurrencyFactor = 0 Then mudtProps.CurrencyFactor = oPC.Configuration.Currencies(oPC.Configuration.DefaultCurrencyID & "k").Factor
'    CurrencyFactor = mudtProps.CurrencyFactor
'End Property
'Public Property Let CurrencyFactor(val As Double)
'    mudtProps.CurrencyFactor = val
'    SetDirty True
'End Property
'
'Public Property Get VATRate() As Double
'    VATRate = mudtProps.VATRate
'End Property
'Public Property Get VATRateFormatted() As String
'    VATRate = Format(mudtProps.VATRate, "##0.00")
'End Property
'Public Property Let VATRate(val As Double)
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.VATRate = val
'    mudtProps.IsDirty = True
'End Property

Public Property Get OrderDate() As Date
    On Error GoTo errHandler
    OrderDate = mudtProps.DOCDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.OrderDate"
End Property
Public Property Get OrderDateFormatted() As String
    On Error GoTo errHandler
    OrderDateFormatted = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.OrderDateFormatted"
End Property
Public Property Let OrderDate(val As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DOCDate = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.OrderDate(val)", val
End Property

Public Property Get DispatchModeID() As Long
    DispatchModeID = FNS(mudtProps.DispatchMethodID)
End Property
Public Function SetDispatchModeID(val As Long) As Boolean
Dim bOK As Boolean
  '  If mcolStack.Count = 0 Then Err.Raise 383
        mudtProps.DispatchMethodID = val
End Function

Public Property Get POLines() As ch_POL
    On Error GoTo errHandler
  Set POLines = mPOL
    Exit Property
errHandler:
    ErrorIn "a_PO.POLines"
End Property
Public Property Get DisplayPO() As d_PO
    On Error GoTo errHandler
    Set DisplayPO = New d_PO
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DisplayPO"
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
    ErrorIn "a_PO.StatusF"
End Property
Public Property Get StatusForPrinting() As String
    On Error GoTo errHandler
    If mudtProps.Status = stVOID Then
        StatusForPrinting = "VOID"
    ElseIf mudtProps.Status = stInProcess Then
        StatusForPrinting = "IN PROCESS"
    ElseIf mudtProps.Status = stISSUED Then
        StatusForPrinting = ""
    ElseIf mudtProps.Status = stCOMPLETE Then
        StatusForPrinting = ""
    ElseIf mudtProps.Status = stPROFORMA Then
        StatusForPrinting = ""
    ElseIf mudtProps.Status = stCANCELLED Then
        StatusForPrinting = "CANCELLED"
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.StatusForPrinting"
End Property
Public Property Get Status() As enStatus
    On Error GoTo errHandler
    Status = mudtProps.Status
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Status"
End Property
Public Function FetchSONumber(strTmp As String) As String
    On Error GoTo errHandler
Dim oCode As z_Code
Dim strDefaultCo As String
Dim strPONumber As String
    Set oCode = New z_Code
    
    strPONumber = oCode.GetNextCode(enSupplierOrder, strTmp)
    
    strDefaultCo = oPC.Configuration.DefaultCompany.CompanyCode
    FetchSONumber = strDefaultCo & "P" & strPONumber
    
    Set oCode = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.FetchSONumber(strTmp)", strTmp
End Function

'Property Get BillTOAddress() As a_Address
'    Set BillTOAddress = oBillToAddress
'End Property
'
'Property Get DelToAddress() As a_Address
'    Set DelToAddress = oDelToAddress
'End Property
Public Function PrintPO_Display(bForeign As Boolean) As String
    On Error GoTo errHandler
Dim bDiscountExists As Boolean
Dim oPO As a_POL
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
    ErrorIn "a_PO.PrintPO_Display(bForeign)", bForeign
End Function
Public Function GetDocumentAddresses(pBillTo As String, pDelto As String) As Boolean
Dim oTmpStore As a_Store

    Set oTmpStore = oPC.Configuration.Stores.FindStoreByID(Me.DELTOStoreID)
    If Not oTmpStore Is Nothing Then
        oPCL.WriteTag "DELTO", PackText(oTmpStore.DelAddress)
        pDelto = oTmpStore.DelAddress
    End If
    Set oTmpStore = oPC.Configuration.Stores.FindStoreByID(mudtProps.DocStoreID)
    If Not oTmpStore Is Nothing Then
        oPCL.WriteTag "BILLTO", PackText(oTmpStore.BillAddress)
        pBillTo = oTmpStore.BillAddress
    End If
End Function
Public Function GetBilltoAddressasString() As String
Dim oTmpStore As a_Store
    Set oTmpStore = oPC.Configuration.Stores.FindStoreByID(mudtProps.DocStoreID)
    If Not oTmpStore Is Nothing Then
        GetBilltoAddressasString = oTmpStore.BillAddress
    Else
        GetBilltoAddressasString = ""
    End If
End Function
Public Function GetDeltoAddressasString() As String
Dim oTmpStore As a_Store
    Set oTmpStore = oPC.Configuration.Stores.FindStoreByID(mudtProps.DELTOStoreID)
    If Not oTmpStore Is Nothing Then
        GetDeltoAddressasString = oTmpStore.DelAddress
    Else
        GetDeltoAddressasString = ""
    End If
End Function
Public Function ExportToXML(bForeign As Boolean, DispatchMode As enTransmitType, _
            pFilename As String, _
            Optional pDestinationEmailAddress As String, _
            Optional pWholeMessage As String, Optional pQtyCopies As Integer, Optional QuickPrint As Boolean) As Boolean
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim strPath As String
Dim strBillto As String
Dim strDelto As String
Dim strFOFile As String
Dim strPDFFile As String
Dim strEDIFilePath As String
Dim strXML As String
Dim strCommand As String
Dim sFormatFilename As String
Dim i As Integer
Dim strHTML As String
Dim fs As New FileSystemObject
Dim objXSL As New MSXML2.DOMDocument60
Dim opXMLDOC As New MSXML2.DOMDocument60
Dim objXMLDOC  As New MSXML2.DOMDocument60
Dim strWorkingFolder As String
Dim strDispatchMethod As String
Dim oDC As a_DocumentControl
Dim strExecutable As String
Dim cnt As Integer
    If DispatchMode = enMail Then
        If oPC.UsesOutlookForPOEmail Then
            strWorkingFolder = oPC.LocalFolder & "TEMP\"
            strDispatchMethod = "Outlook"
        Else
            strWorkingFolder = oPC.SharedFolderRoot & "\Emails\"
            strDispatchMethod = "Emailing"
        End If
    ElseIf DispatchMode = enPrint Then
  '  MsgBox "Code change -tell david Clarke"
  '  oPC.SharedFolderRoot = "\\David-PC\PBKS_S"
        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
        strDispatchMethod = "Printing"
    ElseIf DispatchMode = enView Then
        strWorkingFolder = oPC.SharedFolderRoot & "\PDF\"
        strDispatchMethod = ""
    ElseIf DispatchMode = enEDI Then
        strWorkingFolder = oPC.LocalFolder & "TEMP\"
        strDispatchMethod = "EDI"
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
                        p 2
    If pQtyCopies < 1 Then pQtyCopies = 1

    Set xMLDoc = New ujXML
                        p 21
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "PO_DOC"
            .chCreate "MessageType"
                .elText = "PURCHASEORDER"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "DestinationAddress"
                If oPC.TestMode Then
                    .elText = oPC.EmailAddressForTesting
                    pDestinationEmailAddress = oPC.EmailAddressForTesting
                Else
                    If Me.Supplier Is Nothing Then
                        .elText = ""
                        pDestinationEmailAddress = ""
                    Else
                        If Me.Supplier.OrderToAddress Is Nothing Then
                            .elText = ""
                            pDestinationEmailAddress = ""
                        Else
                            .elText = Me.Supplier.OrderToAddress.EMail
                            pDestinationEmailAddress = Me.Supplier.OrderToAddress.EMail
                        End If
                    End If
                End If
                        p 23
            .elCreateSibling "TemplateName"
                .elText = "PO_DOC"
            .elCreateSibling "SendersEmail"
                If oPC.EmailFrom > "" Then
                    .elText = oPC.EmailFrom
                Else
                    .elText = Me.SendersEmail
                End If
            .elCreateSibling "CopyCount"
                .elText = pQtyCopies
            .elCreateSibling "Printer"
                If Not oDC Is Nothing Then .elText = oDC.PrinterName
            .elCreateSibling "Status"
                .elText = StatusForPrinting
            .elCreateSibling "AccompanyingMessage"
                .elText = oPC.Configuration.EmailPOMsg
            .elCreateSibling "DispatchMethod"
                .elText = Me.Supplier.DispatchModes.ItemByOrdinalIndex(Me.Supplier.DispatchModes.FindIndexByKey(Me.DispatchModeID))
            .elCreateSibling "LogoPath"
                If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
                Else
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
                End If
            .elCreateSibling "DocCode"
                .elText = Me.DOCCode
            .elCreateSibling "DocDate", True
                .elText = Me.DOCDate
            .elCreateSibling "Sender", True
                .elText = oPC.Configuration.DefaultCompany.CompanyName
            .elCreateSibling "SenderAddress", True
                .elText = Replace(oPC.Configuration.DefaultCompany.StreetAddress, Chr(13) & Chr(10), Chr(10))
            .elCreateSibling "SupplierName", True
                .elText = Supplier.NameAndCode(35)
            .elCreateSibling "SupplierWithAddress", True
                If Supplier.OrderToAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Replace(Supplier.OrderToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
                End If
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
            .elCreateSibling "BillTo", True
                If Me.BillToAddressID > 0 Then
                    If Not oPC.Configuration.Stores.FindStoreByID(Me.BillToAddressID) Is Nothing Then
                        .elText = Replace(oPC.Configuration.Stores.FindStoreByID(Me.BillToAddressID).BillAddress, Chr(13) & Chr(10), Chr(10))
                    Else
                        .elText = Replace(oPC.Configuration.DefaultStore.BillAddress, Chr(13) & Chr(10), Chr(10))
                    End If
                End If
            .elCreateSibling "DelTo", True
                If Me.DELTOStoreID > 0 Then
                    If Not oPC.Configuration.Stores.FindStoreByID(Me.DELTOStoreID) Is Nothing Then
                        .elText = Replace(oPC.Configuration.Stores.FindStoreByID(Me.DELTOStoreID).DelAddress, Chr(13) & Chr(10), Chr(10))
                    Else
                        .elText = Replace(oPC.Configuration.DefaultStore.DelAddress, Chr(13) & Chr(10), Chr(10))
                    End If
                End If
                            p 24
            
            For i = 1 To Me.POLines.Count
                    .elCreateSibling "DetailLine", True
                    .chCreate "SKU"
                    .elText = POLines.Item(i).ProductCodeForExport
                    .elCreateSibling "Title", True
                    If POLines(i).Fulfilled <> "CAN" Then
                       .elText = POLines(i).TitleAuthor
                    Else
                       .elText = "***CANCELLED***" & POLines(i).TitleAuthor
                    End If
                    .elCreateSibling "QtyFirm", True
                        .elText = POLines(i).QtyFirmF
                    .elCreateSibling "QtySS", True
                        .elText = POLines(i).QtySSF
                    .elCreateSibling "Price", True
                        .elText = POLines(i).PriceF(bForeign)
                    .elCreateSibling "DiscountRate", True
                        .elText = POLines(i).DiscountF
                    .elCreateSibling "Reference", True
                        .elText = POLines(i).Ref
                    .elCreateSibling "Extension", True
                        .elText = POLines(i).PLessDiscExtF(bForeign)
                    .elCreateSibling "Note", True
                        .elText = POLines(i).Note
                    .navUP
            Next i
                            p 25
            .elCreateSibling "TotalNumberOfLines", True
                .elText = CStr(Me.POLines.Count)
            .elCreateSibling "TotalText", True
                .elText = "Total"
            .elCreateSibling "TotalNumbers", True
                .elText = TotalPayableF(bForeign)
            .elCreateSibling "Memo", True
                .elText = Memo
            .elCreateSibling "CompanyRegistration", True
                .elText = oPC.Configuration.DefaultCompany.CoRegistrationNumber
            .elCreateSibling "VATNumber", True
                .elText = oPC.Configuration.DefaultCompany.VatNumber
            .elCreateSibling "StaffMember", True
                If oPC.Configuration.Staff.FindStaffByID(Me.StaffID) Is Nothing Then
                    .elText = ""
                Else
                    .elText = oPC.Configuration.Staff.FindStaffByID(Me.StaffID).StaffName
                End If
            .elCreateSibling "OrderMessage", True
                .elText = oPC.Configuration.OrderText
    End With
                            p 26
'FINALLY PRODUCE THE .XML FILE
    strXML = strWorkingFolder & "PO_" & Me.DOCCode & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With
                            p 27
    If DispatchMode = enMail Then
        If oPC.EmailPOShowHTML = True Then
'WRITE THE .HTML FILE--------------------------------------------------------------------------------------------------
            objXSL.async = False
            objXSL.ValidateOnParse = False
            objXSL.resolveExternals = False
            strPath = oPC.SharedFolderRoot & "\Templates\PO_DOC_HTML.xslt"
            Set fs = New FileSystemObject
            If fs.FileExists(strPath) Then
                objXSL.Load strPath
            End If
        
            If fs.FileExists(strWorkingFolder & "PO_" & Me.DOCCode & ".HTML") Then
                fs.DeleteFile strWorkingFolder & "PO_" & Me.DOCCode & ".HTML", True
            End If
            oTF.OpenTextFileToAppend strWorkingFolder & "PO_" & Me.DOCCode & ".HTML"
            pWholeMessage = xMLDoc.docObject.transformNode(objXSL)
            oTF.WriteToTextFile pWholeMessage 'xMLDoc.docObject.transformNode(objXSL)
            oTF.CloseTextFile
        End If
    End If
    
                        p 27
'WRITE THE .PDF FILE IF NECESSARY ------------------------------------------------------------------------------------
'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file
    If (Me.Supplier.DispatchMethod = "M" And DispatchMode = enMail) Or DispatchMode = enView Or (DispatchMode = enPrint And QuickPrint = True) Then
        Set objXSL = Nothing
        Set objXSL = New MSXML2.DOMDocument60
        objXSL.async = False
        objXSL.ValidateOnParse = False
        objXSL.resolveExternals = False
        strPath = oPC.SharedFolderRoot & "\Templates\PO_DOC_FO.xsl"
        Set fs = New FileSystemObject
        If fs.FileExists(strPath) Then
            objXSL.Load strPath
        End If
                        p 6
        Set opXMLDOC = New MSXML2.DOMDocument60
        opXMLDOC.async = False
        opXMLDOC.ValidateOnParse = False
        opXMLDOC.resolveExternals = False
        xMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC
        
        strFOFile = strWorkingFolder & "PO_" & Me.DOCCode & ".FO"
        strPDFFile = strWorkingFolder & "PO_" & Me.DOCCode & ".PDF"
                                p 7
        If docWriteTostream(strFOFile, opXMLDOC, "UNICODE") = False Then
            MsgBox "Cannot write to .FO file. Please contact support to check the permissions for writing " & strFOFile & " are correctly set.", vbInformation + vbOKOnly, "Can't do this"
            oTF.CloseTextFile
            Exit Function
        End If
        p 7, "a"
'Stage 2 Convert the .FO file to .PDF and clean up
        If fs.FileExists(strPDFFile) Then
            On Error Resume Next
            fs.DeleteFile strPDFFile
            If Err Then
                MsgBox "It looks like you have this document open already in your PDF reader. Close the document before re-processing it.", vbInformation + vbOKOnly, "Warning"
            End If
        End If
        p 7, "b"
        ChDir "\PBKS\Executables\FOP\"
        strCommand = GetFOPCommandstring(strFOFile, strPDFFile)
        p 7, "c"
        Dim fl As New z_TextFile
        If fs.FileExists("c:\PBKS\TEMP\CREATEPDF.BAT") Then
          fs.DeleteFile ("c:\PBKS\TEMP\CREATEPDF.BAT")
        End If
        fl.OpenTextFile ("c:\PBKS\TEMP\CREATEPDF.BAT")
        fl.WriteToTextFile (strCommand)
        fl.CloseTextFile
        Shell "c:\PBKS\TEMP\CREATEPDF.BAT", vbHide
        Set fl = Nothing
        p 7, "d"
        'IF IT FAILS HERE CHECK THAT JAVA IS INSTALLED
        'We do the following because the batch file executes and is not waited for by the F_7_AB_1_ShellAndWaitSimple routine
        'so the FO file is sometimes deleted before the .PDF file is produced.
    End If
    
    MsgWaitObj 3000   ' allow for closing of file after writing
p 7, "e"
'Prepare EDI file and dispatch---------------------------------------------------------------------------------------------------------
    If (Me.Supplier.DispatchMethod = "E" And DispatchMode = enEDI) Then
p 7, "f"
'        'First create equivalent EDI file (same name different suffix)
            strEDIFilePath = "C:\PBKS\EDI_Files\" & fs.GetBaseName(strXML) & ".EDI"
            sFormatFilename = "C:\mec\eagle\formatdescription\edifact.96.a.orders.xml"
            strCommand = "jstart2.BAT de.mendelson.eagle.converter.edixml.EDIXMLConverter -xmlin " & strXML & " -formatin " & sFormatFilename & " -ediout " & strEDIFilePath & "  -ncs"
            ChDir "c:\MEC\EAGLE"
            F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 600000
    
    End If
'GenerateSAANAMsg
                        p 8
'Log the dispatch
    Set oSM = New z_StockManager
    Set oFSO = New FileSystemObject
    pFilename = strPDFFile
    
    If DispatchMode <> enMail And DispatchMode <> enView Then
        oSM.LogTransmission Me.TRID, "Dispatched: by " & strDispatchMethod & " : " & Format(Date, "dd/mm/yyyy") & vbCrLf
        Log = "Dispatched: by " & strDispatchMethod & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
    End If
    Set oSM = Nothing
    Set oFSO = Nothing
                        p 9
                        
    ExportToXML = True
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
    ErrorIn "a_PO.ExportToXML(bForeign)", bForeign, , , "strPos", Array(strErrPos)
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
    oTF.WriteToTextFile "Sender" & vbTab & oPC.Configuration.DefaultCompany.CompanyName
    oTF.WriteToTextFile "SenderAddress" & vbTab & Replace(oPC.Configuration.DefaultCompany.StreetAddress, Chr(13) & Chr(10), Chr(10))
    oTF.WriteToTextFile "SupplierName" & vbTab & Supplier.NameAndCode(70)
    oTF.WriteToTextFile "ACNO" & vbTab & IIf(Me.Supplier.AcNo > "", "Ac/no. " & Me.Supplier.AcNo, "")
    
    oTF.WriteToTextFile_NoLineTerminator "SupplierWithAddress" & vbTab
    If Not Supplier.OrderToAddress Is Nothing Then
        oTF.WriteToTextFile Replace(Supplier.OrderToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
    End If
    
    oTF.WriteToTextFile_NoLineTerminator "SupplierPhone" & vbTab
    If Not Supplier.OrderToAddress Is Nothing Then
        oTF.WriteToTextFile IIf(Me.Supplier.BillTOAddress.Phone > "", "Phone: " & Supplier.BillTOAddress.Phone, "")
    End If
    oTF.WriteToTextFile_NoLineTerminator "SupplierFax" & vbTab
    If Not Supplier.OrderToAddress Is Nothing Then
        oTF.WriteToTextFile IIf(Supplier.BillTOAddress.Fax > "", "Fax: " & Supplier.BillTOAddress.Fax, "")
    End If

                If Me.BillToAddressID > 0 Then
                    If Not oPC.Configuration.Stores.FindStoreByID(Me.BillToAddressID) Is Nothing Then
                        s = Replace(oPC.Configuration.Stores.FindStoreByID(Me.BillToAddressID).BillAddress, Chr(13) & Chr(10), Chr(10))
                    Else
                        s = Replace(oPC.Configuration.DefaultStore.BillAddress, Chr(13) & Chr(10), Chr(10))
                    End If
                End If
    oTF.WriteToTextFile "BillTo" & vbTab & s

                If Me.DELTOStoreID > 0 Then
                    If Not oPC.Configuration.Stores.FindStoreByID(Me.DELTOStoreID) Is Nothing Then
                        s = Replace(oPC.Configuration.Stores.FindStoreByID(Me.DELTOStoreID).DelAddress, Chr(13) & Chr(10), Chr(10))
                    Else
                        s = Replace(oPC.Configuration.DefaultStore.DelAddress, Chr(13) & Chr(10), Chr(10))
                    End If
                End If
    oTF.WriteToTextFile "DelTo" & vbTab & s

    s = "SKU" & vbTab & "Title" & vbTab & "QtyFirm" & vbTab & "QtySS" & vbTab & "Price" & vbTab & "DiscountRate" & vbTab & "Ref" & vbTab _
    & "Extension" & vbTab & "Note"
    
    oTF.WriteToTextFile s
               
                
  lngNumberOfLines = 0
  For i = 1 To POLines.Count
    lngNumberOfLines = lngNumberOfLines + 1
    s = POLines.Item(i).ProductCodeForExport & vbTab
                    If POLines(i).Fulfilled <> "CAN" Then
                       s2 = POLines(i).TitleAuthor
                    Else
                       s2 = "***CANCELLED***" & POLines(i).TitleAuthor
                    End If
    s = s & s2
    s = s & vbTab & POLines(i).QtyFirmF
    s = s & vbTab & POLines(i).QtySSF
    s = s & vbTab & POLines(i).PriceF(bForeign)
    s = s & vbTab & POLines(i).DiscountF
    s = s & vbTab & POLines(i).Ref
    s = s & vbTab & POLines(i).PLessDiscExtF(bForeign)
    s = s & vbTab & POLines(i).Note
    oTF.WriteToTextFile s
Next

    oTF.WriteToTextFile "TotalNumberOfLines" & vbTab & CStr(Me.POLines.Count)
    oTF.WriteToTextFile "Memo" & vbTab & Memo
    oTF.WriteToTextFile "StaffMember" & vbTab & IIf(Me.StaffName > "", Me.StaffName, "")
    oTF.WriteToTextFile "TotalIncVAT" & vbTab & TotalPayableF(bForeign)
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
    ErrPreserve
    If Err = 3004 Then
        LogSaveToFile Error & " in a_PO:docWriteTostream"
        Exit Function
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.docWriteTostream(FilePath,Charset)", Array(FilePath, CharSet)
End Function

Public Function PrintPO(pShowVAT As Boolean, bForeign As Boolean, Optional ByVal pRoundedUp As Boolean) As Boolean
    On Error GoTo errHandler
Dim bDiscountExists As Boolean
Dim tmp As String
Dim strLocalPrinter As String
Dim oDC As a_DocumentControl
Dim strBillto As String
Dim strDelto As String

    
    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing Then
        PrintPO = False
        Exit Function
    Else
        Set oPCL = New z_PrintClient
        If oPCL.StartRequestPrinting("PO_" & Me.DOCCode) = False Then
            PrintPO = False
            Exit Function
        End If
        With oDC
            oPCL.PrinterSetup .PrinterName(), .PreviewPrint, .Style, .QtyCopies
        End With
        PrintPO = True
    End If
    
    
    
    If Me.Status <> stISSUED And Me.Status <> stCOMPLETE Then
        oPCL.WriteTag "STATUS", Me.StatusF
    End If
    oPCL.LineCOunt POLines.Count
    oPCL.LineTotalCount 1
    oPCL.WriteTag "TOTAL", "Total" & "|" & Me.TotalPayableF(bForeign)
    oPCL.WriteTag "DOCUMENTCODE", Me.DOCCode
    oPCL.WriteTag "SENDER", oPC.Configuration.DefaultCompany.CompanyName
    oPCL.WriteTag "OURADDRESS", PackText(oPC.Configuration.DefaultCompany.StreetandPostalAddress)
    oPCL.WriteTag "SUPPLIERADDRESS", PackText(Supplier.OrderToAddress.AddressMailing)
    oPCL.WriteTag "SUPPLIER", Supplier.NameAndCode(35)
    oPCL.WriteTag "PHONE", IIf(Me.Supplier.BillTOAddress.Phone > "", "Phone: " & Supplier.BillTOAddress.Phone, "")
    oPCL.WriteTag "FAX", IIf(Supplier.BillTOAddress.Fax > "", "Fax: " & Supplier.BillTOAddress.Fax, "")
    
    If GetDocumentAddresses(strBillto, strDelto) Then
    '    oPCL.WriteTag "DELTO", PackText(strDelto)
    '    oPCL.WriteTag "BILLTO", PackText(strBillto)
    End If
    
'    Set oTmpStore = oPC.Configuration.Stores.FindStoreByID(Me.DELTOStoreID)
'    If Not oTmpStore Is Nothing Then oPCL.WriteTag "DELTO", PackText(oTmpStore.DelAddress)
'    If Not oTmpStore Is Nothing Then oPCL.WriteTag "BILLTO", PackText(oTmpStore.BillAddress)
    oPCL.WriteTag "DATE", Me.DOCDate
    oPCL.WriteTag "MEMO", PackText(Memo)
    oPCL.WriteTag "FOOT", "Co. Reg. No. " & oPC.Configuration.DefaultCompany.CoRegistrationNumber & "         V.A.T. Reg. No. " & oPC.Configuration.DefaultCompany.VatNumber
    oPCL.WriteTag "MAINMSG", PackText(oPC.Configuration.OrderText)
    oPCL.WriteTag "SM", IIf(Me.StaffName > "", Me.StaffName, "")
    FillTableRows Me, bForeign, bDiscountExists
    
    Set oSM = New z_StockManager
    Set oFSO = New FileSystemObject
    oSM.LogTransmission Me.TRID, "Printed: " & Format(Date, "dd/mm/yyyy") & vbCrLf
    Log = "File: " & oFSO.GetFileName(oPCL.FilePath) & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
    Set oSM = Nothing
    Set oFSO = Nothing
    
    oPCL.FinishRequest
        
EXIT_Handler:
    Exit Function
errHandler:
    ErrPreserve
    oPCL.FinishRequest
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.PrintPO(pShowVAT,bForeign,pRoundedUp)", Array(pShowVAT, bForeign, pRoundedUp)
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
    ErrorIn "a_PO.DisplayTotals(pLabels,pValues,bForeign)", Array(pLabels, pValues, bForeign)
End Function
Private Sub FillTableRows(pPO As a_PO, bForeign As Boolean, pDiscountExists As Boolean)
    On Error GoTo errHandler
Dim i As Integer

    For i = 1 To pPO.POLines.Count
        With pPO.POLines(i)
            If .Discount > 0 Then pDiscountExists = True
            oPCL.POLSend .ProductCodeF, .QtyFirmF, .QtySSF, .TitleAuthor, .PriceF(bForeign), .PLessDiscExtF(bForeign), .DiscountF, PackText(.Note), .Ref '"InvDate", oInv.TransDateFormatted
        End With
    Next i
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.FillTableRows(pPO,bForeign,pDiscountExists)", Array(pPO, bForeign, pDiscountExists)
End Sub


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
    ErrorIn "a_PO.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Missing supplier", "TP"
    colClassors.Add "Note too short", "NOTE"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.LoadClassorsCollection"
End Sub
Private Function CheckInput(pVal As String, pFld As String) As Boolean
    On Error GoTo errHandler

    CheckInput = True
    If mudtProps.TPID = 0 Then
        mobjValid.BreakRule "TP", True
        If pFld = "TP" Then CheckInput = False
    Else
        mobjValid.BreakRule "TP", False
    End If

    mobjValid.GetStatus
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CheckInput(pVal,pFld)", Array(pVal, pFld)
End Function


Public Function Post() As String
    On Error GoTo errHandler
Dim oSM As z_StockManager
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
   ' oPC.COShort.BeginTrans
    Me.ApplyEdit
    Set oSM = New z_StockManager
    
    Post = ""
    Select Case mOriginalStatus
    Case stVOID      'VOID
            Post = "VOID"
    Case stInProcess      'In Process
            Select Case Status
            Case stVOID
       ''         SetStatus stVOID
            Case stInProcess
                oSM.IssuePO Me, Status
                Post = "In Process"
            Case stCOMPLETE, stISSUED
                oSM.IssuePO Me, Status
              '  SetStatus stCOMPLETE
            Case stPROFORMA
                oSM.IssuePO Me, stISSUED
        '        SetStatus stPROFORMA
            End Select
    Case stISSUED
            Select Case Status
            Case stCANCELLED
                oSM.CancelPO Me
            Case stISSUED, stCOMPLETE
                oSM.IssuePO Me, stISSUED
            End Select
            Post = "ISSUED"
    Case stPROFORMA      'Issued PROFORMA
            Select Case Status
            Case stVOID
                oSM.CancelPO Me
        '        SetStatus stVOID
            Case stInProcess
                Post = "OR: was proforma, can't now be inprocess."
            Case stCOMPLETE, stISSUED
                oSM.IssuePO Me, Status
        '        SetStatus stCOMPLETE
            Case stPROFORMA
                Post = "PROFORMA"
            End Select
    End Select
    Set oSM = Nothing
  '  oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Post", , EA_RERAISE + EA_ADVANCED, oPC.COShort
End Function
Public Sub SetStatus(val As enStatus)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Status = val
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SetStatus(val)", val
End Sub

Public Property Get CanCancel() As Boolean
    On Error GoTo errHandler
Dim bCanCancel As Boolean
Dim oPOL As a_POL
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
    ErrorIn "a_PO.CanCancel"
End Property

Public Function VoidDocument()
    On Error GoTo errHandler
    Me.BeginEdit
    Me.SetStatus stVOID
    Me.ApplyEdit
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.VoidDocument"
End Function

Public Property Get ISForeignCurrency() As Boolean
    On Error GoTo errHandler
    ISForeignCurrency = (Not oPC.Configuration.DefaultCurrency Is Me.CaptureCurrency)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.ISForeignCurrency"
End Property

Public Property Get CurrencyConversionAsText() As String
    On Error GoTo errHandler
   ' CurrencyConversionAsText = mudtProps.CurrencyRate & " " & oPC.Configuration.DefaultCurrency.Description & " per " & Me.CaptureCurrency.Description
    CurrencyConversionAsText = mudtProps.CurrencyRate & " " & Me.CaptureCurrency.Description & " per " & oPC.Configuration.DefaultCurrency.Description
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CurrencyConversionAsText"
End Property

Public Function GenerateSAANAMsg() As Boolean
    On Error GoTo errHandler
Dim str As String
Dim strMsg As String
Dim dte As Date
Dim oTF As New z_TextFile
Dim strFilePath As String
Dim strFilename As String
Dim i As Integer
Dim oSM As z_StockManager
Dim oFSO As New FileSystemObject

    dte = Now()
    strFilename = oPC.Configuration.DefaultStore.code & "_" & Me.DOCCode
    
    strFilePath = oPC.SharedFolderRoot & "\EDI\POs\OUT\" & Me.Supplier.AcNo & "_" & oFSO.GetBaseName(strFilename) & ".EDI"
    
    oTF.OpenTextFile strFilePath
    strMsg = "UNA+1+" & Me.DOCCode & "+"
    oTF.WriteToTextFile_NoLineTerminator strMsg & Chr(10)
    
    strMsg = "UNB+UNOA:2+" & Me.Supplier.GFXNumber & "+" & Format(dte, "YYYYMMDD") & ":" & Format(dte, "HHNN") & "+" & Me.DOCCode & "+PASSWORD+ORDERS"
    oTF.WriteToTextFile_NoLineTerminator strMsg & "'" & Chr(10)
    strMsg = "UNH+" & Me.DOCCode & "+ORDERS:6:0:S"
    oTF.WriteToTextFile_NoLineTerminator strMsg & "'" & Chr(10)
    strMsg = "SOP+" & oPC.Configuration.GFXNumber
    oTF.WriteToTextFile_NoLineTerminator strMsg & "'" & Chr(10)
    strMsg = "CLO+" & oPC.Configuration.GFXNumber & "+" & oPC.Configuration.GFXNumber & "+++" & oPC.Configuration.GFXNumber & "+N"
    oTF.WriteToTextFile_NoLineTerminator strMsg & "'" & Chr(10)
    strMsg = "ORD+" & Me.DOCCode & ":" & Format(dte, "YYYYMMDD") & ":" & Format(DateAdd("m", 4, dte), "YYYYMMDD") & "+N"
    oTF.WriteToTextFile_NoLineTerminator strMsg & "'" & Chr(10)
    strMsg = "DIN+" & Format(DateAdd("d", 1, dte), "YYYYMMDD") & ":" & Format(DateAdd("m", 4, dte), "YYYYMMDD")
    oTF.WriteToTextFile_NoLineTerminator strMsg & "'" & Chr(10)
    
    For i = 1 To Me.POLines.Count
        If Me.POLines(i).Fulfilled <> "CAN" Then
            If Len(Me.POLines(i).EAN) = 13 Then
                strMsg = "OLD+" & CStr(i) & "+:F:" & Me.POLines(i).EAN & ":+" & CStr(Me.POLines(i).QtyFirm + Me.POLines(i).QtySS) & "+" & CStr((POLines(i).Price(False) / oPC.Configuration.DefaultCurrency.Divisor)) & "++" & CStr((POLines(i).PExt(False) / oPC.Configuration.DefaultCurrency.Divisor)) & "+++0.00"
                oTF.WriteToTextFile_NoLineTerminator strMsg & "'" & Chr(10)
            ElseIf Len(Me.POLines(i).ProductCode) = 10 Then
                strMsg = "OLD+" & CStr(i) & "+:F:" & Me.POLines(i).ProductCode & ":+" & CStr(Me.POLines(i).QtyFirm + Me.POLines(i).QtySS) & "+" & CStr((POLines(i).Price(False) / oPC.Configuration.DefaultCurrency.Divisor)) & "++" & CStr((POLines(i).PExt(False) / oPC.Configuration.DefaultCurrency.Divisor)) & "+++0.00"
                oTF.WriteToTextFile_NoLineTerminator strMsg & "'" & Chr(10)
            End If
        End If
    Next
    strMsg = "UNT+" & CStr(6 + Me.POLines.Count) & "+" & Me.DOCCode   'Assuming DIN is not needed
    oTF.WriteToTextFile_NoLineTerminator strMsg & "'" & Chr(10)
    strMsg = "UNZ+1+" & Me.DOCCode
    oTF.WriteToTextFile_NoLineTerminator strMsg & "'" & Chr(10)
    
    Set oSM = New z_StockManager
    oSM.LogTransmission Me.TRID, "File: " & oFSO.GetFileName(oTF.FileName) & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf
    Log = "File: " & oFSO.GetFileName(oTF.FileName) & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
    Set oSM = Nothing
    
    oTF.CloseTextFile
    Exit Function
errHandler:
    ErrorIn "a_PO.GenerateSAANAMsg"
End Function
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

Private Function ValidateObject(pVal As String, pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean
    bValid = True
    If mudtProps.TPID = 0 Then
        mobjValid.BreakRule "TP", True
        If pFld = "TP" Then bValid = False
    Else
        mobjValid.BreakRule "TP", False
    End If

    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.ValidateObject(pVal,pFld)", Array(pVal, pFld)
End Function

Public Sub PasteLine(PID As String, QtyFirm As Long, QtySS As Long, Price As Long, DiscountRate As Double, VATRate As Double, Ref As String, pETA As Date)
Dim oSM As New z_StockManager

    oSM.PastePOLine PID, QtyFirm, QtySS, Price, DiscountRate, VATRate, mudtProps.TRID, Ref, pETA
    
End Sub


Public Function ExportToXML_MEC(bForeign As Boolean, DispatchMode As enTransmitType, pFilename As String) As Boolean
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim strPath As String
Dim strBillto As String
Dim strDelto As String
Dim strFOFile As String
Dim strPDFFile As String
Dim strEDIFilePath As String
Dim strXML As String
Dim strCommand As String
Dim sFormatFilename As String
Dim i As Integer
Dim strHTML As String
Dim fs As New FileSystemObject
Dim objXSL As New MSXML2.DOMDocument60
Dim opXMLDOC As New MSXML2.DOMDocument60
Dim objXMLDOC  As New MSXML2.DOMDocument60
Dim strWorkingFolder As String
Dim strDispatchMethod As String
Dim oDC As a_DocumentControl
Dim strExecutable As String
Dim cnt As Integer
Dim lngCnt As Long
Dim dteTransmission As Date
Dim oPOL As a_POL
Dim j As Integer
Dim iter As Integer
Dim iSegCnt As Integer
Dim sDocRef As String
Dim sAddress As String

    sDocRef = "12345"
    dteTransmission = Now()
    strWorkingFolder = oPC.LocalFolder & "TEMP\"
    strDispatchMethod = "EDI"
                        p 1
                        p 2

    Set xMLDoc = New ujXML
                        p 21
                        
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "ORDERS"
            .chCreate "UNB"
            .chCreate "S001"
            .chCreate "D0001"
                .elText = "UNOC"
            .elCreateSibling "D0002"
                .elText = "3"
            .navUP
            .elCreateSibling "S002"
            .chCreate "D0004"
                .elText = oPC.Configuration.GFXNumber
            .elCreateSibling "D0007"
                .elText = "ZZ"
            .navUP
            .elCreateSibling "S003"
          '  .chCreate "S0010"
           '     .elText = Me.Supplier.GFXNumber
            .chCreate "D0007"
                .elText = "ZZ"
            .navUP
            .navUP
            
            .chCreate "S004"
            .chCreate "D0017"
                .elText = Format(dteTransmission, "YYMMDD")
            .elCreateSibling "D0019"
                .elText = Format(dteTransmission, "HHNN")
            .navUP
            .elCreateSibling "D0020"
                .elText = Me.DOCCode
            .navUP
          '  .navUP
            .elCreateSibling "SG0"
            .chCreate "UNH"
                        iSegCnt = 1
            .chCreate "D0062"
                .elText = sDocRef
            .elCreateSibling "S009"
            .chCreate "D0065"
                .elText = "ORDERS"
            .elCreateSibling "D0052"
                .elText = "D"
            .elCreateSibling "D0054"
                .elText = "96A"
            .elCreateSibling "D0051"
                .elText = "UN"
            .elCreateSibling "D0057"
                .elText = "EAN008"
                
            .navUP
            .navUP
           ' .navUP
            .elCreateSibling "BGM"
                        iSegCnt = iSegCnt + 1
            .chCreate "C002"
            .chCreate "D1001"
                .elText = "220"
            .navUP
            .elCreateSibling "D1004"
                .elText = Me.TRID
            .elCreateSibling "D1225"
                .elText = "9"
            .navUP
            .elCreateSibling "DTM"
                        iSegCnt = iSegCnt + 1
            .chCreate "C507"
            .elCreateSibling "D2005"
                .elText = "137"
            .elCreateSibling "D2380"
                .elText = Format(dteTransmission, "YYYYMMNN")
            .elCreateSibling "D2379"
                .elText = "102"
            .navUP
           ' .navUP
            .elCreateSibling "SG2"
                        iSegCnt = iSegCnt + 1
            .chCreate "NAD"
            .chCreate "D3035"
                .elText = "BY"
            .elCreateSibling "C082"
            .chCreate "D3039"
                .elText = oPC.Configuration.GFXNumber
            .elCreateSibling "D3055"
                .elText = "31B"
            .navUP
            .navUP
            .navUP
            .elCreateSibling "SG2"
                        iSegCnt = iSegCnt + 1
            .chCreate "NAD"
            .chCreate "D3035"
                .elText = "SU"
            .elCreateSibling "C082"
            .chCreate "D3039"
                .elText = Me.Supplier.GFXNumber
            .elCreateSibling "D3055"
                .elText = "31B"
            .navUP
            .navUP
            .navUP
            lngCnt = 1
            For Each oPOL In Me.POLines
                .elCreateSibling "SG25"
                .chCreate "LIN"
                        iSegCnt = iSegCnt + 1
                .chCreate "D1082"
                    .elText = lngCnt
                .elCreateSibling "C212"
                .chCreate "D7140"
                    .elText = oPOL.EAN
                .elCreateSibling "D7143"
                    .elText = "EN"
                .navUP
                .navUP
                    
                If oPOL.MainAuthor > "" Then
                    .elCreateSibling "IMD"
                            iSegCnt = iSegCnt + 1
                    .chCreate "D7077"
                    .elText = "F"
                    .elCreateSibling "D7081"
                        .elText = "BAU"
                    .elCreateSibling "C273"
                    .chCreate "D7008"
                        .elText = Mid(oPOL.MainAuthor, 1, 35)
                    j = 36
                    If j < Len(oPOL.MainAuthor) Then
                        .elCreateSibling "D7008"
                        .elText = Mid(oPOL.MainAuthor, j, 35)
                    End If
                    .navUP
                    .navUP
                End If
                
                
                 If oPOL.Title > "" Then
                    .elCreateSibling "IMD"
                            iSegCnt = iSegCnt + 1
                    .chCreate "D7077"
                    .elText = "F"
                    .elCreateSibling "D7081"
                        .elText = "BTI"
                    .elCreateSibling "C273"
                    .chCreate "D7008"
                        .elText = Mid(oPOL.Title, 1, 35)
                    j = 36
                    If j < Len(oPOL.Title) Then
                        .elCreateSibling "D7008"
                        .elText = Mid(oPOL.Title, j, 35)
                    End If
                    .navUP
                    .navUP
                End If
               
                
                  If oPOL.Publisher > "" Then
                    .elCreateSibling "IMD"
                            iSegCnt = iSegCnt + 1
                    .chCreate "D7077"
                    .elText = "F"
                    .elCreateSibling "D7081"
                        .elText = "BPU"
                    .elCreateSibling "C273"
                    .chCreate "D7008"
                        .elText = Mid(oPOL.Publisher, 1, 35)
                    j = 36
                    If j < Len(oPOL.Publisher) Then
                        .elCreateSibling "D7008"
                        .elText = Mid(oPOL.Publisher, j, 35)
                    End If
                    .navUP
                    .navUP
                End If
               
                
                  If oPOL.PublicationDate > "" Then
                    .elCreateSibling "IMD"
                            iSegCnt = iSegCnt + 1
                    .chCreate "D7077"
                    .elText = "F"
                    .elCreateSibling "D7081"
                        .elText = "BPD"
                    .elCreateSibling "C273"
                    .chCreate "D7008"
                        .elText = Mid(oPOL.PublicationDate, 1, 35)
                    j = 36
                    If j < Len(oPOL.PublicationDate) Then
                        .elCreateSibling "D7008"
                        .elText = Mid(oPOL.PublicationDate, j, 35)
                    End If
                .navUP
                .navUP
                End If
                
                  If oPOL.Edition > "" Then
                    .elCreateSibling "IMD"
                            iSegCnt = iSegCnt + 1
                    .chCreate "D7077"
                    .elText = "F"
                    .elCreateSibling "D7081"
                        .elText = "BEN"
                    .elCreateSibling "C273"
                    .chCreate "D7008"
                        .elText = Mid(oPOL.Edition, 1, 35)
                    j = 36
                    If j < Len(oPOL.Edition) Then
                        .elCreateSibling "D7008"
                        .elText = Mid(oPOL.Edition, j, 35)
                    End If
                    .navUP
                    .navUP
                End If
                
                
                .elCreateSibling "QTY"
                            iSegCnt = iSegCnt + 1
                .chCreate "C186"
                .chCreate "D6063"
                    .elText = "21"
                .elCreateSibling "D6060"
                    .elText = oPOL.QtyFirm
                
                .navUP
                .navUP
                
                   If oPOL.Note > "" Then
                    .elCreateSibling "FTX"
                            iSegCnt = iSegCnt + 1
                    .chCreate "D4451"
                        .elText = "LIN"
               '     .elCreateSibling "D7081"
               '         .elText = "BEN"
                    .elCreateSibling "C107"
                    .elCreateSibling "C108"
                    .chCreate "D4440"
                        .elText = Mid(oPOL.Note, 1, 70)
                    j = 70
                    iter = 0
                    Do While j < Len(oPOL.Note) And iter < 5
                        .elCreateSibling "D4440"
                        .elText = Mid(oPOL.Note, j, 70)
                        j = j + 70
                        iter = iter + 1
                    Loop
                .navUP
                .navUP
                End If
               
                .elCreateSibling "SG27"
                            iSegCnt = iSegCnt + 1
                .chCreate "PRI"
                .chCreate "C509"
                .chCreate "D5125"
                    .elText = "AAA"
                .elCreateSibling "D5118"
                    .elText = Format(oPOL.Price(False) / 100, "0.00")
                    
                .navUP
                .navUP
                .navUP
                If oPOL.Ref > "" Then
                    .elCreateSibling "SG28"
                            iSegCnt = iSegCnt + 1
                    .chCreate "RFF"
                    .chCreate "C506"
                    .chCreate "D1153"
                        .elText = "LI"
                    .elCreateSibling "D1154"
                        .elText = oPOL.Ref
                    .navUP
                    .navUP
                    .navUP
                End If
               
                .elCreateSibling "SG34"
                            iSegCnt = iSegCnt + 1
                .chCreate "NAD"
                .chCreate "D3035"
                   .elText = "DP"
                .elCreateSibling "C082"
                .chCreate "D3039"
                    .elText = oPC.Configuration.GFXNumber
                .elCreateSibling "D3055"
                    .elText = "31B"
                .navUP
                .elCreateSibling "C058"
                .chCreate "D3124"
                    sAddress = stripCRLF(oPC.Configuration.DefaultStore.DelAddress, " ")
                    .elText = Mid(sAddress, 1, 35)
                j = 35
                iter = 0
                Do While j < Len(sAddress) And iter < 5
                    .elCreateSibling "D3124"
                    .elText = Mid(sAddress, j, 35)
                    j = j + 35
                    iter = iter + 1
                Loop
                
                .navUP
                .navUP
                .navUP
               
                
                .elCreateSibling "SG34"
                            iSegCnt = iSegCnt + 1
                .chCreate "NAD"
                .chCreate "D3035"
                   .elText = "OB"
                .elCreateSibling "C082"
                .chCreate "D3039"
                    .elText = oPC.Configuration.GFXNumber
                .elCreateSibling "D3055"
                    .elText = "31B"
                .navUP
                .elCreateSibling "C058"
                .chCreate "D3124"
                    sAddress = stripCRLF(oPC.Configuration.DefaultStore.DelAddress, " ")
                    .elText = Mid(sAddress, 1, 35)
                j = 35
                iter = 0
                Do While j < Len(sAddress) And iter < 5
                    .elCreateSibling "D3124"
                    .elText = Mid(sAddress, j, 35)
                    j = j + 35
                    iter = iter + 1
                Loop
                
                .navUP
                .navUP
                .navUP
                .navUP
              
                lngCnt = lngCnt + 1
            Next
                
            .elCreateSibling "UNS"
                            iSegCnt = iSegCnt + 1
            .chCreate "D0081"
                .elText = "S"
            .navUP
            .elCreateSibling "CNT"
                            iSegCnt = iSegCnt + 1
            .chCreate "C270"
            .chCreate "D6069"
                .elText = "1"
            .elCreateSibling "D6066"
                .elText = Me.TotalQty
            .navUP
            .navUP
            .elCreateSibling "CNT"
                            iSegCnt = iSegCnt + 1
            .chCreate "C270"
            .chCreate "D6069"
                .elText = "2"
            .elCreateSibling "D6066"
                .elText = Me.POLines.Count
            .navUP
            .navUP
            .elCreateSibling "UNT"
                            iSegCnt = iSegCnt + 1
            .chCreate "D0074"
                .elText = CStr(iSegCnt)
            .elCreateSibling "D0062"
                .elText = sDocRef
            .navUP
            .navUP
            .elCreateSibling "UNZ"
            .chCreate "D0036"
                .elText = "1"
            .elCreateSibling "D0020"
                .elText = Me.DOCCode
    End With
                            p 26
'FINALLY PRODUCE THE .XML FILE
    strXML = strWorkingFolder & "PO_" & Me.DOCCode & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UTF-8", "<?xml version=""1.0"" encoding=""UTF-8"" ?>", False
    End With
    sFormatFilename = oPC.SharedFolderRoot & "\MEC\formatdescription\edifact.96.a.orders.xml"
                            p 27
    strEDIFilePath = oPC.SharedFolderRoot & "\EDI_Out\" & fs.GetBaseName(strXML) & ".EDIFACT"
    strCommand = "jstart.BAT de.mendelson.eagle.converter.xmledi.XMLEDIConverter -xmlin " & strXML & " -formatout " & sFormatFilename & " -ediout " & strEDIFilePath & "  -ncs"
    ChDir "\PBKS\MEC"
    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 600000
    
''Log the dispatch
'    Set oSM = New z_StockManager
'    Set oFSO = New FileSystemObject
'    pFilename = strPDFFile
'
'    If DispatchMode <> enMail And DispatchMode <> enView Then
'        oSM.LogTransmission Me.TRID, "Dispatched: by " & strDispatchMethod & " : " & Format(Date, "dd/mm/yyyy") & vbCrLf
'        Log = "Dispatched: by " & strDispatchMethod & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
'    End If
'    Set oSM = Nothing
'    Set oFSO = Nothing
'                        p 9
'    oTF.CloseTextFile
'    ExportToXML = True
'
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.ExportToXML(bForeign)", bForeign, , , "strPos", Array(strErrPos)
End Function

