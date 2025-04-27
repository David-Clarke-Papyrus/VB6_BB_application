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
Attribute VB_Exposed = True
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
Private colClassErrors As Collection
'Private oBillToAddress As a_Address
'Private oDelToAddress As a_Address
Private oBillingCompany As a_Company
'Private CaptureCurrency As a_Currency
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private moCaptureInCurrency As a_Currency 'usually this will be the system default currency but it could be a foreign currency
' if we are ordering from a foreign company and wish to reflect our understanding of their prices in their currency.
' The currency used and the conversionrate as at time of issuing PO are stored with the PO.
' We can thus record differnces between the ordered and delivered prices which reflect the currency flustuations.
Const CDOCCODE = "PO"

Private Sub mPOL_CurrRowStatus(pMsg As String)
    On Error GoTo ErrHandler
    RaiseEvent CurrRowStatus(pMsg)
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.mPOL_CurrRowStatus(pMsg)", pMsg
End Sub
Public Property Get CaptureCurrency() As a_Currency
    Set CaptureCurrency = moCaptureInCurrency
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
    On Error GoTo ErrHandler
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
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.LevelNewValues(pPOL)", pPOL
End Function

Public Function CalculateTotals()
    On Error GoTo ErrHandler
    CalculateTotal_Local
    CalculateTotal_Foreign
    RaiseEvent TotalChange(Me.TotalLessDiscExtF(False), Me.TotalLessDiscExtF(True))
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CalculateTotals"
End Function
Public Function CalculateTotal_Local()
    On Error GoTo ErrHandler
Dim ol As a_POL
Dim bLineDiscountApplied As Boolean
    bLineDiscountApplied = False
    lngLPLessDiscExt_tot = 0
    lngLPExt_tot = 0
    lngVAT_tot = 0
    lngDisc_tot = 0
    'NOTE Some items on the invoice might be nonstock e.g. Insurance, Postage and so on. These items are non subject to any discount
    'and they may mut probably will not carry VAT. e.g. a handling charge would include VAT
    'VAT is calculated on the discounted price, so it is necessary to apply the discount before calculating VAT
    'Both forms of discount (by line or overall(mutually exclusive)) will need to be applied line by line, and the VAT for that line
    'calculated afterwards. Remember VAT rate can vary per line.
    'Any general discount will merely be distributed as a line by line discount on all non nonstock items
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
ErrHandler:
    ErrPreserve
    If Err = 6 Then 'overflow
        Err.Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CalculateTotal_Local"
End Function
Public Function CalculateTotal_Foreign()
    On Error GoTo ErrHandler
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
ErrHandler:
    ErrPreserve
    If Err = 6 Then 'overflow
        Err.Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CalculateTotal_Foreign"
End Function
'Public Property Get TotalInvoiceDiscountAmountF() As Long
'    TotalInvoiceDiscountAmountF = Format(lngTotalInvoiceDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
Public Property Get TotalLessDiscExtF(bForeign As Boolean) As String
    On Error GoTo ErrHandler
    If bForeign Then
        TotalLessDiscExtF = Format(mudtProps.TotalExtension_Foreign / CaptureCurrency.Divisor, CaptureCurrency.FormatString)
    Else
        TotalLessDiscExtF = Format(mudtProps.TotalExtension / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TotalLessDiscExtF(bForeign)", bForeign
End Property
Public Property Get TotalExtF(bForeign As Boolean) As String
    On Error GoTo ErrHandler
    If bForeign Then
        TotalExtF = Format(mudtProps.TotalExtensionSimple_Foreign / CaptureCurrency.Divisor, CaptureCurrency.FormatString)
    Else
        TotalExtF = Format(mudtProps.TotalExtensionSimple / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TotalExtF(bForeign)", bForeign
End Property

Public Property Get TotalLessDiscExt(bForeign As Boolean) As String
    On Error GoTo ErrHandler
    If bForeign Then
        TotalLessDiscExt = mudtProps.TotalExtension_Foreign
    Else
        TotalLessDiscExt = mudtProps.TotalExtension
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TotalLessDiscExt(bForeign)", bForeign
End Property
Public Property Get TotalVATF(bForeign As Boolean) As String
    On Error GoTo ErrHandler
    If bForeign Then
        TotalVATF = Format(mudtProps.TotalVAT_Foreign / CaptureCurrency.Divisor, CaptureCurrency.FormatString)
    Else
        TotalVATF = Format(mudtProps.TotalVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TotalVATF(bForeign)", bForeign
End Property
Public Property Get TotalPayable(bForeign As Boolean) As String
    On Error GoTo ErrHandler
    If bForeign Then
        TotalPayable = mudtProps.TotalPayable_Foreign / CaptureCurrency.Divisor
    Else
        TotalPayable = mudtProps.TotalPayable / oPC.Configuration.DefaultCurrency.Divisor
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TotalPayable(bForeign)", bForeign
End Property
Public Property Get TotalPayableF(bForeign As Boolean) As String
    On Error GoTo ErrHandler
    If bForeign Then
        TotalPayableF = Format(TotalPayable(bForeign), CaptureCurrency.FormatString)
    Else
        TotalPayableF = Format(TotalPayable(bForeign), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TotalPayableF(bForeign)", bForeign
End Property
Public Property Get TotalDiscount(bForeign As Boolean) As Long
    On Error GoTo ErrHandler
    If bForeign Then
        TotalDiscount = mudtProps.TotalDiscount_Foreign
    Else
        TotalDiscount = mudtProps.TotalDiscount
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TotalDiscount(bForeign)", bForeign
End Property
Public Property Get TotalDiscountF(bForeign As Boolean) As String
    On Error GoTo ErrHandler
    If bForeign Then
        TotalDiscountF = Format(mudtProps.TotalDiscount_Foreign / CaptureCurrency.Divisor, CaptureCurrency.FormatString)
    Else
        TotalDiscountF = Format(mudtProps.TotalDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TotalDiscountF(bForeign)", bForeign
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
    On Error GoTo ErrHandler
  Dim udtData As POData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.GetState"
End Function
Private Sub SetState(buffer As String)
    On Error GoTo ErrHandler
  Dim udtData As POData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SetState(buffer)", buffer
End Sub
Public Function GetSuperState() As String
    On Error GoTo ErrHandler
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
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.GetSuperState"
End Function
Public Sub SetSuperState(buffer As String)
    On Error GoTo ErrHandler
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
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SetSuperState(buffer)", buffer
End Sub


Public Sub BeginEdit()
    On Error GoTo ErrHandler
    mPOL.BeginEdit
    mcolStack.Add GetState
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo ErrHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mPOL.CancelEdit
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo ErrHandler
Dim objPersist As a_PO_P
'    CalculateTotal
    mudtProps.TPID = Me.Supplier.ID
    'Set these values fromthe objects before saving -  the objects cant be saved
    mudtProps.DocStoreID = oPC.Configuration.DefaultStoreID
   ' mudtProps.DELTOStoreID = o
        
    If DOCCode = "" Then
        DOCCode = FetchSONumber("GETANDWRITE") & IIf(ContainsCO = True, "X", "")
        DOCDate = Now()
    End If
    If CaptureDate < #1/1/1995# And mudtProps.IsNew Then CaptureDate = Now
    
    If mcolStack.Count = 0 Then Err.Raise 445
    mPOL.ApplyEdit
    Set objPersist = New a_PO_P
    If mudtProps.IsDeleted Then
        objPersist.DeletePO mudtProps.TRID
        mcolStack.Remove mcolStack.Count
        mudtProps.IsNew = True
        mudtProps.IsDeleted = False
    ElseIf IsDirty Or mudtProps.IsNew Then
        If mudtProps.IsNew Then mudtProps.CaptureDate = Now()
        If Not IsValid Then Err.Raise 445
        mcolStack.Remove mcolStack.Count
        SetSuperState objPersist.Save(GetSuperState)
        mudtProps.IsNew = False
    Else
        mcolStack.Remove mcolStack.Count
    End If
    Set objPersist = Nothing
    mudtProps.IsDirty = False
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.ApplyEdit"
End Sub
Public Property Get IsDeleted() As Boolean
    On Error GoTo ErrHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo ErrHandler
  IsEditing = mcolStack.Count > 0
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo ErrHandler
  IsNew = mudtProps.IsNew
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo ErrHandler
  IsDirty = mudtProps.IsDirty Or mPOL.IsDirty
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.IsDirty"
End Property

Public Property Get IsValid() As Boolean
    On Error GoTo ErrHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.IsValid"
End Property
Public Function GetSTatus()
    On Error GoTo ErrHandler
    If IsValid Then
        RaiseEvent Valid("")
    Else
        RaiseEvent Valid(TranslateErrors(mobjValid.AllBrokenRules))
    End If
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.GetSTatus"
End Function
Private Sub Class_Initialize()
    On Error GoTo ErrHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    Set mPOL = New ch_POL
    mPOL.Component Me
    Set mSupplier = New a_Supplier
    Set oPCL = New z_PrintClient
    mudtProps.IsNew = True
    mudtProps.Status = 2
    mOriginalStatus = stInProcess
'    Set oBillToAddress = oPC.Configuration.DefaultStore.BillAddress
    
    mudtProps.DocStoreID = oPC.Configuration.DefaultStoreID
    mobjValid.BreakRule "TP", True
    LoadClassErrorsCollection
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrHandler
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
    Set colClassErrors = Nothing
    Set mcolStack = Nothing
    Set mobjValid = Nothing
    Set mPOL = Nothing
    Set mSupplier = Nothing
    Set oPCL = Nothing
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(errors As String)
    On Error GoTo ErrHandler
    RaiseEvent Valid(TranslateErrors(errors))
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.mobjValid_BrokenRule(errors)", errors
End Sub
Private Sub mobjValid_RuleUnbroken(errors As String)
    On Error GoTo ErrHandler
    RaiseEvent Valid(TranslateErrors(errors))
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.mobjValid_RuleUnbroken(errors)", errors
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo ErrHandler
    RaiseEvent Valid("")
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo ErrHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.mobjValid_Status(pMsg)", pMsg
End Sub

Public Sub Load(TRID As Long, ReadOnly As Boolean)
    On Error GoTo ErrHandler
Dim objPersist As a_PO_P
Dim oADd As a_Address
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
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Load(TRID,ReadOnly)", Array(TRID, ReadOnly)
End Sub

Public Sub Reload()
    On Error GoTo ErrHandler
Dim lngTRID As Long
    lngTRID = Me.TRID
    Class_Terminate
    Class_Initialize
    Load lngTRID, True
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Reload"
End Sub
Public Property Get Supplier() As a_Supplier
    On Error GoTo ErrHandler
    Set Supplier = mSupplier
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Supplier"
End Property
Public Sub Delete()
    On Error GoTo ErrHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  SetDirty True
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Delete"
End Sub



Public Property Get SOID() As Long
    On Error GoTo ErrHandler
    SOID = mudtProps.TRID
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SOID"
End Property
Public Property Let SOID(val As Long)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TRID = val
    SetDirty True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SOID(val)", val
End Property

Public Property Get BillToAddressID() As Long
    On Error GoTo ErrHandler
    BillToAddressID = mudtProps.DocStoreID
    Exit Property
ErrHandler:
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
    On Error GoTo ErrHandler
    DELTOStoreID = mudtProps.DELTOStoreID
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DELTOStoreID"
End Property
Public Sub setDelToStoreID(val As Long)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DELTOStoreID = val
    mudtProps.DelToAddress = oPC.Configuration.Stores.FindStoreByID(val).DelAddress
    SetDirty True
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.setDelToStoreID(val)", val
End Sub
Private Sub SetDirty(pVal As Boolean)
    On Error GoTo ErrHandler
    mudtProps.IsDirty = pVal
    RaiseEvent Dirty(pVal)
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SetDirty(pVal)", pVal
End Sub

Public Property Get TRID() As Long
    On Error GoTo ErrHandler
    TRID = mudtProps.TRID
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TRID"
End Property
Public Property Let TRID(val As Long)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TRID = val
    SetDirty True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TRID(val)", val
End Property
Public Property Let ContainsCO(p As Boolean)
    On Error GoTo ErrHandler
    mudtProps.ContainsCO = p
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.ContainsCO(p)", p
End Property
Public Property Get ContainsCO() As Boolean
    On Error GoTo ErrHandler
    ContainsCO = mudtProps.ContainsCO
    Exit Property
ErrHandler:
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
'        If Not oPO.NONStock Then
'            oPO.Discount = dblDiscount
'        End If
'    Next
'    mPOL_RowsChange  'recalulates andd raises event
'    RaiseEvent reloadlist
'End Function
Public Function SetTP(pTPID As Long) As Boolean
    On Error GoTo ErrHandler
Dim bSuccess As Boolean
    bSuccess = mSupplier.Load(pTPID)
    SetCaptureCurrency Supplier.DefaultCurrency
    SetTP = bSuccess
    If bSuccess Then
        mobjValid.RuleBroken "TP", False
    End If
    SetDirty True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SetTP(pTPID)", pTPID
End Function
Public Function SetSupplierFromAccNum(pAccnum As String) As Boolean
    On Error GoTo ErrHandler
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
    SetDirty True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SetSupplierFromAccNum(pAccnum)", pAccnum
End Function
Public Function SetCaptureCurrency(val As a_Currency) As Boolean
    On Error GoTo ErrHandler
    Set moCaptureInCurrency = val
    mudtProps.CurrencyID = val.ID
    mudtProps.CurrencyRate = val.Factor
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SetCaptureCurrency(val)", val
End Function
Public Property Get StaffID() As Long
    On Error GoTo ErrHandler
    StaffID = mudtProps.StaffID
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.StaffID"
End Property
Public Property Let StaffID(val As Long)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.StaffID = val
    SetDirty True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.StaffID(val)", val
End Property
Public Property Get StaffName() As String
    On Error GoTo ErrHandler
    If mudtProps.StaffID < 1 Then
        StaffName = ""
        Exit Property
    End If
    StaffName = oPC.Configuration.Staff.FindStaffByID(StaffID).StaffName
    Exit Property
ErrHandler:
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
    On Error GoTo ErrHandler
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
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.StaffNameB"
End Property

Public Property Get BillingCompany() As a_Company
    Set BillingCompany = oBillingCompany
End Property
Public Property Get TPName() As String
    On Error GoTo ErrHandler
    TPName = FNS(mudtProps.TPName)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TPName"
End Property
Public Property Let TPName(val As String)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPName = val
    SetDirty True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TPName(val)", val
End Property

Public Property Get TPPhone() As String
    On Error GoTo ErrHandler
    TPPhone = FNS(mudtProps.TPPhone)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TPPhone"
End Property
Public Property Let TPPhone(val As String)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPPhone = val
    SetDirty True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TPPhone(val)", val
End Property

Public Property Get TPFax() As String
    On Error GoTo ErrHandler
    TPFax = FNS(mudtProps.TPFax)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TPFax"
End Property
Public Property Let TPFax(val As String)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPFax = val
    SetDirty True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TPFax(val)", val
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
    On Error GoTo ErrHandler
    Memo = FNS(mudtProps.Memo)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Memo"
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
    On Error GoTo ErrHandler
    OrderType = FNS(mudtProps.OrderType)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.OrderType"
End Property
Public Property Let OrderType(val As String)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.OrderType = val
    SetDirty True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.OrderType(val)", val
End Property

Public Property Get DOCCode() As String
    On Error GoTo ErrHandler
    DOCCode = FNS(mudtProps.DOCCode)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DocCode"
End Property
Public Property Let DOCCode(val As String)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DOCCode = val
    SetDirty True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DocCode(val)", val
End Property
Public Property Get TPACCNum() As String
    On Error GoTo ErrHandler
    TPACCNum = FNS(mudtProps.TPACCNum)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TPACCNum"
End Property
Public Property Let TPACCNum(val As String)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPACCNum = val
    SetDirty True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TPACCNum(val)", val
End Property
Public Property Get DeliverToAddress() As String
    On Error GoTo ErrHandler
    DeliverToAddress = mudtProps.DelToAddress
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DeliverToAddress"
End Property
Public Property Let DeliverToAddress(val As String)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DelToAddress = val
    SetDirty True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DeliverToAddress(val)", val
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
    On Error GoTo ErrHandler
    DOCDate = mudtProps.DOCDate
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DocDate"
End Property
Public Property Get TransDateFormatted() As String
    On Error GoTo ErrHandler
    TransDateFormatted = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TransDateFormatted"
End Property
Public Property Let DOCDate(val As Date)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    If IsDate(val) Then
        mudtProps.DOCDate = CDate(val)
        SetDirty True
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DocDate(val)", val
End Property

Public Property Get CaptureDate() As Date
    On Error GoTo ErrHandler
    CaptureDate = mudtProps.CaptureDate
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CaptureDate"
End Property
Public Property Get CaptureDateFormatted() As Date
    On Error GoTo ErrHandler
    CaptureDateFormatted = Format(mudtProps.CaptureDate, "dd/mm/yyyy")
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CaptureDateFormatted"
End Property
Public Property Let CaptureDate(val As Date)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CaptureDate = val
    SetDirty True
    Exit Property
ErrHandler:
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
    On Error GoTo ErrHandler
    OrderDate = mudtProps.DOCDate
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.OrderDate"
End Property
Public Property Get OrderDateFormatted() As String
    On Error GoTo ErrHandler
    OrderDateFormatted = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.OrderDateFormatted"
End Property
Public Property Let OrderDate(val As Date)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DOCDate = val
    SetDirty True
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.OrderDate(val)", val
End Property

'Public Property Get Vatable() As Boolean
'  '  Vatable = mudtProps.Vatable
'  Vatable = Me.Customer.Vatable
'End Property
'Public Property Let Vatable(val As Boolean)
'    mudtProps.Vatable = val
'    SetDirty True
'End Property


Public Property Get POLines() As ch_POL
  Set POLines = mPOL
End Property
Public Property Get DisplayPO() As d_PO
    On Error GoTo ErrHandler
    Set DisplayPO = New d_PO
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DisplayPO"
End Property
Public Property Get StatusF() As String
    On Error GoTo ErrHandler
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
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.StatusF"
End Property
Public Property Get StatusForPrinting() As String
    On Error GoTo ErrHandler
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
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.StatusForPrinting"
End Property
Public Property Get Status() As enStatus
    On Error GoTo ErrHandler
    Status = mudtProps.Status
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Status"
End Property
Public Function FetchSONumber(strTmp As String) As String
    On Error GoTo ErrHandler
Dim oCode As z_Code
Dim strDefaultCo As String
Dim strPONumber As String
    Set oCode = New z_Code
    
    strPONumber = oCode.GetNextCode(enSupplierOrder, strTmp)
    
    strDefaultCo = oPC.Configuration.DefaultCompany.CompanyCode
    FetchSONumber = strDefaultCo & "P" & strPONumber
    
    Set oCode = Nothing
    Exit Function
ErrHandler:
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
    On Error GoTo ErrHandler
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
'ERR_Handler:
'    oPCL.FinishRequest
'    Select Case Err
'    Case 5941
'        MsgBox "Bookmark missing"
'        Resume Next
'    Case Else
'        MsgBox Error
'        GoTo EXIT_Handler
'        Resume
'    End Select
'
    Exit Function
ErrHandler:
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
Public Function ExportToXML(bForeign As Boolean, DispatchMode As enTransmitType, Optional pQtyCopies As Integer) As Boolean
    On Error GoTo ErrHandler
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
Dim objXSL As New MSXML2.DOMDocument30
Dim opXMLDOC As New MSXML2.DOMDocument30
Dim objXMLDOC  As New MSXML2.DOMDocument30
Dim strWorkingFolder As String
Dim strDispatchMethod As String
Dim oDC As a_DocumentControl
'Const strHead = "<?xml version=""1.0"" encoding=""UTF-8""?>" _
'            & "<PO_1 xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation=""C:\PBKS\Templates\PO_FO_1.xsd"">"

    If DispatchMode = enMail Then
        strWorkingFolder = oPC.SharedFolderRoot & "\Emails\"
        strDispatchMethod = "Emailing"
    ElseIf DispatchMode = enPrint Then
        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
        strDispatchMethod = "Printing"
    End If

    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing And strDispatchMethod <> "Emailing" Then
        ExportToXML = False
        Exit Function
    End If
    If pQtyCopies < 1 Then pQtyCopies = 1

    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "PO_DOC"
            .chCreate "DestinationAddress"
                If Me.Supplier Is Nothing Then
                    .elText = ""
                Else
                    .elText = Me.Supplier.OrderToAddress.Email
                End If
'            If bDiscountExists Then
'                    .elText = "PO_DISC_REF_FO.XSL"
'            ElseIf bReferenceExists Then
'                    .elText = "PO_REF_FO.XSL"
'            Else
            .elCreateSibling "TemplateName"
                .elText = "PO_DOC_FO.XSL"
            .elCreateSibling "SendersEmail"
                .elText = Me.SendersEmail
            .elCreateSibling "CopyCount"
                .elText = pQtyCopies
            .elCreateSibling "Printer"
                If Not oDC Is Nothing Then .elText = oDC.PrinterName
            .elCreateSibling "Status"
                .elText = StatusForPrinting
            
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
                .elText = Replace(oPC.Configuration.Stores.FindStoreByID(Me.BillToAddressID).BillAddress, Chr(13) & Chr(10), Chr(10))
            .elCreateSibling "DelTo", True
                .elText = Replace(oPC.Configuration.Stores.FindStoreByID(Me.DELTOStoreID).DelAddress, Chr(13) & Chr(10), Chr(10))
            
            For i = 1 To Me.POLines.Count
                If POLines(i).Fulfilled <> "CAN" Then
                    .elCreateSibling "DetailLine", True
                    .chCreate "SKU"
                    .elText = POLines(i).ProductCodeForExport
                    .elCreateSibling "Title", True
                        .elText = POLines(i).TitleAuthor
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
                End If
            Next i
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
                .elText = IIf(Me.Signature > "", Me.Signature, "")
            .elCreateSibling "OrderMessage", True
                .elText = oPC.Configuration.OrderText
    End With
    
'FINALLY PRODUCE THE .XML FILE
    strXML = strWorkingFolder & "PO_" & Me.DOCCode & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

    If DispatchMode = enMail Then
        'WRITE THE .HTML FILE
            objXSL.async = False
            objXSL.validateOnParse = False
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
            oTF.WriteToTextFile xMLDoc.docObject.transformNode(objXSL)
            oTF.CloseTextFile
    End If
    
    
'WRITE THE .PDF FILE IF NECESSARY
'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file
    If Me.Supplier.DispatchMethod = "M" And DispatchMode = enMail Then
        Set objXSL = Nothing
        Set objXSL = New MSXML2.DOMDocument30
        objXSL.async = False
        objXSL.validateOnParse = False
        objXSL.resolveExternals = False
        strPath = oPC.SharedFolderRoot & "\Templates\PO_DOC_FO.xsl"
        Set fs = New FileSystemObject
        If fs.FileExists(strPath) Then
            objXSL.Load strPath
        End If
    
        
        Set opXMLDOC = New MSXML2.DOMDocument30
        opXMLDOC.async = False
        opXMLDOC.validateOnParse = False
        opXMLDOC.resolveExternals = False
        xMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC
        
        strFOFile = strWorkingFolder & "PO_" & Me.DOCCode & ".FO"
        strPDFFile = strWorkingFolder & "PO_" & Me.DOCCode & ".PDF"
        
 '       If DispatchMode = enMail Then
 '           strDispatchMethod = "Emailing"
 '       ElseIf DispatchMode = enPrint Then
 '           strDispatchMethod = "Printing"
 '       End If
        
        docWriteTostream strFOFile, opXMLDOC, "UNICODE"
        
'Stage 2 Convert the .FO file to .PDF and clean up
        strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFile & " " & strPDFFile
        F_7_AB_1_ShellAndWaitSimple strCommand, vbHide
        If fs.FileExists(strFOFile) Then
            fs.DeleteFile strFOFile
        End If
    End If
    
'Log the dispatch
    Set oSM = New z_StockManager
    Set oFSO = New FileSystemObject
    oSM.LogTransmission Me.TRID, "Dispatched: by " & strDispatchMethod & " : " & Format(Date, "dd/mm/yyyy") & vbCrLf
    Log = "Dispatched: by " & strDispatchMethod & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
    Set oSM = Nothing
    Set oFSO = Nothing
    
    oTF.CloseTextFile
    
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.ExportToXML(bForeign)", bForeign
End Function
Public Sub docWriteTostream(ByVal FilePath As String, obj As MSXML2.DOMDocument30, _
                Optional ByVal CharSet As String = "UNICODE")
    On Error GoTo ErrHandler
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
ErrHandler:
    ErrorIn "ujXML.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
End Sub

Public Function PrintPO(pShowVAT As Boolean, bForeign As Boolean, Optional ByVal pRoundedUp As Boolean) As Boolean
    On Error GoTo ErrHandler
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
ErrHandler:
    ErrPreserve
    oPCL.FinishRequest
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.PrintPO(pShowVAT,bForeign,pRoundedUp)", Array(pShowVAT, bForeign, pRoundedUp)
End Function
Public Function DisplayTotals(pLabels As String, pValues As String, bForeign As Boolean) 'This must be the same as PRINTINVOICE
    On Error GoTo ErrHandler
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
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.DisplayTotals(pLabels,pValues,bForeign)", Array(pLabels, pValues, bForeign)
End Function
Private Sub FillTableRows(pPO As a_PO, bForeign As Boolean, pDiscountExists As Boolean)
    On Error GoTo ErrHandler
Dim i As Integer

    For i = 1 To pPO.POLines.Count
        With pPO.POLines(i)
            If .Discount > 0 Then pDiscountExists = True
            oPCL.POLSend .ProductCodeF, .QtyFirmF, .QtySSF, .TitleAuthor, .PriceF(bForeign), .PLessDiscExtF(bForeign), .DiscountF, PackText(.Note), .Ref '"InvDate", oInv.TransDateFormatted
        End With
    Next i
EXIT_Handler:
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.FillTableRows(pPO,bForeign,pDiscountExists)", Array(pPO, bForeign, pDiscountExists)
End Sub


Private Function TranslateErrors(pRawErrors As String) As String
    On Error GoTo ErrHandler
Dim strOut As String
Dim strRule, strAllRules As String
Dim NoMoreRules As Boolean
Dim iMarker, iStart As Integer
    iMarker = 1
    strAllRules = ""
    If Len(pRawErrors) > 0 Then
        iMarker = InStr(iMarker + 1, pRawErrors, ",")
        If iMarker > 0 Then
            strAllRules = colClassErrors(Left(pRawErrors, iMarker - 1))
        Else
            strAllRules = colClassErrors(pRawErrors)
        End If
        Do Until iMarker = 0
            iStart = iMarker + 1
            iMarker = InStr(iStart, pRawErrors, ",")
            If iMarker > 0 Then
                strRule = colClassErrors(Mid(pRawErrors, iStart, iMarker - iStart))
            Else
                strRule = colClassErrors(Mid(pRawErrors, iStart))
            End If
                
            strAllRules = strAllRules & vbCrLf & strRule
        Loop
    End If
    TranslateErrors = strAllRules
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.TranslateErrors(pRawErrors)", pRawErrors
End Function
Private Sub LoadClassErrorsCollection()
    On Error GoTo ErrHandler
    Set colClassErrors = New Collection
    colClassErrors.Add "Missing supplier", "TP"
    colClassErrors.Add "Note too short", "NOTE"
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.LoadClassErrorsCollection"
End Sub
Private Function CheckInput(pVal As String, pFld As String) As Boolean
    On Error GoTo ErrHandler

    CheckInput = True
    If mudtProps.TPID = 0 Then
        mobjValid.BreakRule "TP", True
        If pFld = "TP" Then CheckInput = False
    Else
        mobjValid.BreakRule "TP", False
    End If

    mobjValid.GetSTatus
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CheckInput(pVal,pFld)", Array(pVal, pFld)
End Function


Public Function Post() As String
    On Error GoTo ErrHandler
Dim oSM As z_StockManager
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.BeginTrans
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
                Post = "In Process"
            Case stCOMPLETE, stISSUED
                oSM.IssuePO Me
              '  SetStatus stCOMPLETE
            Case stPROFORMA
                oSM.IssuePO Me
        '        SetStatus stPROFORMA
            End Select
    Case stCOMPLETE      'Complete
            Post = "ISSUED"
    Case stPROFORMA      'Issued PROFORMA
            Select Case Status
            Case stVOID
                oSM.CancelPO Me
        '        SetStatus stVOID
            Case stInProcess
                Post = "ERROR: was proforma, can't now be inprocess."
            Case stCOMPLETE, stISSUED
                oSM.IssuePO Me
        '        SetStatus stCOMPLETE
            Case stPROFORMA
                Post = "PROFORMA"
            End Select
    End Select
    Set oSM = Nothing
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.Post", , EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Function
Public Sub SetStatus(val As enStatus)
    On Error GoTo ErrHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Status = val
    SetDirty True
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.SetStatus(val)", val
End Sub

Public Property Get CanCancel() As Boolean
    On Error GoTo ErrHandler
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
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CanCancel"
End Property

Public Function VoidDocument()
    On Error GoTo ErrHandler
    Me.BeginEdit
    Me.SetStatus stVOID
    Me.ApplyEdit
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.VoidDocument"
End Function

Public Property Get ISForeignCurrency() As Boolean
    On Error GoTo ErrHandler
    ISForeignCurrency = (Not oPC.Configuration.DefaultCurrency Is Me.CaptureCurrency)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.ISForeignCurrency"
End Property

Public Property Get CurrencyConversionAsText() As String
    On Error GoTo ErrHandler
    CurrencyConversionAsText = mudtProps.CurrencyRate & " " & oPC.Configuration.DefaultCurrency.Description & " per " & Me.CaptureCurrency.Description
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO.CurrencyConversionAsText"
End Property

Public Function GenerateSAANAMsg() As Boolean
    On Error GoTo ErrHandler
Dim str As String
Dim strMsg As String
Dim dte As Date
Dim oTF As New z_TextFile
Dim strfilePath As String
Dim strFileName As String
Dim i As Integer
Dim oSM As z_StockManager
Dim oFSO As New FileSystemObject

    dte = Now()
    strFileName = Me.DOCCode & Format(dte, "HHNN")
    strfilePath = oPC.SharedFolderRoot & "\EDI\UP\" & strFileName & ".TXT"
    
    oTF.OpenTextFile strfilePath
    strMsg = "UNB+UNOA:2+" & oPC.Configuration.GFXNumber & "+" & Me.Supplier.GFXNumber & "+" & Format(dte, "YYMMDD") & ":" & Format(dte, "HHNN") & "+" & strFileName & "+PASSWORD+ORDERS"
    oTF.WriteToTextFile strMsg & "'" & Chr(10)
    strMsg = "UNH+" & strFileName & "+ORDERS:6:0:S"
    oTF.WriteToTextFile strMsg & "'" & Chr(10)
    strMsg = "SOP+5013546146168"
    oTF.WriteToTextFile strMsg & "'" & Chr(10)
    strMsg = "CLO+0144096000000"
    oTF.WriteToTextFile strMsg & "'" & Chr(10)
    strMsg = "ORD+" & Me.DOCCode & "::" & Format(dte, "YYMMDD") & "+N"  '????????????????????????????????
    oTF.WriteToTextFile strMsg & "'" & Chr(10)
    strMsg = "DIN+" & Format(DateAdd("d", 1, dte), "YYMMDD") & "+" & Format(DateAdd("m", 6, dte), "YYMMDD")
    oTF.WriteToTextFile strMsg & "'" & Chr(10)
    
    For i = 1 To Me.POLines.Count
    'QtyFirm + QtySS?
        If Len(Me.POLines(i).ProductCode) = 10 Or Len(Me.POLines(i).ProductCode) = 13 Then
            strMsg = "OLD+" & CStr(i) & "+::" & Me.POLines(i).ProductCode & ":+" & CStr(Me.POLines(i).QtyFirm + Me.POLines(i).QtySS) & "+" & CStr((POLines(i).Price(False) / oPC.Configuration.DefaultCurrency.Divisor)) & "++" & CStr((POLines(i).PExt(False) / oPC.Configuration.DefaultCurrency.Divisor)) & "+++0.00"
            oTF.WriteToTextFile strMsg & "'" & Chr(10)
        End If
    Next
    strMsg = "UNT+" & CStr(6 + Me.POLines.Count) & "+" & strFileName   'Assuming DIN is not needed
    oTF.WriteToTextFile strMsg & "'" & Chr(10)
    strMsg = "UNZ+1+" & strFileName
    oTF.WriteToTextFile strMsg & "'" & Chr(10)
    
    Set oSM = New z_StockManager
    oSM.LogTransmission Me.TRID, "File: " & oFSO.GetFileName(oTF.FileName) & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf
    Log = "File: " & oFSO.GetFileName(oTF.FileName) & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
    Set oSM = Nothing
    
    oTF.CloseTextFile
    Exit Function
ErrHandler:
    ErrorIn "a_PO.GenerateSAANAMsg"
End Function
Public Property Get Log() As String
    On Error GoTo ErrHandler
    Log = Trim(mudtProps.Log)
    Exit Property
ErrHandler:
    ErrorIn "a_PO.Log"
End Property

Public Property Let Log(val As String)
    On Error GoTo ErrHandler
    mudtProps.Log = Trim(val)
    Exit Property
ErrHandler:
    ErrorIn "a_PO.Log(val)", val
End Property

Private Function ValidateObject(pVal As String, pFld As String) As Boolean
    On Error GoTo ErrHandler
Dim bValid As Boolean
    bValid = True
    If mudtProps.TPID = 0 Then
        mobjValid.BreakRule "TP", True
        If pFld = "TP" Then bValid = False
    Else
        mobjValid.BreakRule "TP", False
    End If

    mobjValid.GetSTatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO.ValidateObject(pVal,pFld)", Array(pVal, pFld)
End Function

