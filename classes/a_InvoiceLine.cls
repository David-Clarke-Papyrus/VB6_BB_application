VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_InvoiceLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------------------------------------------
'HANDLING OUTPUT IN FOREIGN CURRENCY and capture and local display in local currency formats
'Definitions:
    'Localprice: a number representing the price in the default currency as determined in the configuration table (stored in the smallest denominated currency unit e.g. cent)
    'ForeignPrice: a number representing the price in the selected foreign currency as determined in the Invoice table (stored in the smallest denominated currency unit e.g. cent)
    'Factor: a number that when applied to the localprice yields the foreignprice (both stored in the smallest denominated currency unit e.g. cent)
    'Divisor: a number that when divided into the localprice or the foreignprice yields a number representing the price in the larger currency unit (e.g. Pounds, dollars, rands)
    'FormatString: a text string used with the Format function to represent a localprice or a foreignprice (After dividing by Divisor) as a properly formatted currency
'Each IL keeps the price and deposit (the only two actual prices, the other are percentages) in two forms,
'one as local currency (Long representing smallest denominated currency unit e.g. cent)
'and the other as foreign currency in the same representation
'The Invoice object records a currency to use for output (it may be just the local currency) and the factor at time of issue
'When the invoice is to be shown as it would appear printed, the calcTotalsOP method of the invoice object is called,
'being passed the foreign currency factor to use. All the properties showing prices anc totals formatted now use the
'foreign currency Divisor and factor and formatstring
'------------------------------------------------------------------------------------------------------------------------

Event Valid(msg As String)
'Event ExtensionChange(lngExtension As Long, strExtension As String)
Private mudtProps As ILProps
Private mcolStack As Collection
Private mobjParent As Object
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassors As Collection
Private mInvoiceParent As Object
Private WithEvents mchCOFFs As ch_COFF
Attribute mchCOFFs.VB_VarHelpID = -1

Private mKey As String
Private moProd As a_Product

Private mFPAfterDiscount As Long
Private mLPAfterDiscount As Long

Private mFPAfterDiscountExt As Long
Private mLPAfterDiscountLineValue As Long
Private mLPAfterDiscountLineValueExVAT As Long
Private mLPExt As Long
Private mFPExt As Long
Private mFPDiscExt As Long
Private mLPLineDiscountValue As Long
Private mLPDiscExVATExt As Long

Private mFPAfterDiscountExtNonVAT As Long
Private mLPAfterDiscountLineValueNonVAT As Long

Private mFPAfterDiscountLineValueVAT As Long
Private mLPAfterDiscountLineValueVAT As Long

Private mFPDisc_Unit As Long
Private mLPDisc_Unit As Long

Private mFPDisc_Ext As Long
Private mLPDisc_Ext As Long
Private mExVATUnitPrice As Long
Private mSelected As Boolean

Public Property Let Selected(val As Boolean)
    mSelected = val
End Property
Public Property Get Selected() As Boolean
    Selected = mSelected
End Property

Private Sub mchCoffs_CoffQtyChanged()
    On Error GoTo errHandler
    MsgBox "Here"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.mchCoffs_CoffQtyChanged"
End Sub
Public Property Get COFFs() As ch_COFF
    Set COFFs = mchCOFFs
End Property

Public Property Get Properties(pSrt As enSortField) As Variant
    On Error GoTo errHandler
    Select Case pSrt
    Case enSequence
        Properties = Me.Sequence
    Case enTitle
        Properties = Me.Title
    Case enAuthor
        Properties = Me.MainAuthor
    Case enCode
        Properties = Me.code
    Case enRef
        Properties = Me.Ref
    End Select
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Properties(pSrt)", pSrt
End Property

Public Property Get Sequence() As Long
    On Error GoTo errHandler
    Sequence = mudtProps.Sequence
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Sequence"
End Property
Public Property Let Sequence(val As Long)
    On Error GoTo errHandler
    If val <> mudtProps.Sequence Then
        mudtProps.Sequence = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Sequence(val)", val
End Property

Private Function SetForeignPriceAndDeposit()
    On Error GoTo errHandler
Dim tmpPrice As Long
Dim lngRoundTo As Long
Dim ConvertedPrice As Long
Dim oCurr As a_Currency
    mobjValid.RuleBroken "OVERFLOW", False
    tmpPrice = mudtProps.Price * mInvoiceParent.CurrencyFactor
    If Me.ServiceItem Then
        mudtProps.ForeignPrice = tmpPrice
        mudtProps.ForeignDeposit = mudtProps.Deposit * mInvoiceParent.CurrencyFactor
    Else
        lngRoundTo = oPC.Configuration.RoundingRules.GetRoundTo(tmpPrice)
        ConvertedPrice = (RoundUp(tmpPrice, lngRoundTo))
        mudtProps.ForeignPrice = ConvertedPrice
        mudtProps.ForeignDeposit = mudtProps.Deposit / mInvoiceParent.CurrencyFactor
    End If
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'ov erflow
        mobjValid.RuleBroken "OVERFLOW", True
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.SetForeignPriceAndDeposit"
End Function
Public Function CalculateLine()
    On Error GoTo errHandler
Dim xVATPrice As Double
Dim bDone As Boolean
'First in local currency
    mobjValid.RuleBroken "OVERFLOW", False
    If Not mudtProps.ServiceItem Then
        If oPC.AllowsSSInvoicing Then
            mudtProps.Qty = mudtProps.QtyFirm + mudtProps.QtySS
        End If
    End If
    bDone = False
    If Not mInvoiceParent Is Nothing Then
        If mInvoiceParent.Customer.ShowVAT = False And mInvoiceParent.Customer.VATable = False Then  'We base all calculations on an ex-VAT price
            mExVATUnitPrice = mudtProps.Price
            mLPAfterDiscountLineValue = Round(CDbl(mExVATUnitPrice * mudtProps.Qty) * (CDbl(100) - mudtProps.DiscountPercent) / 100)
            mLPLineDiscountValue = (CDbl(mudtProps.Price * mudtProps.Qty)) - mLPAfterDiscountLineValue
            mLPAfterDiscountLineValueVAT = 0
            mLPDiscExVATExt = mLPLineDiscountValue
            mLPAfterDiscountLineValueNonVAT = mLPAfterDiscountLineValue - mLPAfterDiscountLineValueVAT
            mLPExt = mudtProps.Price * mudtProps.Qty
            bDone = True
        End If
    End If
    If bDone = False Then
        mLPAfterDiscountLineValue = Round(CDbl(mudtProps.Price * mudtProps.Qty) * (CDbl(100) - mudtProps.DiscountPercent) / 100)
        mLPLineDiscountValue = (CDbl(mudtProps.Price * mudtProps.Qty)) - mLPAfterDiscountLineValue
        mLPAfterDiscountLineValueVAT = Round(mLPAfterDiscountLineValue - (CDbl(mLPAfterDiscountLineValue) * 100 / (CDbl(100) + mudtProps.VATRate)))
        
        mLPDiscExVATExt = mLPLineDiscountValue * 100# / (CDbl(100) + mudtProps.VATRate)
    
        mLPAfterDiscountLineValueNonVAT = mLPAfterDiscountLineValue - mLPAfterDiscountLineValueVAT
        mLPExt = mudtProps.Price * mudtProps.Qty
    End If
'Then in Foreign if appropriate
    If mInvoiceParent Is Nothing Then Exit Function
    If Not mInvoiceParent.ForeignCurrency Is Nothing Then
        If mInvoiceParent.ForeignCurrency.ID <> oPC.Configuration.DefaultCurrency.ID Then
            SetForeignPriceAndDeposit
            mFPAfterDiscountExt = Round(CDbl(mudtProps.ForeignPrice * mudtProps.Qty) * ((CDbl(100) - mudtProps.DiscountPercent) / 100))
            mFPDiscExt = (CDbl(mudtProps.ForeignPrice * mudtProps.Qty)) - mFPAfterDiscountExt
            mFPAfterDiscountLineValueVAT = Round(mFPAfterDiscountExt - (CDbl(mFPAfterDiscountExt) * 100 / (CDbl(100) + mudtProps.VATRate)))
            mFPAfterDiscountExtNonVAT = mFPAfterDiscountExt - mFPAfterDiscountLineValueVAT
            mFPExt = mudtProps.ForeignPrice * mudtProps.Qty
        End If
    End If
    Exit Function
    
errHandler:
    ErrPreserve
    If Err = 6 Then  'overflow
        mobjValid.RuleBroken "OVERFLOW", True
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.CalculateLine"
End Function

Public Property Get PAfterDiscount(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PAfterDiscount = mFPAfterDiscount
    Else
        PAfterDiscount = mLPAfterDiscount
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PAfterDiscount(val)", val
End Property
Public Property Get PAfterDiscountF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PAfterDiscountF = Format(CDbl(mFPAfterDiscount / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        PAfterDiscountF = Format(CDbl(mLPAfterDiscount / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PAfterDiscountF(val)", val
End Property

Public Property Get PAfterDiscountExt(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PAfterDiscountExt = mFPAfterDiscountExt
    Else
        PAfterDiscountExt = mLPAfterDiscountLineValue
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PAfterDiscountExt(val)", val
End Property

Public Property Get PAfterDiscountExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PAfterDiscountExtF = Format(CDbl(mFPAfterDiscountExt / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        PAfterDiscountExtF = Format(CDbl(mLPAfterDiscountLineValue / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PAfterDiscountExtF(val)", val
End Property
'mLPDiscExVATExt
Public Property Get PDiscExtExVAT() As Long
    On Error GoTo errHandler
        PDiscExtExVAT = mLPDiscExVATExt
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PDiscExtExVAT"
End Property
Public Property Get PDiscExtExVATF(Optional BlankWhenZero As Boolean = False) As String
    On Error GoTo errHandler
        If mLPDiscExVATExt = 0 And BlankWhenZero Then
            PDiscExtExVATF = ""
        Else
            PDiscExtExVATF = Format(CDbl(mLPDiscExVATExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
        End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PDiscExtF"
End Property

Public Property Get PDiscExt(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PDiscExt = mFPDiscExt
    Else
        PDiscExt = mLPLineDiscountValue
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PDiscExt(val)", val
End Property
Public Property Get PDiscExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PDiscExtF = Format(CDbl(mFPDiscExt / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        PDiscExtF = Format(CDbl(mLPLineDiscountValue / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PDiscExtF(val)", val
End Property
Public Property Get PExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PExtF = Format(CDbl(mFPExt / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        PExtF = Format(CDbl(mLPExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PExtF(val)", val
End Property
Public Property Get PExt(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PExt = mFPExt
    Else
        PExt = mFPExt
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PExt(val)", val
End Property

Public Property Get PAfterDiscountLineValueVAT(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PAfterDiscountLineValueVAT = mFPAfterDiscountLineValueVAT
    Else
        PAfterDiscountLineValueVAT = mLPAfterDiscountLineValueVAT
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PAfterDiscountLineValueVAT(val)", val
End Property
Public Property Get PAfterDiscountLineValueVATF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PAfterDiscountLineValueVATF = Format(CDbl(mFPAfterDiscountLineValueVAT / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        PAfterDiscountLineValueVATF = Format(CDbl(mLPAfterDiscountLineValueVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PAfterDiscountLineValueVATF(val)", val
End Property

Public Property Get PAfterDiscountExtNonVAT(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PAfterDiscountExtNonVAT = mFPAfterDiscountExtNonVAT
    Else
        PAfterDiscountExtNonVAT = mLPAfterDiscountLineValueNonVAT
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PAfterDiscountExtNonVAT(val)", val
End Property
Public Property Get PAfterDiscountExtNonVATF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PAfterDiscountExtNonVATF = Format(CDbl(mFPAfterDiscountExtNonVAT / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        PAfterDiscountExtNonVATF = Format(CDbl(mLPAfterDiscountLineValueNonVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PAfterDiscountExtNonVATF(val)", val
End Property
Public Property Get PAfterDiscountExtNonVAT_withDecimal(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PAfterDiscountExtNonVAT_withDecimal = Format(mFPAfterDiscountExtNonVAT / mInvoiceParent.ForeignCurrency.Divisor, "#######0.00")
    Else
        PAfterDiscountExtNonVAT_withDecimal = Format(mLPAfterDiscountLineValueNonVAT / oPC.Configuration.DefaultCurrency.Divisor, "#######0.00")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PAfterDiscountExtNonVAT_withDecimal(val)", val
End Property

Private Function GetState() As String
    On Error GoTo errHandler
Dim mudtData As ILData
  
    LSet mudtData = mudtProps
    GetState = mudtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.GetState"
End Function
    
Private Sub SetState(buffer As String)
    On Error GoTo errHandler
Dim mudtData As ILData
  
    mudtData.buffer = buffer
    LSet mudtProps = mudtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
'Dim objPB As PropertyBag
'
'  Set objPB = New PropertyBag
'  With objPB
'    .WriteProperty "State", GetState
'    .WriteProperty "COFFs", mchCOFFs.GetSuperState
'    GetSuperState = .Contents
'  End With
'  Set objPB = Nothing
    GetSuperState = GetState

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
'Dim objPB As PropertyBag
'Dim arBuffer() As Byte
  
  SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.SetSuperState(buffer)", buffer
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 445
    'look for any other COFFs that may be applicable
    mchCOFFs.ApplyEdit
    mudtProps.SalesValue = Me.PAfterDiscountExtNonVAT(False)
    mudtProps.VATAmount = Me.PAfterDiscountLineValueVAT(False)
    If Not mobjParent Is Nothing Then
        mobjParent.AddDocLine Me
        Set mobjParent = Nothing
    End If
    mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.ApplyEdit"
End Sub
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property
Public Sub BeginEdit()
    On Error GoTo errHandler
    mchCOFFs.BeginEdit
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
    mchCOFFs.CancelEdit
  
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.CancelEdit"
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    Set mchCOFFs = New ch_COFF
    LoadClassorsCollection
    mobjValid.RuleBroken "Qty", True
    mobjValid.RuleBroken "Price", True
    mudtProps.IsNew = True
    Set moProd = New a_Product
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Class_Initialize"
End Sub
Friend Sub Initialize(Parent As Object)
    On Error GoTo errHandler
    If mudtProps.IsNew Then
        Set mobjParent = Parent
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Initialize(Parent)", Parent
End Sub
Public Sub SetParentInvoice(pParentInvoice As Object)
    On Error GoTo errHandler
    Set mInvoiceParent = pParentInvoice
    If Not mInvoiceParent Is Nothing Then
        If mInvoiceParent.DocDiscountRate > 0 Then
            If Not Me.ServiceItem Then
                mudtProps.DiscountPercent = mInvoiceParent.DocDiscountRate
            End If
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.SetParentInvoice(pParentInvoice)", pParentInvoice
End Sub
Friend Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.IsValid"
End Property

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set moProd = Nothing
    Set mchCOFFs = Nothing
    If mobjParent Is Nothing Then
      If mcolStack.Count > 0 Then _
        Err.Raise vbObjectError + 1001, , "State stack is not empty"
    End If
    Set mcolStack = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.mobjValid_Status(pMsg)", pMsg
End Sub

'Friend Sub Load(InvoiceLineID As Long)
'    If mcolStack.Count > 0 Then Err.Raise 445
'    If Not mudtProps.IsNew Then Err.Raise 445
'
'    mudtProps.IsNew = False
'    Fetch InvoiceLineID
'End Sub

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Delete"
End Sub
Public Property Get BottomOfDocument() As Boolean
    On Error GoTo errHandler
    BottomOfDocument = (Right(CodeF, 2) = "_B")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.BottomOfDocument"
End Property
Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty Or mchCOFFs.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.IsDirty"
End Property


Public Property Get InvoiceLineID() As Long
    On Error GoTo errHandler
    InvoiceLineID = mudtProps.InvoiceLineID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.InvoiceLineID"
End Property
Public Property Let InvoiceLineID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.InvoiceLineID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.InvoiceLineID"
End Property

Public Property Let InvoiceID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.InvoiceID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.InvoiceID(val)", val
End Property
Public Property Get InvoiceID() As Long
    On Error GoTo errHandler
  InvoiceID = mudtProps.InvoiceID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.InvoiceID"
End Property
Public Property Get Product() As a_Product
    Set Product = moProd
End Property
'Public Function SetLineProduct(Optional PID As String, Optional pCode As String) As Boolean
'    If Not moProd.IsNew Then
'        moProd.PrepareForReload
'    End If
'    With moProd
'        If .Load(PID, 0, Trim$(pCode)) <> 99 Then   'product found
'            Me.Title = .TitleAuthorPublisher
'            Me.PID = .PID
'            Me.Price = .RRP
'            If Not .DefaultCopy Is Nothing Then
'                Me.CopyID = .DefaultCopy.ID
'            End If
'            Me.ProductCode = .Code
'            Me.ProductCodeF = .Code
'            SetLineProduct = True
'        Else
'            SetLineProduct = False   'Not found
'        End If
'    End With
'
'    mobjValid.GetStatus
'End Function

Public Property Get PID() As String
    On Error GoTo errHandler
    PID = FNS(mudtProps.PID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PID"
End Property
Public Property Let PID(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PID = val
    If (mudtProps.PID <= "") Then
        mobjValid.RuleBroken "PID", True
    Else
        mobjValid.RuleBroken "PID", False
    End If
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PID(val)", val
End Property
Public Property Get PIID() As Long
    On Error GoTo errHandler
    PIID = mudtProps.PIID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PIID"
End Property
Public Property Let PIID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PIID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PIID(val)", val
End Property

Public Property Get Qty() As Long
    On Error GoTo errHandler
    Qty = mudtProps.Qty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Qty"
End Property
Public Property Get QtyF() As String
    On Error GoTo errHandler
    If ServiceItem Then   'e.g. postage, insurance etc
        QtyF = ""
    Else
        QtyF = Format(Qty, "##0")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.QtyF"
End Property
Public Property Let Qty(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If mudtProps.Qty <> val Then
        mudtProps.Qty = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Qty(val)", val
End Property
'============
Public Property Get QtySS() As Long
    On Error GoTo errHandler
    QtySS = mudtProps.QtySS
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.QtySS"
End Property
Public Property Get QtySSF() As String
    On Error GoTo errHandler
    If ServiceItem Then   'e.g. postage, insurance etc
        QtySSF = ""
    Else
        QtySSF = Format(QtySS, "##0")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.QtySSF"
End Property
Public Property Let QtySS(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If mudtProps.QtySS <> val Then
        mudtProps.QtySS = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.QtySS(val)", val
End Property
'==============
Public Property Get QtyFirm() As Long
    On Error GoTo errHandler
    QtyFirm = mudtProps.QtyFirm
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.QtyFirm"
End Property
Public Property Get QtyFirmF() As String
    On Error GoTo errHandler
    If ServiceItem Then   'e.g. postage, insurance etc
        QtyFirmF = ""
    Else
        QtyFirmF = Format(QtyFirm, "##0")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.QtyFirmF"
End Property
Public Property Let QtyFirm(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If mudtProps.QtyFirm <> val Then
        mudtProps.QtyFirm = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.QtyFirm(val)", val
End Property
'================




Public Property Get COLID() As Long
    On Error GoTo errHandler
    COLID = mudtProps.COLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.COLID"
End Property
Public Property Let COLID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If mudtProps.COLID <> val Then
        mudtProps.COLID = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.COLID(val)", val
End Property

Public Property Get Publisher() As String
    On Error GoTo errHandler
    Publisher = stripCRLF(FNS(mudtProps.Publisher))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Publisher"
End Property
Public Property Let Publisher(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Publisher = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Publisher(val)", val
End Property
Public Property Get DistributorName() As String
    On Error GoTo errHandler
    DistributorName = stripCRLF(FNS(mudtProps.DistributorName))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.DistributorName"
End Property
Public Property Let DistributorName(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DistributorName = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.DistributorName(val)", val
End Property
Public Property Get DistributorAcno() As String
    On Error GoTo errHandler
    DistributorAcno = stripCRLF(FNS(mudtProps.DistributorAcno))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.DistributorAcno"
End Property
Public Property Let DistributorAcno(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DistributorAcno = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.DistributorAcno(val)", val
End Property

'Public Property Get Description() As String
'    Description = Trim$(mudtProps.Description)
'End Property
'Public Property Let Description(Val As String)
'  If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.Description = Val
'    mudtProps.IsDirty = True
'End Property

Public Property Get Note() As String
    On Error GoTo errHandler
    Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Note"
End Property
Public Property Let Note(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Note = stripCRLF(val)
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Note(val)", val
End Property

Public Property Get Ref() As String
    On Error GoTo errHandler
    Ref = stripCRLF(FNS(mudtProps.Ref))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Ref"
End Property
Public Property Let Ref(val As String)
    On Error GoTo errHandler
    mudtProps.Ref = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Ref(val)", val
End Property

Public Property Get code() As String
    On Error GoTo errHandler
    code = stripCRLF(FNS(mudtProps.code))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Code"
End Property
Public Property Get EAN() As String
    On Error GoTo errHandler
    EAN = stripCRLF(FNS(mudtProps.EAN))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.EAN"
End Property
Sub TestEvent()
    On Error GoTo errHandler
RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.TestEvent"
End Sub
Public Property Let code(val As String)
    On Error GoTo errHandler
Dim oProdCode As z_ProdCode
Dim oSM As New z_StockManager

  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.code = val
    If Left(val, 1) <> "2" Then
    If Len(val) = 13 And mudtProps.CodeF = "" Then
        mudtProps.CodeF = FormatISBN13(val)
    Else
       ' mudtProps.CodeF = val
    End If
    End If
'    Set oProdCode = New z_ProdCode
'    If Len(val) = 13 Then
'        oProdCode.LoadNew "", val, , "", True
'    Else
'        oProdCode.LoadNew val, "", , "", True
'    End If
'    mudtProps.CodeF = oProdCode.FormatCode(False)
'    Set oProdCode = Nothing
    
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Code(val)", val
End Property
Public Property Let CNLQty(val As Integer)
    mudtProps.tmpCNLQty = val
End Property
Public Property Get CNLQty() As Integer
    CNLQty = FNN(mudtProps.tmpCNLQty)
End Property
Public Sub SetMarker(bON As Boolean)
    mudtProps.Marker = bON
End Sub
Public Property Get Marker() As Boolean
    Marker = FNB(mudtProps.Marker)
End Property
Public Property Get CodeF() As String
    On Error GoTo errHandler
    CodeF = FNS(mudtProps.CodeF)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.CodeF"
End Property
Public Property Let CodeF(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CodeF = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.CodeF(val)", val
End Property
Public Property Get CodeForExport() As String
    On Error GoTo errHandler
  CodeForExport = FNS(mudtProps.CodeForExport)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.CodeForExport"
End Property
Public Property Let CodeForExport(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CodeForExport = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.CodeForExport(val)", val
End Property
Public Property Get TitleAuthor()
    On Error GoTo errHandler
Dim tmp As String

    tmp = Title
    If Len(tmp) > 0 Then
        tmp = tmp & IIf(Len(MainAuthor) > 0, "(" & MainAuthor & ")", "")
    End If
    TitleAuthor = stripCRLF(tmp)

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.TitleAuthor"
End Property
Public Property Get TitleAuthorPublisher()
    On Error GoTo errHandler
Dim tmp As String

    tmp = Title
    If Len(tmp) > 0 Then
        tmp = tmp & IIf(Len(MainAuthor) > 0, "(" & MainAuthor & ")", "")
    End If
    If Len(tmp) > 0 Then
        tmp = tmp & IIf(Len(Publisher) > 0, "(" & Publisher & ")", "")
    End If
    TitleAuthorPublisher = stripCRLF(tmp)

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.TitleAuthorPublisher"
End Property
Public Property Get TitleAuthorPublisherWithoutBRackets()
    On Error GoTo errHandler
Dim tmp As String

    tmp = Title
    If Len(tmp) > 0 Then
        tmp = tmp & IIf(Len(MainAuthor) > 0, " " & MainAuthor & " ", "")
    End If
    If Len(tmp) > 0 Then
        tmp = tmp & IIf(Len(Publisher) > 0, "(" & Publisher & ")", "")
    End If
    TitleAuthorPublisherWithoutBRackets = stripCRLF(tmp)

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.TitleAuthorPublisherWithoutBRackets"
End Property
Public Property Get TitleAuthorPublisherL(pLen As Integer)
    On Error GoTo errHandler
    TitleAuthorPublisherL = stripCRLF(Left(TitleAuthorPublisher, pLen) & IIf(Len(TitleAuthorPublisher) > pLen, "...", ""))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.TitleAuthorPublisherL(pLen)", pLen
End Property
Public Property Get SpecialVATMsg() As String
    On Error GoTo errHandler
    If Me.VATRate <> oPC.Configuration.VATRate And oPC.Configuration.IsVATRegion Then
         SpecialVATMsg = "(" & VATRateF & " VAT) "
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.SpecialVATMsg"
End Property
Public Property Get Title() As String
    On Error GoTo errHandler
    Title = stripCRLF(FNS(mudtProps.Title))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Title"
End Property
Public Property Let Title(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Title = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Title(val)", val
End Property
Public Property Get TitleF(pLen As Integer) As String
    On Error GoTo errHandler
    TitleF = Left(Title, pLen) & IIf(Len(Title) > pLen, "...", "")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.TitleF(pLen)", pLen
End Property
Public Property Get GDNCode() As String
    On Error GoTo errHandler
    GDNCode = FNS(mudtProps.GDNCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.GDNCode"
End Property

Public Property Get MainAuthor() As String
    On Error GoTo errHandler
    MainAuthor = stripCRLF(FNS(mudtProps.MainAuthor))
     '   tmp = tmp & IIf(Len(MainAuthor) > 0, " " & MainAuthor & " ", "")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.MainAuthor"
End Property

Public Property Get CO_StaffShortname() As String
    On Error GoTo errHandler
    CO_StaffShortname = FNS(mudtProps.CO_StaffShortname)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.CO_StaffShortname"
End Property
Public Property Let MainAuthor(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.MainAuthor = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.MainAuthor(val)", val
End Property

Public Property Get DiscountPercent() As Double
    On Error GoTo errHandler
    DiscountPercent = mudtProps.DiscountPercent
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.DiscountPercent"
End Property
Public Property Let DiscountPercent(val As Double)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If Not ServiceItem Then  'non ServiceItem items cant have discount
     '   If mInvoiceParent.DocDiscountRate > 0 Then 'we only allow discount here if no general discount applicable
     '       mudtProps.DiscountPercent = mInvoiceParent.DocDiscountRate
     '   Else
            mudtProps.DiscountPercent = val
     '   End If
        If (val < 0 Or val > 100) Then
            mobjValid.RuleBroken "Discount", True
        Else
            mobjValid.RuleBroken "Discount", False
        End If
        mudtProps.IsDirty = True
  '      RaiseEvent ExtensionChange(Extension, ExtensionF)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.DiscountPercent(val)", val
End Property
Public Function SetDiscountPercent(val As String) As Boolean
    On Error GoTo errHandler
Dim dblDiscountPercent As Double
    SetDiscountPercent = True
    If Not ServiceItem Then  'non-stock items can't have discount
        If Not mInvoiceParent Is Nothing Then
            If mInvoiceParent.DocDiscountRate > 0 Then 'we only allow discount here if no general discount applicable
                mudtProps.DiscountPercent = mInvoiceParent.DocDiscountRate
                Exit Function
            End If
        End If
        
        If mcolStack.Count = 0 Then Err.Raise 383
        If Trim$(val) = "" Then
            dblDiscountPercent = 0
        ElseIf Not ConvertToDBL(val, dblDiscountPercent) Then
            SetDiscountPercent = False
            Exit Function
        End If
        If (dblDiscountPercent < 0 Or dblDiscountPercent > 100) Then
            mobjValid.RuleBroken "Discount", True
        Else
            mobjValid.RuleBroken "Discount", False
        End If
        mudtProps.DiscountPercent = dblDiscountPercent
        mudtProps.IsDirty = True
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.SetDiscountPercent(val)", val
End Function
Public Function DiscountPercentF() As String
    On Error GoTo errHandler
    If mudtProps.DiscountPercent = 0 Then
        DiscountPercentF = ""
    Else
        DiscountPercentF = Format(mudtProps.DiscountPercent, gPercentFormatString)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.DiscountPercentF"
End Function
Public Property Get VATRate() As Double
    On Error GoTo errHandler
    VATRate = mudtProps.VATRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.VATRate"
End Property
Public Property Get VATRateF() As String
    On Error GoTo errHandler
    VATRateF = Format(mudtProps.VATRate, gPercentFormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.VATRateF"
End Property

Public Property Let VATRate(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.VATRate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.VATRate(val)", val
End Property
Public Property Get ServiceItem() As Boolean
    On Error GoTo errHandler
    ServiceItem = mudtProps.ServiceItem
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.ServiceItem"
End Property
Public Property Let ServiceItem(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ServiceItem = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.ServiceItem(val)", val
End Property
'---------------------------------
'Public Property Get VATAmount() As Long
'    VATAmount = CLng(CDbl((VATRate / (100 + VATRate))) * ExtensionInclDeposit)
'End Property
'Public Property Get VATAmount_foreign() As Long
'    VATAmount_foreign = CLng((VATRate / (100 + VATRate)) * ExtensionInclDeposit_Foreign)
'End Property
'Public Property Get VATAmountF(bForeign As Boolean) As String
'    If bForeign Then
'        VATAmountF = Format(CDbl(VATAmount / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
'    Else
'        VATAmountF = Format(CDbl(VATAmount / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
'    End If
'End Property
'Public Property Get ExtensionExVAT() As Long
'    ExtensionExVAT = (CLng((mudtProps.Price * mudtProps.Qty) * ((100 - mudtProps.DiscountPercent) / 100))) * (100 / (100 + mudtProps.VATRate))
'End Property
'Public Property Get ExtensionExVATF(bForeign As Boolean) As String
'    If bForeign Then
'        ExtensionExVATF = Format(CDbl(ExtensionExVAT / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
'    Else
'        ExtensionExVATF = Format(CDbl(ExtensionExVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
'    End If
'End Property
'Public Property Get ExtensionInclDeposit() As Long
'    ExtensionInclDeposit = CLng(CLng(mudtProps.Price * mudtProps.Qty) * ((100 - mudtProps.DiscountPercent) / 100))
'End Property
'Public Property Get ExtensionInclDeposit_Foreign() As Long
'    ExtensionInclDeposit_Foreign = CLng(CLng(ForeignPrice * mudtProps.Qty) * ((100 - mudtProps.DiscountPercent) / 100))
'End Property
'Public Property Get Extension() As Long
'    Extension = CLng(CLng(mudtProps.Price * mudtProps.Qty) * ((100 - mudtProps.DiscountPercent) / 100))
'End Property
'Public Property Get Extension_Foreign() As Long
'    Extension_Foreign = CLng(CLng(ForeignPrice * mudtProps.Qty) * ((100 - mudtProps.DiscountPercent) / 100))
'End Property
'Public Property Get ForeignPrice() As Long
'    ForeignPrice = mudtProps.ForeignPrice 'RoundUp(mudtProps.ForeignPrice / mInvoiceParent.ForeignCurrency.Divisor, oPC.Configuration.RoundingRules.GetRoundTo(mudtProps.ForeignPrice / mInvoiceParent.ForeignCurrency.Divisor)) * mInvoiceParent.ForeignCurrency.Divisor
'End Property
'
'Public Property Get ExtensionF(bForeign As Boolean) As String
'    If bForeign Then
'        ExtensionF = Format(CDbl(Extension_Foreign / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
'    Else
'        ExtensionF = Format(CDbl(Extension / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
'    End If
'End Property
'Public Property Get ExtensionSimple() As Long
'    ExtensionSimple = CLng(mudtProps.Price * mudtProps.Qty)
'End Property
'Public Property Get ExtensionSimple_foreign() As Long
'    ExtensionSimple_foreign = CLng(mudtProps.ForeignPrice * mudtProps.Qty)
'End Property
'Public Property Get ExtensionSimpleF(bForeign As Boolean) As String
'    If bForeign Then
'        ExtensionSimpleF = Format(CDbl(ExtensionSimple_foreign / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
'    Else
'        ExtensionSimpleF = Format(CDbl(ExtensionSimple / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
'    End If
'End Property
Public Function PriceF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        PriceF = Format(CDbl(mudtProps.ForeignPrice / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        PriceF = Format(CDbl(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PriceF(bForeign)", bForeign
End Function
Public Function PriceExVatF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        'surely this wont be used
        PriceExVatF = Format(CDbl(mudtProps.ForeignPrice * (100 / (100 + oPC.Configuration.VATRate)) / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        PriceExVatF = Format(CDbl(Me.PriceExVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PriceExVatF(bForeign)", bForeign
End Function

'Public Property Get DiscountAmount() As Long
'    DiscountAmount = CLng(ExtensionSimple - ExtensionInclDeposit)
'End Property
'Public Property Get DiscountAmount_Foreign() As Long
'    DiscountAmount_Foreign = CLng(ExtensionSimple_foreign - ExtensionInclDeposit_Foreign)
'End Property
'Public Property Get DiscountAmountF(bForeign As Boolean) As String
'    If bForeign Then
'        DiscountAmountF = Format(CDbl(DiscountAmount / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
'    Else
'        DiscountAmountF = Format(CDbl(DiscountAmount / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
'    End If
'End Property

Public Property Get Deposit_Foreign() As Long
    On Error GoTo errHandler
    Deposit_Foreign = mudtProps.ForeignDeposit
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Deposit_Foreign"
End Property
Public Property Let Deposit_Foreign(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ForeignDeposit = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Deposit_Foreign(val)", val
End Property
Public Property Get Deposit() As Long
    On Error GoTo errHandler
    Deposit = mudtProps.Deposit
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Deposit"
End Property
Public Property Let Deposit(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.Deposit Then
        mudtProps.Deposit = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Deposit(val)", val
End Property


Public Property Get APPLID() As Long
    On Error GoTo errHandler
    APPLID = mudtProps.APPLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.APPLID"
End Property
Public Property Let APPLID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.APPLID Then
        mudtProps.APPLID = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.APPLID(val)", val
End Property

Public Property Get APPLQTY() As Long
    On Error GoTo errHandler
    APPLQTY = mudtProps.APPLQTY
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.APPLQTY"
End Property
Public Property Let APPLQTY(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.APPLQTY Then
        mudtProps.APPLQTY = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.APPLQTY(val)", val
End Property
Public Property Let SubstitutesAvailable(pVal As Boolean)
    mudtProps.SubstitutesAvailable = pVal
    SetDirty (True)
End Property
Public Property Get SubstitutesAvailable() As Boolean
    SubstitutesAvailable = FNB(mudtProps.SubstitutesAvailable)
End Property
Public Property Get CreditedQty() As Long
    On Error GoTo errHandler
    CreditedQty = mudtProps.CreditedQty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.CreditedQty"
End Property

Public Function SetDeposit(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Deposit
    bOK = SetField_strAsCurrencyToLong(mudtProps.Deposit, val, mcolStack.Count, "DEPOSIT", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
    If bOK And (lngTemp <> mudtProps.Deposit) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("DEPOSIT")
    End If
    SetDeposit = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.SetDeposit(val)", val
End Function
Public Function DepositF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        DepositF = Format(CDbl(mudtProps.ForeignDeposit / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        DepositF = Format(CDbl(mudtProps.Deposit / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.DepositF(bForeign)", bForeign
End Function
Public Function SetNote(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Note) Then
        Err.Raise 384
    End If
    If val <> mudtProps.Note Then
        mudtProps.Note = val
        SetDirty True
    End If
    SetNote = True
    Exit Function
End Function
Public Function SetRef(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Ref) Then
        Err.Raise 384
    End If
    If val <> mudtProps.Ref Then
        mudtProps.Ref = val
        SetDirty True
    End If
    SetRef = True
    Exit Function
End Function

'------------------------------
'
'Public Property Get ForeignPrice() As Long
'    ForeignPrice = mudtProps.ForeignPrice
'End Property
Public Property Let ForeignPrice(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.ForeignPrice Then
        mudtProps.ForeignPrice = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.ForeignPrice(val)", val
End Property
Public Property Get ForeignPrice() As Long
    On Error GoTo errHandler
    ForeignPrice = mudtProps.ForeignPrice
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.ForeignPrice"
End Property
Public Property Get PriceExVAT() As Long
    On Error GoTo errHandler
    If (mInvoiceParent.Customer.VATable = False And mInvoiceParent.Customer.ShowVAT = False) Then
        PriceExVAT = mudtProps.Price
    Else
        PriceExVAT = Round(mudtProps.Price * (100 / (100 + mudtProps.VATRate)))
    End If
    'Can't do this as some items are really zero vat - like postage etc
 '   PriceExVAT = Round(mudtProps.Price * (100 / (100 + IIf(mudtProps.VATRate = 0, oPC.Configuration.VATRate, mudtProps.VATRate))))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PriceExVAT"
End Property
Public Property Get Price() As Long
    On Error GoTo errHandler
    Price = mudtProps.Price
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Price"
End Property
Public Property Let Price(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.Price Then
        mudtProps.Price = val
        SetDirty True
    End If
    ValidateObject "PRICE"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Price(val)", val
End Property
Public Property Get Cost() As Long
    On Error GoTo errHandler
    Cost = mudtProps.Cost
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Cost"
End Property
Public Property Let Cost(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.Cost Then
        mudtProps.Cost = val
        SetDirty True
    End If
    ValidateObject "Cost"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Cost(val)", val
End Property
Public Function CostF() As String
    On Error GoTo errHandler
        CostF = Format(CDbl(mudtProps.Cost / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.PCostF"
End Function

Public Property Get CodeForEditing() As String
    If Not IsISBN10(FNS(mudtProps.code)) And FNS(mudtProps.code) > "" Then
        CodeForEditing = FNS(mudtProps.code)
    Else
        If FNS(mudtProps.CodeF) > "" Then
            CodeForEditing = FNS(mudtProps.CodeF)
        Else
            If IsISBN13(FNS(mudtProps.EAN), True) Then
                CodeForEditing = FNS(mudtProps.EAN)
            Else
                If FNS(mudtProps.code) > "" Then
                    CodeForEditing = FNS(mudtProps.code)
                Else
                    CodeForEditing = FNS(mudtProps.EAN)
                End If
            End If
        End If
    End If
    
End Property
Public Function SetPrice(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Price
    bOK = SetField_strAsCurrencyToLong(mudtProps.Price, val, mcolStack.Count, "PRICE", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
    If bOK And (lngTemp <> mudtProps.Price) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("PRICE")
    End If
    SetPrice = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.SetPrice(val)", val
End Function
Public Function SetQty(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Qty
    bOK = SetField_LONG(mudtProps.Qty, val, "QTY", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.Qty) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QTY")
    End If
    SetQty = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.SetQty(val)", val
End Function

Public Function SetQtyFirm(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.QtyFirm
    bOK = SetField_LONG(mudtProps.QtyFirm, val, "QTY", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.QtyFirm) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QTY")
    End If
    SetQtyFirm = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.SetQtyFirm(val)", val
End Function
Public Function SetQtySS(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.QtySS
    bOK = SetField_LONG(mudtProps.QtySS, val, "QTY", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.QtySS) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QTY")
    End If
    SetQtySS = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.SetQtySS(val)", val
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
    ErrorIn "a_InvoiceLine.TranslateErrors(pRawors)", pRawors
End Function

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.SetDirty(pVal)", pVal
End Sub

Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Invalid price", "PRICE"
    colClassors.Add "Invalid quantity", "QTY"
    colClassors.Add "Invalid ETA", "ETA"
    colClassors.Add "Invalid discount", "DISCOUNT"
    colClassors.Add "Invalid deposit", "DEPOSIT"
    colClassors.Add "A numeric value is too great for calculation", "OVERFLOW"
    colClassors.Add "A pre-delivery invoice must match a customer order", "PREDEL"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.LoadClassorsCollection"
End Sub
Private Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean
    bValid = True
    If mudtProps.Price > 0 Then
        If pFld = "PRICE" Then bValid = False
        mobjValid.RuleBroken "PRICE", False
    Else
        mobjValid.RuleBroken "PRICE", True
    End If
    If mudtProps.Qty >= 0 And mudtProps.QtyFirm >= 0 And mudtProps.QtySS >= 0 Then
        mobjValid.RuleBroken "QTY", False
    Else
        If pFld = "QTY" Then bValid = False
        mobjValid.RuleBroken "QTY", True
    End If
    If (mudtProps.DiscountPercent < 0 Or mudtProps.DiscountPercent > 100) Then
        If pFld = "DISCOUNT" Then bValid = False
        mobjValid.RuleBroken "DISCOUNT", True
    Else
        mobjValid.RuleBroken "DISCOUNT", False
    End If
    If Not mInvoiceParent Is Nothing Then
        If mInvoiceParent.IsInvoice Then
            If mInvoiceParent.IsPreDelivery = True And mudtProps.COLID = 0 Then
                mobjValid.RuleBroken "PREDEL", True
            Else
                mobjValid.RuleBroken "PREDEL", False
            End If
        Else
                mobjValid.RuleBroken "PREDEL", False
        End If
    End If
    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
EXIT_Handler:
'H:
'    MsgBox "CheckOK: " & or
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.ValidateObject(pFld)", pFld
End Function

Public Sub GetStatus()
    On Error GoTo errHandler
    ValidateObject ""
    mobjValid.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.GetSTatus"
End Sub

Public Function CoffsValid(pResult As Integer)
    On Error GoTo errHandler
Dim lngTotalQtys As Long
Dim lngTotalQtysAwaiting As Long
Dim oCoff As a_COFF
    lngTotalQtys = 0
    For Each oCoff In COFFs
        lngTotalQtys = oCoff.COFFQTY + lngTotalQtys
        lngTotalQtysAwaiting = oCoff.COLQty - oCoff.COLQtyDispatched + lngTotalQtysAwaiting
    Next
    If lngTotalQtys = Qty Then
        pResult = 1    'We have allocated all we can
    ElseIf lngTotalQtys < Qty Then
        If lngTotalQtysAwaiting > lngTotalQtys Then
            pResult = 2  'We could allocate more
        Else
            pResult = 4 ' WE have tried to allocate more than requested
        End If
    Else
        pResult = 3   'We have tried to allocate more than we can
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.CoffsValid(pResult)", pResult
End Function


Public Function SetForeignPrice(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.ForeignPrice
    bOK = SetField_strAsCurrencyToLong(mudtProps.ForeignPrice, val, mcolStack.Count, "ForeignPrice", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
    If bOK And (lngTemp <> mudtProps.ForeignPrice) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("ForeignPrice")
    End If
    SetForeignPrice = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.SetForeignPrice(val)", val
End Function
'Public Property Let ForeignPrice(val As Long)
'    On Error GoTo errHandler
'  If mcolStack.Count = 0 Then Err.Raise 383
'    If val <> mudtProps.ForeignPrice Then
'        mudtProps.ForeignPrice = val
'        SetDirty True
'    End If
'    ValidateObject "ForeignPrice"
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_QUL.ForeignPrice(val)", val
'End Property
Public Property Get ForeignPriceF() As String
    On Error GoTo errHandler
Dim oFC As New a_Currency
    If mudtProps.FCID < 1 Then
        ForeignPriceF = ""
        Exit Sub
    End If
    Set oFC = oPC.Configuration.Currencies.FindCurrencyByID(mudtProps.FCID)
    ForeignPriceF = Format(CDbl(mudtProps.ForeignPrice / oFC.Divisor), oFC.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.ForeignPriceF"
End Property

Public Function SetFCFactor(val As Double) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim Temp As Double
    Temp = mudtProps.FCFactor
    bOK = SetField_DOUBLE(mudtProps.FCFactor, CStr(val), "FCFactor", mcolStack.Count)
    If bOK And (Temp <> mudtProps.FCFactor) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("FCFactor")
    End If
    SetFCFactor = bOK
    Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.SetFCFactor(val)", val
End Function

Public Property Get FCFactor() As Double
    FCFactor = mudtProps.FCFactor
End Property
Public Property Get FCFactorF() As String
    FCFactorF = Format(FCFactor, "##0.####")
End Property
Public Property Get FCFactorInvF() As String
    If FCFactor <> 0 Then
        FCFactorInvF = Format(Round(1# / FCFactor, 3), "##0.###")
    End If
End Property
Public Property Get FCID() As Long
    FCID = mudtProps.FCID
End Property
Public Property Let FCID(val As Long)
    mudtProps.FCID = val
End Property

