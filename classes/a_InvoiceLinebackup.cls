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
Attribute VB_Exposed = True
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

Event Valid(Msg As String)
'Event ExtensionChange(lngExtension As Long, strExtension As String)
Private mudtProps As ILProps
Private mcolStack As Collection
Private mobjParent As ch_InvoiceLine
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassErrors As Collection
Private mInvoiceParent As a_Invoice
Private WithEvents mchCOFFs As ch_COFF
Attribute mchCOFFs.VB_VarHelpID = -1

Private mKey As String
Private moProd As a_Product

Private Sub mchCoffs_CoffQtyChanged()
    MsgBox "Here"
End Sub
Public Property Get COFFs() As ch_COFF
    Set COFFs = mchCOFFs
End Property

Public Property Get Properties(pSrt As enInvoiceSortField) As Variant
    Select Case pSrt
    Case enSequence
        Properties = Me.Sequence
    Case enTitle
        Properties = Me.Title
    Case enAuthor
        Properties = Me.MainAuthor
    Case enCode
        Properties = Me.Code
    End Select
End Property

Public Property Get Sequence() As Long
    Sequence = mudtProps.Sequence
End Property
Public Property Let Sequence(val As Long)
    If val <> mudtProps.Sequence Then
        mudtProps.Sequence = val
        SetDirty True
    End If
End Property

Public Function SetForeignPriceAndDeposit()
Dim tmpPrice As Long
    tmpPrice = mudtProps.Price / mInvoiceParent.CurrencyFactor
    If Me.NONStock Then
        mudtProps.ForeignPrice = (RoundUp(tmpPrice / mInvoiceParent.ForeignCurrency.Divisor, tmpPrice / mInvoiceParent.ForeignCurrency.Divisor)) * mInvoiceParent.ForeignCurrency.Divisor
        mudtProps.ForeignDeposit = mudtProps.Deposit / mInvoiceParent.CurrencyFactor
    Else
        mudtProps.ForeignPrice = (RoundUp(tmpPrice / mInvoiceParent.ForeignCurrency.Divisor, oPC.Configuration.RoundingRules.GetRoundTo(tmpPrice / mInvoiceParent.ForeignCurrency.Divisor))) * mInvoiceParent.ForeignCurrency.Divisor
        mudtProps.ForeignDeposit = mudtProps.Deposit / mInvoiceParent.CurrencyFactor
    End If
End Function
Private Function GetState() As String
  Dim mudtData As ILData
  
  LSet mudtData = mudtProps
  GetState = mudtData.buffer
End Function

Private Sub SetState(buffer As String)
  Dim mudtData As ILData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Function GetSuperState() As String
Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "COFFs", mchCOFFs.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
End Function
Friend Sub SetSuperState(buffer As String)
Dim objPB As PropertyBag
Dim arBuffer() As Byte
  
  Set objPB = New PropertyBag
  arBuffer = buffer
  With objPB
    .Contents = arBuffer
    SetState .ReadProperty("State")
    mchCOFFs.SetSuperState .ReadProperty("COFFs")
  End With
  Set objPB = Nothing
End Sub

Public Sub ApplyEdit()
    If mcolStack.Count = 0 Then Err.Raise 445
    'look for any other COFFs that may be applicable
    mchCOFFs.ApplyEdit
    mudtProps.SalesValue = ExtensionExVAT
    mudtProps.VATAmount = VATAmount
    If Not mobjParent Is Nothing Then
        mobjParent.AddInvoiceLine Me
        Set mobjParent = Nothing
    End If
    mcolStack.Remove mcolStack.Count
End Sub
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property
Public Sub BeginEdit()
    mchCOFFs.BeginEdit
  mcolStack.Add GetState
End Sub

Public Sub CancelEdit()
  If mcolStack.Count = 0 Then Err.Raise 445
    mchCOFFs.CancelEdit
  
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
End Sub

Private Sub Class_Initialize()
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    Set mchCOFFs = New ch_COFF
    LoadClassErrorsCollection
    mobjValid.RuleBroken "Qty", True
    mobjValid.RuleBroken "Price", True
    mudtProps.IsNew = True
    Set moProd = New a_Product
End Sub
Friend Sub Initialize(Parent As ch_InvoiceLine)
    If mudtProps.IsNew Then
        Set mobjParent = Parent
    End If
End Sub
Friend Sub SetParentInvoice(pParentInvoice As a_Invoice)
    Set mInvoiceParent = pParentInvoice
    If Not mInvoiceParent Is Nothing Then
        If mInvoiceParent.InvoiceDiscountRate > 0 Then
            If Not Me.NONStock Then
                mudtProps.DiscountPercent = mInvoiceParent.InvoiceDiscountRate
            End If
        End If
    End If
End Sub
Friend Property Get IsValid() As Boolean
  IsValid = (mobjValid.Count = 0)
End Property

Private Sub Class_Terminate()
    Set moProd = Nothing
    Set mchCOFFs = Nothing
    If mobjParent Is Nothing Then
      If mcolStack.Count > 0 Then _
        Err.Raise vbObjectError + 1001, , "State stack is not empty"
    End If
    Set mcolStack = Nothing
End Sub

Private Sub mobjValid_BrokenRule(errors As String)
    RaiseEvent Valid(TranslateErrors(errors))
End Sub
Private Sub mobjValid_RuleUnbroken(errors As String)
    RaiseEvent Valid(TranslateErrors(errors))
End Sub

Private Sub mobjValid_NoBrokenRules()
    RaiseEvent Valid("")
End Sub
Private Sub mobjValid_Status(pMsg As String)
    RaiseEvent Valid(TranslateErrors(pMsg))
End Sub

'Friend Sub Load(InvoiceLineID As Long)
'    If mcolStack.Count > 0 Then Err.Raise 445
'    If Not mudtProps.IsNew Then Err.Raise 445
'
'    mudtProps.IsNew = False
'    Fetch InvoiceLineID
'End Sub

Public Sub Delete()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
End Sub
Public Property Get BottomOfDocument() As Boolean
    BottomOfDocument = (Right(Code, 2) = "_B")
End Property
Public Property Get IsDeleted() As Boolean
  IsDeleted = mudtProps.IsDeleted
End Property
Public Property Get IsEditing() As Boolean
  IsEditing = (mcolStack.Count > 0)
End Property

Public Property Get IsNew() As Boolean
  IsNew = mudtProps.IsNew
End Property

Public Property Get IsDirty() As Boolean
  IsDirty = mudtProps.IsDirty Or mchCOFFs.IsDirty
End Property


Public Property Get InvoiceLineID() As Long
    InvoiceLineID = mudtProps.InvoiceLineID
End Property

Public Property Let InvoiceID(val As Long)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.InvoiceID = val
    mudtProps.IsDirty = True
End Property
Public Property Get InvoiceID() As Long
  InvoiceID = mudtProps.InvoiceID
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
    PID = mudtProps.PID
End Property
Public Property Let PID(val As String)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PID = val
    If (mudtProps.PID <= "") Then
        mobjValid.RuleBroken "PID", True
    Else
        mobjValid.RuleBroken "PID", False
    End If
    mudtProps.IsDirty = True
End Property
Public Property Get PIID() As Long
    PIID = mudtProps.PIID
End Property
Public Property Let PIID(val As Long)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PIID = val
    mudtProps.IsDirty = True
End Property

Public Property Get Qty() As Long
    Qty = mudtProps.Qty
End Property
Public Property Get QtyF() As String
    If NONStock Then   'e.g. postage, insurance etc
        QtyF = ""
    Else
        QtyF = Format(Qty, "##0")
    End If
End Property
Public Property Let Qty(val As Long)
  If mcolStack.Count = 0 Then Err.Raise 383
    If mudtProps.Qty <> val Then
        mudtProps.Qty = val
        SetDirty True
    End If
End Property
Public Property Get COLID() As Long
    COLID = mudtProps.COLID
End Property
Public Property Let COLID(val As Long)
  If mcolStack.Count = 0 Then Err.Raise 383
    If mudtProps.COLID <> val Then
        mudtProps.COLID = val
        SetDirty True
    End If
End Property

Public Property Get Publisher() As String
    Publisher = FNS(mudtProps.Publisher)
End Property
Public Property Let Publisher(val As String)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Publisher = val
    mudtProps.IsDirty = True
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
    Note = FNS(mudtProps.Note)
End Property
Public Property Let Note(val As String)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Note = val
    mudtProps.IsDirty = True
End Property

Public Property Get Ref() As String
    Ref = FNS(mudtProps.Ref)
End Property
Public Property Let Ref(val As String)
    mudtProps.Ref = val
End Property

Public Property Get Code() As String
    Code = FNS(mudtProps.Code)
End Property
Sub TestEvent()
RaiseEvent Valid("")
End Sub
Public Property Let Code(val As String)
Dim oCode As z_ProdCode
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Code = val
    Set oCode = New z_ProdCode
    oCode.Load val
    mudtProps.CodeF = oCode.FormatCode
    mudtProps.IsDirty = True
    Set oCode = Nothing
End Property

Public Property Get CodeF() As String
    CodeF = FNS(mudtProps.CodeF)
End Property
Public Property Let CodeF(val As String)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CodeF = val
    mudtProps.IsDirty = True
End Property

Public Property Get CodeForExport() As String
  CodeForExport = FNS(mudtProps.CodeForExport)
End Property
Public Property Let CodeForExport(val As String)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CodeForExport = val
End Property
Public Property Get TitleAuthorPublisher()
Dim tmp As String

    tmp = Title
    If Len(tmp) > 0 Then
        tmp = tmp & IIf(Len(MainAuthor) > 0, "(" & MainAuthor & ")", "")
    End If
    If Len(tmp) > 0 Then
        tmp = tmp & IIf(Len(Publisher) > 0, "(" & Publisher & ")", "")
    End If
    TitleAuthorPublisher = tmp

End Property
Public Property Get TitleAuthorPublisherL(pLen As Integer)
    TitleAuthorPublisherL = Left(TitleAuthorPublisher, pLen) & IIf(Len(TitleAuthorPublisher) > pLen, "...", "")
End Property
Public Property Get SpecialVATMsg() As String
    If Me.VATRate <> oPC.Configuration.VATRate And oPC.Configuration.IsVATRegion Then
         SpecialVATMsg = " (" & VATRateF & " VAT)"
    End If
End Property
Public Property Get Title() As String
    Title = FNS(mudtProps.Title)
End Property
Public Property Let Title(val As String)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Title = val
    mudtProps.IsDirty = True
End Property
Public Property Get TitleF(pLen As Integer) As String
    TitleF = Left(Title, pLen) & IIf(Len(Title) > pLen, "...", "")
End Property
Public Property Get MainAuthor() As String
    MainAuthor = FNS(mudtProps.MainAuthor)
End Property
Public Property Let MainAuthor(val As String)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.MainAuthor = val
    mudtProps.IsDirty = True
End Property

Public Property Get DiscountPercent() As Double
    DiscountPercent = mudtProps.DiscountPercent
End Property
Public Property Let DiscountPercent(val As Double)
  If mcolStack.Count = 0 Then Err.Raise 383
    If Not NONStock Then  'non nonstock items cant have discount
        If mInvoiceParent.InvoiceDiscountRate > 0 Then 'we only allow discount here if no general discount applicable
            mudtProps.DiscountPercent = mInvoiceParent.InvoiceDiscountRate
        Else
            mudtProps.DiscountPercent = val
        End If
        If (val < 0) Then
            mobjValid.RuleBroken "DiscountPercent", True
        Else
            mobjValid.RuleBroken "DiscountPercent", False
        End If
        mudtProps.IsDirty = True
  '      RaiseEvent ExtensionChange(Extension, ExtensionF)
    End If
End Property
Public Function SetDiscountPercent(val As String) As Boolean
Dim dblDiscountPercent As Double
    SetDiscountPercent = True
    If Not NONStock Then  'non nonstock items cant have discount
        If Not mInvoiceParent Is Nothing Then
            If mInvoiceParent.InvoiceDiscountRate > 0 Then 'we only allow discount here if no general discount applicable
                mudtProps.DiscountPercent = mInvoiceParent.InvoiceDiscountRate
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
        If (dblDiscountPercent < 0) Then
            mobjValid.RuleBroken "DiscountPercent", True
        Else
            mobjValid.RuleBroken "DiscountPercent", False
        End If
        mudtProps.DiscountPercent = dblDiscountPercent
        mudtProps.IsDirty = True
    End If
End Function
Public Function DiscountPercentF() As String
    If mudtProps.DiscountPercent = 0 Then
        DiscountPercentF = ""
    Else
        DiscountPercentF = Format(mudtProps.DiscountPercent, gPercentFormatString)
    End If
End Function
Public Property Get VATRate() As Double
    VATRate = mudtProps.VATRate
End Property
Public Property Get VATRateF() As String
    VATRateF = Format(mudtProps.VATRate, gPercentFormatString)
End Property

Public Property Let VATRate(val As Double)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.VATRate = val
    mudtProps.IsDirty = True
End Property
Public Property Get NONStock() As Boolean
    NONStock = mudtProps.NONStock
End Property
Public Property Let NONStock(val As Boolean)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.NONStock = val
    mudtProps.IsDirty = True
End Property
'---------------------------------
Public Property Get VATAmount() As Long
    VATAmount = CLng(CDbl((VATRate / (100 + VATRate))) * ExtensionInclDeposit)
End Property
Public Property Get VATAmount_foreign() As Long
    VATAmount_foreign = CLng((VATRate / (100 + VATRate)) * ExtensionInclDeposit_Foreign)
End Property
Public Property Get VATAmountF(bForeign As Boolean) As String
    If bForeign Then
        VATAmountF = Format(CDbl(VATAmount / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        VATAmountF = Format(CDbl(VATAmount / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Public Property Get ExtensionExVAT() As Long
    ExtensionExVAT = (CLng((mudtProps.Price * mudtProps.Qty) * ((100 - mudtProps.DiscountPercent) / 100))) * (100 / (100 + mudtProps.VATRate))
End Property
Public Property Get ExtensionExVATF(bForeign As Boolean) As String
    If bForeign Then
        ExtensionExVATF = Format(CDbl(ExtensionExVAT / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        ExtensionExVATF = Format(CDbl(ExtensionExVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Public Property Get ExtensionInclDeposit() As Long
    ExtensionInclDeposit = CLng(CLng(mudtProps.Price * mudtProps.Qty) * ((100 - mudtProps.DiscountPercent) / 100))
End Property
Public Property Get ExtensionInclDeposit_Foreign() As Long
    ExtensionInclDeposit_Foreign = CLng(CLng(ForeignPrice * mudtProps.Qty) * ((100 - mudtProps.DiscountPercent) / 100))
End Property
Public Property Get Extension() As Long
    Extension = CLng(CLng(mudtProps.Price * mudtProps.Qty) * ((100 - mudtProps.DiscountPercent) / 100))
End Property
Public Property Get Extension_Foreign() As Long
    Extension_Foreign = CLng(CLng(ForeignPrice * mudtProps.Qty) * ((100 - mudtProps.DiscountPercent) / 100))
End Property
Public Property Get ForeignPrice() As Long
    ForeignPrice = mudtProps.ForeignPrice 'RoundUp(mudtProps.ForeignPrice / mInvoiceParent.ForeignCurrency.Divisor, oPC.Configuration.RoundingRules.GetRoundTo(mudtProps.ForeignPrice / mInvoiceParent.ForeignCurrency.Divisor)) * mInvoiceParent.ForeignCurrency.Divisor
End Property

Public Property Get ExtensionF(bForeign As Boolean) As String
    If bForeign Then
        ExtensionF = Format(CDbl(Extension_Foreign / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        ExtensionF = Format(CDbl(Extension / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Public Property Get ExtensionSimple() As Long
    ExtensionSimple = CLng(mudtProps.Price * mudtProps.Qty)
End Property
Public Property Get ExtensionSimple_foreign() As Long
    ExtensionSimple_foreign = CLng(mudtProps.ForeignPrice * mudtProps.Qty)
End Property
Public Property Get ExtensionSimpleF(bForeign As Boolean) As String
    If bForeign Then
        ExtensionSimpleF = Format(CDbl(ExtensionSimple_foreign / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        ExtensionSimpleF = Format(CDbl(ExtensionSimple / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Public Function PriceF(bForeign As Boolean) As String
    If bForeign Then
        PriceF = Format(CDbl(ForeignPrice / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        PriceF = Format(CDbl(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Function
Public Property Get DiscountAmount() As Long
    DiscountAmount = CLng(ExtensionSimple - ExtensionInclDeposit)
End Property
Public Property Get DiscountAmount_Foreign() As Long
    DiscountAmount_Foreign = CLng(ExtensionSimple_foreign - ExtensionInclDeposit_Foreign)
End Property
Public Property Get DiscountAmountF(bForeign As Boolean) As String
    If bForeign Then
        DiscountAmountF = Format(CDbl(DiscountAmount / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        DiscountAmountF = Format(CDbl(DiscountAmount / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property

Public Property Get Deposit_Foreign() As Long
    Deposit_Foreign = mudtProps.ForeignDeposit
End Property
Public Property Let Deposit_Foreign(val As Long)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ForeignDeposit = val
    mudtProps.IsDirty = True
End Property
Public Property Get Deposit() As Long
    Deposit = mudtProps.Deposit
End Property
Public Property Let Deposit(val As Long)
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.Deposit Then
        mudtProps.Deposit = val
        SetDirty True
    End If
End Property
Public Function SetDeposit(val As String) As Boolean
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Deposit
    bOK = SetField_strAsCurrencyToLong(mudtProps.Deposit, val, mcolStack.Count, "DEPOSIT")
    If bOK And (lngTemp <> mudtProps.Deposit) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("DEPOSIT")
    End If
    SetDeposit = bOK
End Function
Public Function DepositF(bForeign As Boolean) As String
    If bForeign Then
        DepositF = Format(CDbl(mudtProps.ForeignDeposit / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
    Else
        DepositF = Format(CDbl(mudtProps.Deposit / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
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
End Function

'------------------------------
'
'Public Property Get ForeignPrice() As Long
'    ForeignPrice = mudtProps.ForeignPrice
'End Property
Public Property Let ForeignPrice(val As Long)
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.ForeignPrice Then
        mudtProps.ForeignPrice = val
        SetDirty True
    End If
End Property
Public Property Get Price() As Long
    Price = mudtProps.Price
End Property
Public Property Let Price(val As Long)
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.Price Then
        mudtProps.Price = val
        SetDirty True
    End If
    ValidateObject "PRICE"
End Property
Public Function SetPrice(val As String) As Boolean
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Price
    bOK = SetField_strAsCurrencyToLong(mudtProps.Price, val, mcolStack.Count, "PRICE")
    If bOK And (lngTemp <> mudtProps.Price) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("PRICE")
    End If
    SetPrice = bOK
End Function
Public Function SetQty(val As String) As Boolean
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Qty
    bOK = SetField_LONG(mudtProps.Qty, val, "QTY", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.Qty) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QTY")
    End If
    SetQty = bOK
End Function

Private Function TranslateErrors(pRawErrors As String) As String
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
End Function

Private Sub SetDirty(pVal As Boolean)
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
End Sub

Private Sub LoadClassErrorsCollection()
    Set colClassErrors = New Collection
    colClassErrors.Add "Invalid price", "PRICE"
    colClassErrors.Add "Invalid quantity", "QTY"
    colClassErrors.Add "Invalid ETA", "ETA"
    colClassErrors.Add "Invalid discount", "DISCOUNT"
    colClassErrors.Add "Invalid deposit", "DEPOSIT"
End Sub
Private Function ValidateObject(pFld As String) As Boolean
On Error GoTo ERRH
Dim bValid As Boolean
    bValid = True
    If mudtProps.Price > 0 Then
        If pFld = "PRICE" Then bValid = False
        mobjValid.RuleBroken "PRICE", False
    Else
        mobjValid.RuleBroken "PRICE", True
    End If
    If mudtProps.Qty >= 0 Then
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
    mobjValid.GetSTatus
    SetDirty bValid
    ValidateObject = bValid
EXIT_Handler:
    Exit Function
ERRH:
    MsgBox "CheckOK: " & Error
End Function

Public Sub GetSTatus()
    ValidateObject ""
    mobjValid.GetSTatus
End Sub

Public Function CoffsValid(pResult As Integer)
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
End Function
