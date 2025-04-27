VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Sale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------------------------------------------

Event Valid(msg As String)
Event ProvisionalPrice()
Private mudtProps As SaleProps
Private mcolStack As Collection
Private mobjParent As ch_sale
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassErrors As Collection
Private mExchangeparent As a_Exchange
Private bNeedsRecalc As Boolean
Private bSpecialPrice As Boolean
Private mMID As Long
Private bNDA As Boolean
Private mKey As String
Private mNominalLineValue As Long
Private mLPLessDisc As Long
Private mLPLessDiscExt As Long
Private mLPDiscExt As Long
Private mLPLessDiscExtNonVAT As Long
Private mLPLessDiscExtVAT As Long
Private mLPDisc_Unit As Long
Private mLPDisc_Ext As Long

Private Type DiscountRule
    MID As Long
    NDA As Boolean
    Discount As Double
    DiscountDescription As String
    TotalValue As Long
    SP As Long
    SPSP As Long
    IDCust As Boolean
End Type
Dim Rules() As DiscountRule
Dim mSP As Long
Dim mSSP As Long

'Public Property Let Prices(val As adodb.Recordset)
'    Set rsPrices = val
'End Property
'Public Property Get Prices() As adodb.Recordset
'    Set Prices = rsPrices
'End Property


Public Property Get Properties(pSrt As enSortField) As Variant
    On Error GoTo errHandler
    Select Case pSrt
    Case enSequence
        Properties = Me.Sequence
    Case enTitle
        Properties = Me.title
    Case enAuthor
        Properties = Me.MainAuthor
    Case enCode
        Properties = Me.Code
    End Select
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.Properties(pSrt)", pSrt
End Property

Public Property Get Sequence() As Long
    On Error GoTo errHandler
    Sequence = mudtProps.Sequence
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.Sequence"
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
    ErrorIn "a_Sale.Sequence(val)", val
End Property

Public Function CalculateLine() As Boolean
    On Error GoTo errHandler
Dim bStoreAlteration As Boolean

    
    If Not mudtProps.IsDepositItem Then
        If mExchangeparent.Customer.CustomerType = oPC.GetLoyaltyCode Then
            'we must always give the actual discount if it is > 0
            If mudtProps.DiscountRate > 0 Then
                mudtProps.LoyaltyRate = 0
            End If
        End If
     '   If mExchangeparent.Customer.DefaultDiscount > 0 Then
        If mExchangeparent.Customer.DefaultDiscount > 0 And bSpecialPrice = False Then
            If mudtProps.DiscountRate > mExchangeparent.Customer.DefaultDiscount Then
                'Discount remains as is
            Else
                mudtProps.DiscountRate = mExchangeparent.Customer.DefaultDiscount
                mudtProps.DiscountRule = "Cust. std. disc."
                mudtProps.LoyaltyRate = 0
            End If
        End If
        If mudtProps.NoDiscountAllowed = True Then
                mudtProps.DiscountRate = 0
                mudtProps.DiscountRule = ""
                mudtProps.LoyaltyRate = 0
        End If
    End If
    mNominalLineValue = mudtProps.Price * mudtProps.Qty
   ' If Not FNS(mudtProps.MBCode) > "" Then
        mLPLessDisc = mudtProps.Price * ((100 - mudtProps.DiscountRate) / 100)
        If mudtProps.DiscountRate > 0 And bStoreAlteration = True Then
            mudtProps.PriceAlteration = mLPLessDisc - mudtProps.Price
        End If
    'End If
    mLPLessDiscExt = mLPLessDisc * mudtProps.Qty
    mLPDiscExt = (mudtProps.Price * mudtProps.Qty) - mLPLessDiscExt
    mLPLessDiscExtVAT = (mLPLessDiscExt - (mLPLessDiscExt / (100 + mudtProps.VATRate)) * 100)
    mLPLessDiscExtNonVAT = mLPLessDiscExt - mLPLessDiscExtVAT
    CalculateLine = True
    Exit Function
errHandler:
    ErrPreserve
    CalculateLine = False
    If Err = 6 Then
        mLPLessDiscExt = 999999
        mLPDiscExt = 999999
        mLPLessDiscExtVAT = 9999999
        mLPLessDiscExtNonVAT = 999999
        Exit Function
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.CalculateLine"
End Function
Public Property Get PNominalLineValue() As Long
    On Error GoTo errHandler
    
    PNominalLineValue = mudtProps.Price * mudtProps.Qty
    Exit Property
    
errHandler:
    ErrorIn "a_Sale.PNominalLineValue"
End Property
Public Property Get PLessDisc() As Long
        PLessDisc = mLPLessDisc
End Property
Public Property Get PLessDiscF() As String
    On Error GoTo errHandler
        PLessDiscF = Format(CDbl(mLPLessDisc / oPC.CurrencyDivisor), oPC.CurrencyFormat)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.PLessDiscF"
End Property

Public Property Get PLessDiscExt() As Long
        PLessDiscExt = mLPLessDiscExt
End Property

Public Property Get PLessDiscExtF() As String
    On Error GoTo errHandler
        PLessDiscExtF = Format(CDbl(mLPLessDiscExt / oPC.CurrencyDivisor), oPC.CurrencyFormat)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.PLessDiscExtF"
End Property

Public Property Get PDiscExt() As Long
        PDiscExt = FNN(mLPDiscExt)
End Property
Public Property Get PDiscExtF() As String
    On Error GoTo errHandler
        PDiscExtF = Format(CDbl(mLPDiscExt / oPC.CurrencyDivisor), oPC.CurrencyFormat)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.PDiscExtF"
End Property

Public Property Get PLessDiscExtVAT() As Long
        PLessDiscExtVAT = mLPLessDiscExtVAT
End Property
Public Property Get PLessDiscExtVATF() As String
    On Error GoTo errHandler
        PLessDiscExtVATF = Format(CDbl(mLPLessDiscExtVAT / oPC.CurrencyDivisor), oPC.CurrencyFormat)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.PLessDiscExtVATF"
End Property

Public Property Get PLessDiscExtNonVAT() As Long
        PLessDiscExtNonVAT = mLPLessDiscExtNonVAT
End Property
Public Property Get PLessDiscExtNonVATF() As String
    On Error GoTo errHandler
        PLessDiscExtNonVATF = Format(CDbl(mLPLessDiscExtNonVAT / oPC.CurrencyDivisor), oPC.CurrencyFormat)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.PLessDiscExtNonVATF"
End Property

Private Function GetState() As String
    On Error GoTo errHandler
  Dim mudtData As SaleData
  LSet mudtData = mudtProps
  GetState = mudtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.GetState"
End Function
    
Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim mudtData As SaleData
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
    GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
    SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.SetSuperState(buffer)", buffer
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 445
    mudtProps.Payable = Me.PLessDiscExtNonVAT()
    mudtProps.VAT = Me.PLessDiscExtVAT()
    If Not mobjParent Is Nothing Then
        mobjParent.AddSaleLine Me
        Set mobjParent = Nothing
    End If
    mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.ApplyEdit"
End Sub
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property
Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.CancelEdit"
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    LoadClassErrorsCollection
    mobjValid.RuleBroken "Qty", True
    mobjValid.RuleBroken "Price", True
    mudtProps.IsNew = True
    mudtProps.SaleGUID = CreateGUID
 '   Set moProd = New a_Product
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.Class_Initialize"
End Sub
Friend Sub Initialize(Parent As ch_sale)
    On Error GoTo errHandler
    If mudtProps.IsNew Then
        Set mobjParent = Parent
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.Initialize(Parent)", Parent
End Sub
Friend Sub SetParentExchange(pParentExchange As a_Exchange)
    On Error GoTo errHandler
    Set mExchangeparent = pParentExchange
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.SetParentExchange(pParentExchange)", pParentExchange
End Sub
Friend Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.IsValid"
End Property

Private Sub Class_Terminate()
10        On Error GoTo errHandler
20        If mobjParent Is Nothing Then
30          If Not mcolStack Is Nothing Then
40              If mcolStack.Count > 0 Then _
                    LogSaveToFile "a_sale:Class_Terminate: State stack is not empty"
60          End If
70        End If
80        Set mcolStack = Nothing
90        Exit Sub
errHandler:
100       If ErrMustStop Then Debug.Assert False: Resume
110       ErrorIn "a_Sale.Class_Terminate", , , , "Line", Array(Erl())
End Sub

Private Sub mobjValid_BrokenRule(errors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(errors))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.mobjValid_BrokenRule(errors)", errors
End Sub
Private Sub mobjValid_RuleUnbroken(errors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(errors))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.mobjValid_RuleUnbroken(errors)", errors
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.mobjValid_Status(pMsg)", pMsg
End Sub

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.Delete"
End Sub
Public Property Get BottomOfDocument() As Boolean
    On Error GoTo errHandler
    BottomOfDocument = (Right(Code, 2) = "_B")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.BottomOfDocument"
End Property
Public Property Let IsDepositItem(val As Boolean)
    On Error GoTo errHandler
  mudtProps.IsDepositItem = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.IsDepositItem(val)", val
End Property

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.IsDirty"
End Property


Public Property Get saleID() As Long
    On Error GoTo errHandler
    saleID = mudtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.SaleID"
End Property

Public Property Let saleID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.SaleID(val)", val
End Property
Public Property Get ExchangeID() As String
    On Error GoTo errHandler
  ExchangeID = mudtProps.EXCHANGE_GUID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.ExchangeID"
End Property
Public Property Get SaleGUID() As String
    SaleGUID = mudtProps.SaleGUID
End Property

Public Property Get PID() As String
    On Error GoTo errHandler
    PID = mudtProps.PID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.PID"
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
    ErrorIn "a_Sale.PID(val)", val
End Property
Public Property Get MBCode() As String
    On Error GoTo errHandler
    MBCode = FNS(mudtProps.MBCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.MBCode"
End Property
Public Property Let MBCode(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.MBCode = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.MBCode(val)", val
End Property

Public Property Get CatID() As Long
    On Error GoTo errHandler
    CatID = FNN(mudtProps.CatID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.CATID"
End Property
Public Property Let CatID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If mudtProps.CatID <> val Then
        mudtProps.CatID = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.CATID(val)", val
End Property
Public Property Get ActionSignatureID() As Long
    On Error GoTo errHandler
    ActionSignatureID = FNN(mudtProps.ActionSignatureID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.ActionSignatureID"
End Property
Public Property Let ActionSignatureID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If mudtProps.ActionSignatureID <> val Then
        mudtProps.ActionSignatureID = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.ActionSignatureID(val)", val
End Property


Public Property Get Qty() As Long
    On Error GoTo errHandler
    Qty = FNN(mudtProps.Qty)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.Qty"
End Property
Public Property Get QtyF() As String
    On Error GoTo errHandler
    If NonStock Then   'e.g. postage, insurance etc
        QtyF = ""
    Else
        QtyF = Format(Qty, "##0")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.QtyF"
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
    ErrorIn "a_Sale.Qty(val)", val
End Property

Public Property Get Counterfoil() As String
    On Error GoTo errHandler
    Counterfoil = FNS(mudtProps.Counterfoil)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.Counterfoil"
End Property
Public Property Let Counterfoil(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Counterfoil = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.Counterfoil(val)", val
End Property


Public Property Get Note() As String
    On Error GoTo errHandler
    Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.Note"
End Property
Public Property Let Note(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Note = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.Note(val)", val
End Property


Public Property Get Code() As String
    On Error GoTo errHandler
    Code = FNS(mudtProps.Code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.Code"
End Property
Sub TestEvent()
    On Error GoTo errHandler
RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.TestEvent"
End Sub
Public Property Let Code(val As String)
    On Error GoTo errHandler
Dim oCode As z_ProdCode_Short
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Code = val

  '  Set oCode = New z_ProdCode_Short
  '  oCode.Load val
  '  mudtProps.CodeF = oCode.FormatCode
    mudtProps.IsDirty = True
    Set oCode = Nothing
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.Code(val)", val
End Property

Public Property Get CodeF() As String
    On Error GoTo errHandler
    CodeF = FNS(mudtProps.CodeF)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.CodeF"
End Property
Public Property Let CodeF(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CodeF = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.CodeF(val)", val
End Property

'Public Property Get CodeForExport() As String
'  CodeForExport = FNS(mudtProps.CodeForExport)
'End Property
'Public Property Let CodeForExport(val As String)
'  If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.CodeForExport = val
'End Property
'Public Property Get TitleAuthorPublisher()
'Dim tmp As String
'
'    tmp = Title
'    If Len(tmp) > 0 Then
'        tmp = tmp & IIf(Len(MainAuthor) > 0, "(" & MainAuthor & ")", "")
'    End If
'    If Len(tmp) > 0 Then
'        tmp = tmp & IIf(Len(Publisher) > 0, "(" & Publisher & ")", "")
'    End If
'    TitleAuthorPublisher = tmp
'
'End Property
'Public Property Get TitleAuthorPublisherWithoutBRackets()
'Dim tmp As String
'
'    tmp = Title
'    If Len(tmp) > 0 Then
'        tmp = tmp & IIf(Len(MainAuthor) > 0, " " & MainAuthor & " ", "")
'    End If
'    If Len(tmp) > 0 Then
'        tmp = tmp & IIf(Len(Publisher) > 0, "(" & Publisher & ")", "")
'    End If
'    TitleAuthorPublisherWithoutBRackets = tmp
'
'End Property
'Public Property Get TitleAuthorPublisherL(pLen As Integer)
'    TitleAuthorPublisherL = Left(TitleAuthorPublisher, pLen) & IIf(Len(TitleAuthorPublisher) > pLen, "...", "")
'End Property
'Public Property Get SpecialVATMsg() As String
'    If Me.VATRate <> oPC.ZSession.VATRate And oPC.Configuration.IsVATRegion Then
'         SpecialVATMsg = "(" & VATRateF & " VAT) "
'    End If
'End Property
Public Property Get DiscountRule() As String
    On Error GoTo errHandler
    DiscountRule = FNS(mudtProps.DiscountRule)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.DiscountRule"
End Property
Public Property Let DiscountDescription(val As String)
    On Error GoTo errHandler
    mudtProps.DiscountDescription = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.DiscountDescription"
End Property

Public Property Get DiscountDescription() As String
    On Error GoTo errHandler
    DiscountDescription = FNS(mudtProps.DiscountDescription)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.DiscountDescription"
End Property
Public Property Get DiscountRuleF() As String
    On Error GoTo errHandler
    If FNS(mudtProps.DiscountRule) > "" Then
        DiscountRuleF = FNS(mudtProps.DiscountRule)
    Else
        DiscountRuleF = FNS(mudtProps.DiscountDescription)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.DiscountRuleF"
End Property
Public Property Let DiscountRule(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DiscountRule = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.DiscountRule(val)", val
End Property

Public Property Get title() As String
    On Error GoTo errHandler
    title = FNS(mudtProps.title)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.title"
End Property
Public Property Let title(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.title = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.title(val)", val
End Property
Public Property Get TitleF(pLen As Integer) As String
    On Error GoTo errHandler
    TitleF = Left(title, pLen) & IIf(Len(title) > pLen, "...", "")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.TitleF(pLen)", pLen
End Property
Public Property Get TitleIncDiscF(pLen As Integer) As String
    On Error GoTo errHandler
    If Me.DiscountRate > 0 Then
        TitleIncDiscF = Left(title, pLen - 10) & " (Disc " & Me.DiscountRateF & ")" & IIf(Len(title) > pLen - 10, "...", "")
    Else
        TitleIncDiscF = Left(title, pLen) & IIf(Len(title) > pLen, "...", "")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.TitleIncDiscF(pLen)", pLen
End Property

Public Property Get MainAuthor() As String
    On Error GoTo errHandler
    MainAuthor = FNS(mudtProps.Author)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.MainAuthor"
End Property
Public Property Let MainAuthor(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Author = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.MainAuthor(val)", val
End Property
Public Property Get NoDiscountAllowed() As Boolean
    On Error GoTo errHandler
    NoDiscountAllowed = FNB(mudtProps.NoDiscountAllowed)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.NoDiscountAllowed"
End Property
Public Property Let NoDiscountAllowed(val As Boolean)
    On Error GoTo errHandler
    mudtProps.NoDiscountAllowed = FNB(val)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.NoDiscountAllowed(val)", val
End Property
Public Property Get DiscountRate() As Double
    On Error GoTo errHandler
    DiscountRate = FNDBL(mudtProps.DiscountRate)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.DiscountRate"
End Property
Public Property Let DiscountRate(val As Double)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If Not NonStock Then  'non nonstock items can't have discount
      '  If mExchangeparent.DiscountRate > 0 Then 'we only allow discount here if no general discount applicable
      '      mudtProps.DiscountRate = mExchangeparent.DiscountRate
      '  Else
            mudtProps.DiscountRate = val
      '  End If
        If (val < 0) Then
            mobjValid.RuleBroken "DiscountRate", True
        Else
            mobjValid.RuleBroken "DiscountRate", False
        End If
        mudtProps.IsDirty = True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.DiscountRate(val)", val
End Property
Public Function SetDiscountRateDbl(val As Double, Descr As String) As Boolean
    On Error GoTo errHandler
Dim dblDiscountRate As Double
    SetDiscountRateDbl = True
    If Not NonStock Then  'non nonstock items cant have discount
        If Not mExchangeparent Is Nothing Then
            If mExchangeparent.DiscountRate > 0 Then 'we only allow discount here if no general discount applicable
                mudtProps.DiscountRate = mExchangeparent.DiscountRate
                Exit Function
            End If
        End If
        
        dblDiscountRate = val
        If (dblDiscountRate < 0) Then
            SetDiscountRateDbl = False
            mobjValid.RuleBroken "DiscountRate", True
        Else
            mobjValid.RuleBroken "DiscountRate", False
        End If
        mudtProps.DiscountRate = dblDiscountRate
        mudtProps.DiscountDescription = Descr
        mudtProps.IsDirty = True
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.SetDiscountRateDbl(val,Descr)", Array(val, Descr)
End Function
Public Function SetDiscountRate(val As String, Descr As String) As Boolean
    On Error GoTo errHandler
Dim dblDiscountRate As Double
    SetDiscountRate = True
    If Not NonStock Then  'non nonstock items cant have discount
        If Not mExchangeparent Is Nothing Then
            If mExchangeparent.DiscountRate > 0 Then 'we only allow discount here if no general discount applicable
                mudtProps.DiscountRate = mExchangeparent.DiscountRate
                Exit Function
            End If
        End If
        
        If mcolStack.Count = 0 Then Err.Raise 383
        If Trim$(val) = "" Then
            dblDiscountRate = 0
        ElseIf Not ConvertToDBL(val, dblDiscountRate) Then
            SetDiscountRate = False
            Exit Function
        End If
        If (dblDiscountRate < 0) Then
            mobjValid.RuleBroken "DiscountRate", True
        Else
            mobjValid.RuleBroken "DiscountRate", False
        End If
        mudtProps.DiscountRate = dblDiscountRate
        mudtProps.DiscountDescription = Descr
        mudtProps.IsDirty = True
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.SetDiscountRate(val,Descr)", Array(val, Descr)
End Function
Public Function DiscountRateF() As String
    On Error GoTo errHandler
    If mudtProps.DiscountRate = 0 Then
        DiscountRateF = ""
    Else
        DiscountRateF = PBKSPercentF(mudtProps.DiscountRate)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.DiscountRateF"
End Function
Public Property Get VATRate() As Double
    On Error GoTo errHandler
    VATRate = FNDBL(mudtProps.VATRate)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.VATRate"
End Property
Public Property Get VATRateF() As String
    On Error GoTo errHandler
    VATRateF = PBKSPercentF(mudtProps.VATRate)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.VATRateF"
End Property

Public Property Let VATRate(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.VATRate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.VATRate(val)", val
End Property
Public Property Let LoyaltyDiscount(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.LoyaltyRate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.LoyaltyDiscount(val)", val
End Property
Public Property Get LoyaltyDiscount() As Long
    On Error GoTo errHandler
    LoyaltyDiscount = mudtProps.LoyaltyRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.LoyaltyDiscount"
End Property

Public Property Get NonStock() As Boolean
    On Error GoTo errHandler
    NonStock = mudtProps.ServiceItem
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.NonStock"
End Property
Public Property Let NonStock(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ServiceItem = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.NonStock(val)", val
End Property

Public Function PriceAlteration() As Long
    On Error GoTo errHandler
        PriceAlteration = FNN(mudtProps.PriceAlteration)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.PriceAlteration"
End Function

Public Function PriceF() As String
    On Error GoTo errHandler
        PriceF = Format(CDbl(mudtProps.Price / oPC.CurrencyDivisor), oPC.CurrencyFormat)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.PriceF"
End Function

Public Function SetNote(val As String) As Boolean
    On Error GoTo errHandler
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.SetNote(val)", val
End Function
Public Function SetRef(val As String) As Boolean
    On Error GoTo errHandler
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.SetRef(val)", val
End Function
'Public Property Get IdentifyCustomer() As Boolean
'    IdentifyCustomer = FNB(mudtProps.IdentifyCustomer)
'End Property
Public Property Let SSP(val As Long)
    mSSP = val
End Property
Public Property Get SSP() As Long
    SSP = mSSP
End Property
Public Property Let SP(val As Long)
    mSP = val
End Property
Public Property Get SP() As Long
    SP = mSP
End Property
Public Property Get Price() As Long
    On Error GoTo errHandler
    Price = FNN(mudtProps.Price)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.Price"
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
    ErrorIn "a_Sale.Price(val)", val
End Property
Public Property Get TempPrice() As Long
    On Error GoTo errHandler
    TempPrice = mudtProps.TempPrice
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.TempPrice"
End Property
Public Property Let TempPrice(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.TempPrice Then
        mudtProps.TempPrice = val
        SetDirty True
    End If
    ValidateObject "PRICE"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.TempPrice(val)", val
End Property
Public Function SetPrice(val As String, Optional bListPrice As Boolean) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Price
    bOK = SetField_strAsCurrencyToLong(mudtProps.Price, val, mcolStack.Count, "PRICE", False, oPC.CurrencyDivisor)
    If bOK And (lngTemp <> mudtProps.Price) Then 'only validate if conversion is OK and the new value is other than the original
        If Not bListPrice And lngTemp > 0 Then ' don't mark as alteration if the price is Zero
            mudtProps.PriceAlteration = mudtProps.Price - lngTemp
        End If
        bOK = ValidateObject("PRICE")
    End If
    SetPrice = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.SetPrice(val,bListPrice)", Array(val, bListPrice)
End Function
Public Function SetQty(val As String, Optional bNegative As Boolean) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Qty
    bOK = SetField_LONG(mudtProps.Qty, val, "QTY", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.Qty) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QTY")
    End If
    SetQty = bOK
    If bOK Then
        If bNegative Then
            mudtProps.Qty = mudtProps.Qty * -1
        End If
    Else
        mudtProps.Qty = lngTemp
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.SetQty(val,bNegative)", Array(val, bNegative)
End Function
Public Property Get COLID() As Long
    On Error GoTo errHandler
    COLID = FNN(mudtProps.COLID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.COLID"
End Property
Public Property Let COLID(val As Long)
    On Error GoTo errHandler
    mudtProps.COLID = FNN(val)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.COLID(val)", val
End Property
Private Function TranslateErrors(pRawErrors As String) As String
    On Error GoTo errHandler
Dim strOut As String
Dim strRule, strAllRules As String
Dim NoMoreRules As Boolean
Dim iMarker, istart As Integer
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
            istart = iMarker + 1
            iMarker = InStr(istart, pRawErrors, ",")
            If iMarker > 0 Then
                strRule = colClassErrors(MID(pRawErrors, istart, iMarker - istart))
            Else
                strRule = colClassErrors(MID(pRawErrors, istart))
            End If
                
            strAllRules = strAllRules & vbCrLf & strRule
        Loop
    End If
    TranslateErrors = strAllRules
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.TranslateErrors(pRawErrors)", pRawErrors
End Function

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.SetDirty(pVal)", pVal
End Sub

Private Sub LoadClassErrorsCollection()
    On Error GoTo errHandler
    Set colClassErrors = New Collection
    colClassErrors.Add "Invalid price", "PRICE"
    colClassErrors.Add "Invalid quantity", "QTY"
    colClassErrors.Add "Invalid discount", "DISCOUNT"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.LoadClassErrorsCollection"
End Sub
Private Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean
    bValid = True
    If mudtProps.Price > 0 Then
        mobjValid.RuleBroken "PRICE", False
    Else
        mobjValid.RuleBroken "PRICE", True
        If pFld = "PRICE" Then bValid = False
    End If
    If mudtProps.Qty >= 0 And mudtProps.Qty < 100000 Then
        mobjValid.RuleBroken "QTY", False
    Else
        If pFld = "QTY" Then bValid = False
        mobjValid.RuleBroken "QTY", True
    End If
    If (mudtProps.DiscountRate < 0 Or mudtProps.DiscountRate > 100) Then
        If pFld = "DISCOUNT" Then bValid = False
        mobjValid.RuleBroken "DISCOUNT", True
    Else
        mobjValid.RuleBroken "DISCOUNT", False
    End If
    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.ValidateObject(pFld)", pFld
End Function

Public Sub GetStatus()
    On Error GoTo errHandler
    ValidateObject ""
    mobjValid.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale.GetSTatus"
End Sub

Public Function LoadRules(val As ADODB.Recordset, IsAppro As Boolean)
Dim rs As ADODB.Recordset
Dim i As Integer

    Set rs = val
    ReDim Rules(1 To rs.RecordCount)
    rs.MoveFirst
'Set properties that do not change per rule
    Me.PID = FNS(rs.Fields("TPID"))
    Me.title = FNS(rs.Fields("TTITLE"))
    Me.Code = FNS(rs.Fields("TCODE"))
    Me.CodeF = FNS(rs.Fields("TEAN"))
    If Me.Qty = 0 Then Me.Qty = 1
'    Me.Author = FNS(rs.Fields("TAUTHOR"))
    Me.VATRate = FNDBL(rs.Fields("TVATRATE"))
    Me.SP = FNN(rs.Fields("TSP"))
    Me.SSP = FNN(rs.Fields("TSSP"))
    Me.CatID = FNN(rs.Fields("TCATID"))
    Me.MBCode = FNS(rs.Fields("TMBCODE"))
    Me.NoDiscountAllowed = FNS(rs.Fields("TNDA"))
    i = 1
    Do While Not rs.EOF
        If IsAppro Then
            Rules(i).NDA = True
        Else
            Rules(i).NDA = IIf(rs.Fields("TNDA") = 0, False, True)
        End If
        Rules(i).MID = FNN(rs.Fields("TMID"))
        Rules(i).Discount = FNDBL(rs.Fields("TDISC"))
        Rules(i).DiscountDescription = FNS(rs.Fields("TDISCDESC"))
        Rules(i).IDCust = IIf(rs.Fields("TIDCUST") = 0, False, True)
        Rules(i).TotalValue = FNN(rs.Fields("TMINVALUE"))
        rs.MoveNext
        i = i + 1
    Loop
    rs.Close
    Set rs = Nothing
End Function

Public Function FindRule(lngTotalSaleValue As Long, bIDCust As Boolean)
Dim i As Integer
    'We must look through the rules bearing in mind the current total sale value
    'and determine which rule if any applies
    bSpecialPrice = False
    'This next code covers the situation where a rule that used to apply no longer does
    'e.g. the total value of the sale has reduced because an item was removed from the sale
    If mMID > 0 Then   'Set the Price to either SP or SSP - because whatever rule applied no longer does
        If SSP > 0 Then
            Price = SSP
            bSpecialPrice = True
        Else
            If SP > 0 Then
                If Price = 0 Then
                    Price = SP
                End If
            End If
        End If
        DiscountRate = 0
        DiscountDescription = ""
        bNDA = False
        mudtProps.NoDiscountAllowed = False
        bIDCust = False
        mMID = 0
    End If
    On Error Resume Next
    If (Not Rules) = True Then 'check if array is empty
        Exit Function
    End If
    
    For i = 1 To UBound(Rules)
        If Rules(i).TotalValue <= lngTotalSaleValue And Me.Qty > 0 Then
            'Check whether SSP exists and if so whether this rule yields a lower price
            'Else move to next rule
            If Me.SSP > 0 Then
                If Me.SP * ((100 - Rules(i).Discount) / 100) < Me.SSP Then 'Rule applies
                    Me.DiscountDescription = Rules(i).DiscountDescription
                    Me.DiscountRate = Rules(i).Discount
                    mMID = Rules(i).MID
                    bIDCust = Rules(i).IDCust
                    bNDA = Rules(i).NDA
                    Me.Price = SP '* ((100 - Rules(i).Discount) / 100)
                    bSpecialPrice = False
                    Exit For
                Else
                    Me.Price = Me.SSP
                    bNDA = True
                    bSpecialPrice = True
                    Exit For
                End If
            End If
            If Rules(i).Discount >= Me.DiscountRate Or Rules(i).NDA = True Then
                DiscountDescription = Rules(i).DiscountDescription
                DiscountRate = Rules(i).Discount
                mMID = Rules(i).MID
                bIDCust = Rules(i).IDCust
                bNDA = Rules(i).NDA
                mudtProps.NoDiscountAllowed = bNDA

                If Me.Price = 0 Then
                    Me.Price = SP 'SP or overridden price!!!
                End If
                Exit For
            End If
        Else
            If SP <> 0 And Me.PriceAlteration = 0 Then Me.Price = SP  'e.g. newspapers or inaccurately captured items
            bNeedsRecalc = True
            RaiseEvent ProvisionalPrice
        End If
    Next i
End Function

Public Property Get IsDiscountAllowed() As Boolean
    IsDiscountAllowed = Not bNDA
End Property

Public Property Get IsSpecialPrice() As Boolean
    IsSpecialPrice = FNB(bSpecialPrice)
End Property

Public Sub LogSalesline()
Dim oSaleP As New a_Sale_P

    oSaleP.LogSalesline Me.SaleGUID, Me.PID, Me.Price, Me.Qty
    
End Sub
