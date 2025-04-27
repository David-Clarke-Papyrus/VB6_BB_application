VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_QUL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Event Valid(msg As String)
'Event ExtensionChange(lngExtension As Long, strExtension As String)
Private mudtProps As QULProps
Private mcolStack As Collection
Private mobjParent As ch_QUL
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassors As Collection
Private mQuoteParent As a_QU
      Dim oGUID As guid

Private mKey As String
Private moProd As a_Product

Private mFPLessDisc As Long
Private mLPLessDisc As Long

Private mFExt As Long
Private mExt As Long
Private mExtS As Long
Private mFExtS As Long
Private mFDisc As Long
Private mDisc As Long

Private mFVAT As Long
Private mVAT As Long

Private mFExt_Ex As Long
Private mExt_Ex As Long

Private mFPDisc_Unit As Long
Private mLPDisc_Unit As Long

Private mFPDisc_Ext As Long
Private mLPDisc_Ext As Long
Private mSelected As Boolean

Public Property Let Selected(val As Boolean)
    mSelected = val
End Property
Public Property Get Selected() As Boolean
    Selected = mSelected
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
    ErrorIn "a_QUL.Properties(pSrt)", pSrt
End Property

Public Property Get Sequence() As Long
    On Error GoTo errHandler
    Sequence = mudtProps.Sequence
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Sequence"
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
    ErrorIn "a_QUL.Sequence(val)", val
End Property

Private Function SetForeignPriceAndDeposit()
    On Error GoTo errHandler
Dim tmpPrice As Long
Dim lngRoundTo As Long
Dim ConvertedPrice As Long
Dim oCurr As a_Currency
    mobjValid.RuleBroken "OVERFLOW", False
    tmpPrice = mudtProps.Price * mQuoteParent.CurrencyFactor
    If Me.ServiceItem Then
        mudtProps.ForeignPrice = tmpPrice
        mudtProps.ForeignDeposit = mudtProps.Deposit * mQuoteParent.CurrencyFactor
    Else
        lngRoundTo = oPC.Configuration.RoundingRules.GetRoundTo(tmpPrice)
        ConvertedPrice = (RoundUp(tmpPrice, lngRoundTo))
        mudtProps.ForeignPrice = ConvertedPrice
        mudtProps.ForeignDeposit = mudtProps.Deposit / mQuoteParent.CurrencyFactor
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
    ErrorIn "a_QUL.SetForeignPriceAndDeposit"
End Function
Public Function CalculateLine()
    On Error GoTo errHandler
    mobjValid.RuleBroken "OVERFLOW", False
'=========================================
'Qty * Price = mExtS
'mExtS Less disc = mExt
'mExtS less Vat = mExtS_Ex
'mExt less Vat = mExt_Ex
'mDisc = mExtS - mExt
'=========================================

'First in local currency
    mExtS = mudtProps.Price * mudtProps.Qty
    mExt = CLng(CDbl(mudtProps.Price * mudtProps.Qty) * CDbl((100 - (mudtProps.DiscountPercent))) / 100)
    mDisc = (mudtProps.Price * mudtProps.Qty) - mExt
    mVAT = (mExt - (CDbl(mExt) * 100 / (100 + mudtProps.VATRate)))
    mExt_Ex = mExt - mVAT

'Then in Foreign if appropriate
    If mQuoteParent Is Nothing Then Exit Function
    If Not mQuoteParent.ForeignCurrency Is Nothing Then
        If mQuoteParent.ForeignCurrency.ID <> oPC.Configuration.DefaultCurrency.ID Then
            SetForeignPriceAndDeposit
            mFExtS = mudtProps.ForeignPrice * mudtProps.Qty
            mFExt = mudtProps.ForeignPrice * mudtProps.Qty * ((100 - mudtProps.DiscountPercent) / 100)
            mFDisc = (mudtProps.ForeignPrice * mudtProps.Qty) - mFExt
            mFVAT = (mFExt - CLng((CCur(mFExt) * 100 / (100 + mudtProps.VATRate))))
            mFExt_Ex = mFExt - mFVAT
        End If
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
    ErrorIn "a_QUL.CalculateLine"
End Function
Public Property Get Ext(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        Ext = mFExt
    Else
        Ext = mExt
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Ext(val)", val
End Property

Public Property Get ExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        ExtF = Format(CDbl(mFExt / mQuoteParent.ForeignCurrency.Divisor), mQuoteParent.ForeignCurrency.FormatString)
    Else
        ExtF = Format(CDbl(mExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ExtF(val)", val
End Property

Public Property Get Disc(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        Disc = mFDisc
    Else
        Disc = mDisc
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Disc(val)", val
End Property
Public Property Get DiscF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        DiscF = Format(CDbl(mFDisc / mQuoteParent.ForeignCurrency.Divisor), mQuoteParent.ForeignCurrency.FormatString)
    Else
        DiscF = Format(CDbl(mDisc / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.DiscF(val)", val
End Property
Public Property Get ExtSF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        ExtSF = Format(CDbl(mFExtS / mQuoteParent.ForeignCurrency.Divisor), mQuoteParent.ForeignCurrency.FormatString)
    Else
        ExtSF = Format(CDbl(mExtS / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ExtSF(val)", val
End Property
Public Property Get ExtS(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        ExtS = mFExtS
    Else
        ExtS = mExtS
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ExtS(val)", val
End Property

Public Property Get Ext_Ex(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        Ext_Ex = mFExt_Ex
    Else
        Ext_Ex = mExt_Ex
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Ext_Ex(val)", val
End Property
Public Property Get Ext_ExFF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        Ext_ExFF = Format(CDbl(mFExt_Ex / mQuoteParent.ForeignCurrency.Divisor), mQuoteParent.ForeignCurrency.FormatString)
    Else
        Ext_ExFF = Format(CDbl(mExt_Ex / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Ext_ExFF(val)", val
End Property

Public Property Get VAT(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        VAT = mFVAT
    Else
        VAT = mVAT
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.VAT(val)", val
End Property
Public Property Get VATF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        VATF = Format(CDbl(mFVAT / mQuoteParent.ForeignCurrency.Divisor), mQuoteParent.ForeignCurrency.FormatString)
    Else
        VATF = Format(CDbl(mVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.VATF(val)", val
End Property

Private Function GetState() As String
    On Error GoTo errHandler
Dim mudtData As QULData
  
    LSet mudtData = mudtProps
    GetState = mudtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.GetState"
End Function
    
Private Sub SetState(buffer As String)
    On Error GoTo errHandler
Dim mudtData As QULData
  
    mudtData.buffer = buffer
    LSet mudtProps = mudtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler

    GetSuperState = GetState

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler

  SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.SetSuperState(buffer)", buffer
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 445
    mudtProps.SalesValue = Me.VAT(False)
    mudtProps.VATAmount = Me.Ext_Ex(False)
    If Not mobjParent Is Nothing Then
        mobjParent.AddQuoteLine Me
        Set mobjParent = Nothing
    End If
    mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ApplyEdit"
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
    ErrorIn "a_QUL.BeginEdit"
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
    ErrorIn "a_QUL.CancelEdit"
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    LoadClassorsCollection
    mobjValid.RuleBroken "Qty", True
    mobjValid.RuleBroken "Price", True
    mudtProps.IsNew = True
    Set moProd = New a_Product
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Class_Initialize"
End Sub
Friend Sub Initialize(Parent As ch_QUL)
    On Error GoTo errHandler
    Set oGUID = New guid

    If mudtProps.IsNew Then
        Set mobjParent = Parent
        mudtProps.ID = oGUID.Value
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Initialize(Parent)", Parent
End Sub
Friend Sub SetParentQuote(pParentQuote As a_QU)
    On Error GoTo errHandler
    Set mQuoteParent = pParentQuote
'    If Not mQuoteParent Is Nothing Then
'        If mQuoteParent.InvoiceDiscountRate > 0 Then
'            If Not Me.ServiceItem Then
'                mudtProps.DiscountPercent = mQuoteParent.InvoiceDiscountRate
'            End If
'        End If
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.SetParentQuote(pParentQuote)", pParentQuote
End Sub
Friend Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.IsValid"
End Property

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set moProd = Nothing
    If mobjParent Is Nothing Then
      If mcolStack.Count > 0 Then _
        Err.Raise vbObjectError + 1001, , "State stack is not empty"
    End If
    Set mcolStack = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.mobjValid_Status(pMsg)", pMsg
End Sub



Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Delete"
End Sub
Public Property Get BottomOfDocument() As Boolean
    On Error GoTo errHandler
    BottomOfDocument = (Right(code, 2) = "_B")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.BottomOfDocument"
End Property
Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.IsDirty"
End Property


Public Property Get QuoteLineID() As Long
    On Error GoTo errHandler
    QuoteLineID = mudtProps.QULID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.QuoteLineID"
End Property

Public Property Let QuoteID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.QUID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.QuoteID(val)", val
End Property
Public Property Get QuoteID() As Long
    On Error GoTo errHandler
  QuoteID = mudtProps.QUID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.QuoteID"
End Property
Public Property Get Product() As a_Product
    Set Product = moProd
End Property

Public Property Get PID() As String
    On Error GoTo errHandler
    PID = FNS(mudtProps.PID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.PID"
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
    ErrorIn "a_QUL.PID(val)", val
End Property

Public Property Get Qty() As Long
    On Error GoTo errHandler
    Qty = mudtProps.Qty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Qty"
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
    ErrorIn "a_QUL.QtyF"
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
    ErrorIn "a_QUL.Qty(val)", val
End Property

Public Property Get COLID() As Long
    On Error GoTo errHandler
    COLID = mudtProps.COLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.COLID"
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
    ErrorIn "a_QUL.COLID(val)", val
End Property

Public Property Get Publisher() As String
    On Error GoTo errHandler
    Publisher = stripCRLF(FNS(mudtProps.Publisher))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Publisher"
End Property
Public Property Let Publisher(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Publisher = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Publisher(val)", val
End Property


Public Property Get Note() As String
    On Error GoTo errHandler
    Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Note"
End Property
Public Property Let Note(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Note = stripCRLF(val)
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Note(val)", val
End Property

Public Property Get Ref() As String
    On Error GoTo errHandler
    Ref = stripCRLF(FNS(mudtProps.Ref))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Ref"
End Property
Public Property Let Ref(val As String)
    On Error GoTo errHandler
    mudtProps.Ref = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Ref(val)", val
End Property

Public Property Get code() As String
    On Error GoTo errHandler
    code = stripCRLF(FNS(mudtProps.code))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Code"
End Property
Public Property Get EAN() As String
    On Error GoTo errHandler
    EAN = stripCRLF(FNS(mudtProps.EAN))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.EAN"
End Property
Public Property Let EAN(val As String)
    On Error GoTo errHandler
    mudtProps.EAN = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.EAN(val)", val
End Property
Sub TestEvent()
    On Error GoTo errHandler
RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.TestEvent"
End Sub
Public Property Let code(val As String)
    On Error GoTo errHandler
Dim oSM As New z_StockManager
Dim oProdCode As z_ProdCode
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.code = val
    mudtProps.CodeF = FormatISBN13(val)
    
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
    ErrorIn "a_QUL.Code(val)", val
End Property
'Public Property Let CNLQty(val As Integer)
'    mudtProps.tmpCNLQty = val
'End Property
'Public Property Get CNLQty() As Integer
'    CNLQty = FNN(mudtProps.tmpCNLQty)
'End Property
'Public Sub SetMarker(bOn As Boolean)
'    mudtProps.Marker = bOn
'End Sub
'Public Property Get Marker() As Boolean
'    Marker = FNB(mudtProps.Marker)
'End Property
Public Property Get CodeF() As String
    On Error GoTo errHandler
    CodeF = FNS(mudtProps.CodeF)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.CodeF"
End Property
Public Property Let CodeF(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CodeF = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.CodeF(val)", val
End Property
Public Property Get CodeForExport() As String
    On Error GoTo errHandler
  CodeForExport = FNS(mudtProps.CodeForExport)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.CodeForExport"
End Property
Public Property Let CodeForExport(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CodeForExport = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.CodeForExport(val)", val
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
    ErrorIn "a_QUL.TitleAuthor"
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
    ErrorIn "a_QUL.TitleAuthorPublisher"
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
    ErrorIn "a_QUL.TitleAuthorPublisherWithoutBRackets"
End Property
Public Property Get TitleAuthorPublisherL(pLen As Integer)
    On Error GoTo errHandler
    TitleAuthorPublisherL = stripCRLF(Left(TitleAuthorPublisher, pLen) & IIf(Len(TitleAuthorPublisher) > pLen, "...", ""))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.TitleAuthorPublisherL(pLen)", pLen
End Property
Public Property Get SpecialVATMsg() As String
    On Error GoTo errHandler
    If Me.VATRate <> oPC.Configuration.VATRate And oPC.Configuration.IsVATRegion Then
         SpecialVATMsg = "(" & VATRateF & " VAT) "
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.SpecialVATMsg"
End Property
Public Property Get Title() As String
    On Error GoTo errHandler
    Title = stripCRLF(FNS(mudtProps.Title))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Title"
End Property
Public Property Let Title(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Title = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Title(val)", val
End Property
Public Property Get TitleF(pLen As Integer) As String
    On Error GoTo errHandler
    TitleF = Left(Title, pLen) & IIf(Len(Title) > pLen, "...", "")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.TitleF(pLen)", pLen
End Property
Public Property Get MainAuthor() As String
    On Error GoTo errHandler
    MainAuthor = stripCRLF(FNS(mudtProps.MainAuthor))
     '   tmp = tmp & IIf(Len(MainAuthor) > 0, " " & MainAuthor & " ", "")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.MainAuthor"
End Property
Public Property Let MainAuthor(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.MainAuthor = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.MainAuthor(val)", val
End Property

Public Property Get DiscountPercent() As Double
    On Error GoTo errHandler
    DiscountPercent = mudtProps.DiscountPercent
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.DiscountPercent"
End Property
Public Property Let DiscountPercent(val As Double)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If Not ServiceItem Then  'non ServiceItem items cant have discount
     '   If mQuoteParent.InvoiceDiscountRate > 0 Then 'we only allow discount here if no general discount applicable
     '       mudtProps.DiscountPercent = mQuoteParent.InvoiceDiscountRate
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
    ErrorIn "a_QUL.DiscountPercent(val)", val
End Property
Public Function SetDiscountPercent(val As String) As Boolean
    On Error GoTo errHandler
Dim dblDiscountPercent As Double
    SetDiscountPercent = True
    If Not ServiceItem Then  'non-stock items can't have discount
        If Not mQuoteParent Is Nothing Then
            If mQuoteParent.DiscountRate > 0 Then 'we only allow discount here if no general discount applicable
                mudtProps.DiscountPercent = mQuoteParent.DiscountRate
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
    ErrorIn "a_QUL.SetDiscountPercent(val)", val
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
    ErrorIn "a_QUL.DiscountPercentF"
End Function
Public Property Get VATRate() As Double
    On Error GoTo errHandler
    VATRate = mudtProps.VATRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.VATRate"
End Property
Public Property Get VATRateF() As String
    On Error GoTo errHandler
    VATRateF = Format(mudtProps.VATRate, gPercentFormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.VATRateF"
End Property

Public Property Let VATRate(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.VATRate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.VATRate(val)", val
End Property
Public Property Get ServiceItem() As Boolean
    On Error GoTo errHandler
    ServiceItem = mudtProps.ServiceItem
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ServiceItem"
End Property
Public Property Let ServiceItem(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ServiceItem = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ServiceItem(val)", val
End Property

Public Function PriceF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        PriceF = Format(CDbl(mudtProps.ForeignPrice / mQuoteParent.ForeignCurrency.Divisor), mQuoteParent.ForeignCurrency.FormatString)
    Else
        PriceF = Format(CDbl(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.PriceF(bForeign)", bForeign
End Function
Public Property Get PriceExVAT() As Long
    On Error GoTo errHandler
    PriceExVAT = Round(mudtProps.Price * (100 / (100 + IIf(mudtProps.VATRate = 0, oPC.Configuration.VATRate, mudtProps.VATRate))))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.PriceExVAT"
End Property

Public Function PriceExVatF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        'surely this wont be used
        PriceExVatF = Format(CDbl(mudtProps.ForeignPrice * (100 / (100 + oPC.Configuration.VATRate)) / mQuoteParent.ForeignCurrency.Divisor), mQuoteParent.ForeignCurrency.FormatString)
    Else
        PriceExVatF = Format(CDbl(Me.PriceExVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.PriceF(bForeign)", bForeign
End Function

Public Property Let SubstitutesAvailable(pVal As Boolean)
    mudtProps.SubstitutesAvailable = pVal
    SetDirty (True)
End Property
Public Property Get SubstitutesAvailable() As Boolean
    SubstitutesAvailable = FNB(mudtProps.SubstitutesAvailable)
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
    ErrorIn "a_QUL.SetDeposit(val)", val
End Function
Public Function DepositF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        DepositF = Format(CDbl(mudtProps.ForeignDeposit / mQuoteParent.ForeignCurrency.Divisor), mQuoteParent.ForeignCurrency.FormatString)
    Else
        DepositF = Format(CDbl(mudtProps.Deposit / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.DepositF(bForeign)", bForeign
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

Public Property Get Price() As Long
    On Error GoTo errHandler
    Price = mudtProps.Price
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.Price"
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
    ErrorIn "a_QUL.Price(val)", val
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
    ErrorIn "a_QUL.SetPrice(val)", val
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
    ErrorIn "a_QUL.SetQty(val)", val
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
    ErrorIn "a_QUL.TranslateErrors(pRawors)", pRawors
End Function

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.SetDirty(pVal)", pVal
End Sub

Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Invalid price", "PRICE"
    colClassors.Add "Invalid quantity", "QTY"
    colClassors.Add "Invalid ETA", "ETA"
    colClassors.Add "Invalid discount", "DISCOUNT"
    colClassors.Add "Invalid extra charge", "EXTRACHARGE"
    colClassors.Add "Invalid deposit", "DEPOSIT"
    colClassors.Add "A numeric value is too great for calculation", "OVERFLOW"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.LoadClassorsCollection"
End Sub
Private Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean
    bValid = True
    If mudtProps.Price >= 0 Then
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
    If mudtProps.ExtraCharge <= 0 And FNS(mudtProps.ExtraCode) > "" Then
        If pFld = "EXTRACHARGE" Then bValid = False
        mobjValid.RuleBroken "EXTRACHARGE", True
    Else
        mobjValid.RuleBroken "EXTRACHARGE", False
    End If
    If (mudtProps.DiscountPercent < 0 Or mudtProps.DiscountPercent > 100) Then
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
    ErrorIn "a_QUL.ValidateObject(pFld)", pFld
End Function

Public Sub GetStatus()
    On Error GoTo errHandler
    ValidateObject ""
    mobjValid.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.GetSTatus"
End Sub
'==============
Public Property Get ExtraVATRate() As Double
    On Error GoTo errHandler
    ExtraVATRate = mudtProps.ExtraVATRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ExtraVATRate"
End Property
Public Property Get ExtraVATRateF() As String
    On Error GoTo errHandler
    ExtraVATRateF = Format(mudtProps.ExtraVATRate, gPercentFormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ExtraVATRateF"
End Property

Public Property Let ExtraVATRate(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ExtraVATRate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ExtraVATRate(val)", val
End Property
Public Property Get ExtraVATF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
    Else
        ExtraVATF = Format((Me.ExtraCharge - (CDbl(Me.ExtraCharge) * 100 / (100 + Me.ExtraVATRate))) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ExtraVATRateF"
End Property

Public Property Get ExtraChargeF() As String
    On Error GoTo errHandler
    ExtraChargeF = Format(mudtProps.ExtraCharge / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ExtraChargeF"
End Property

Public Property Get ExtraCharge() As Long
    On Error GoTo errHandler
    ExtraCharge = mudtProps.ExtraCharge
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ExtraCharge"
End Property
Public Function SetExtraCharge(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.ExtraCharge
    bOK = SetField_LONG(mudtProps.ExtraCharge, val, "EXTRACHARGE", mcolStack.Count)
        bOK = ValidateObject("EXTRACHARGE")
    SetExtraCharge = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.SetExtraCharge(val)", val
End Function

Public Property Get ExtraChargeDescription() As String
    On Error GoTo errHandler
    ExtraChargeDescription = stripCRLF(FNS(mudtProps.ExtraChargeDescription))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ExtraChargeDescription"
End Property
Public Property Let ExtraChargeDescription(val As String)
    On Error GoTo errHandler
    If val <> mudtProps.ExtraChargeDescription Then
        mudtProps.ExtraChargeDescription = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ExtraChargeDescription(val)", val
End Property
Public Property Get ExtraPID() As String
    On Error GoTo errHandler
    ExtraPID = FNS(mudtProps.ExtraPID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ExtraPID"
End Property
Public Property Let ExtraPID(val As String)
    On Error GoTo errHandler
    mudtProps.ExtraPID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ExtraPID(val)", val
End Property

Public Property Get ExtraCode() As String
    On Error GoTo errHandler
    ExtraCode = FNS(mudtProps.ExtraCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ExtraCode"
End Property
Public Property Let ExtraCode(val As String)
    On Error GoTo errHandler
    mudtProps.ExtraCode = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ExtraCode(val)", val
End Property
Public Property Get CodeForEditing() As String
    If Not IsISBN10(mudtProps.code) And FNS(mudtProps.code) > "" Then
        CodeForEditing = FNS(mudtProps.code)
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
    
End Property
Public Function SetLineExtraProduct(pCode As String) As Boolean
    On Error GoTo errHandler
Dim oProd As a_Product
    
    Set oProd = New a_Product
    With oProd
        If .Load("", 0, Trim$(pCode)) <> 99 Then   'product found
            If .IsServiceItem Then
                Me.ExtraChargeDescription = .Title
                Me.ExtraPID = .PID
                Me.ExtraCode = .code
                Me.ExtraVATRate = .VATRate
                SetLineExtraProduct = True
            Else
                SetLineExtraProduct = False   'cant add a stock item here
            End If
        Else
            SetLineExtraProduct = False   'Not found
        End If
    End With
    mobjValid.GetStatus
    Set oProd = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.SetLineExtraProduct(pCODE)", pCode
End Function

Public Property Get ForeignPrice() As Long
    On Error GoTo errHandler
    ForeignPrice = mudtProps.ForeignPrice
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ForeignPrice"
End Property
Public Property Let ForeignPrice(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.ForeignPrice Then
        mudtProps.ForeignPrice = val
        SetDirty True
    End If
    ValidateObject "ForeignPrice"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.ForeignPrice(val)", val
End Property
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

Public Property Get ForeignPriceF() As String
Dim oFC As New a_Currency
    Set oFC = oPC.Configuration.Currencies.FindCurrencyByID(mudtProps.FCID)
    ForeignPriceF = Format(CDbl(mudtProps.ForeignPrice / oFC.Divisor), oFC.FormatString)
End Property

Public Function SetFCFactor(val As Double) As Boolean
Dim bOK As Boolean
Dim Temp As Double
    Temp = mudtProps.FCFactor
    bOK = SetField_DOUBLE(mudtProps.FCFactor, CStr(val), "FCFactor", mcolStack.Count)
    If bOK And (Temp <> mudtProps.FCFactor) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("FCFactor")
    End If
    SetFCFactor = bOK
    Exit Function

End Function

Public Property Get FCFactor() As Double
    FCFactor = mudtProps.FCFactor
End Property
Public Property Get FCFactorF() As String
    FCFactorF = Format(FCFactor, "##0.####")
End Property
Public Property Get FCID() As Long
    FCID = mudtProps.FCID
End Property
Public Property Let FCID(val As Long)
    mudtProps.FCID = val
End Property
Public Property Get FCFactorInvF() As String
    FCFactorInvF = Format(Round(1# / FCFactor, 3), "##0.###")
End Property

