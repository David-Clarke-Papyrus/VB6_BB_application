VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_DeliveryLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event Valid(msg As String)
Event ValueChanges()

Private mudtProps As DELLProps
Private mcolStack As Collection
Private mobjParent As ch_DeliveryLine
Private mKey As String
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private mDELParent As a_Delivery
Private colClassors As Collection
Private moProd As a_Product

Private mFPExt As Long
Private mLPExt As Long

Private mFPLessDisc As Long
Private mLPLessDisc As Long

Private mFPLessDiscExt As Long
Private mFPSHortLessDiscExt As Long
Private mLPLessDiscExt As Long
Private mLPSHortLessDiscExt As Long

Private mFPDiscExt As Long
Private mLPDiscExt As Long

Private mFPLessDiscExtNonVAT As Long
Private mLPLessDiscExtNonVAT As Long

Private mFPLessDiscExtVAT As Long
Private mLPLessDiscExtVAT As Long

Private mFPDisc_Unit As Long
Private mLPDisc_Unit As Long

Private mFPDisc_Ext As Long
Private mLPDisc_Ext As Long
Dim bPriceChanged As Boolean
Dim dblMargin As Double

Public Property Get Properties(pSrt As enSortField) As Variant
    On Error GoTo errHandler
    Select Case pSrt
'    Case enSequence
'        Properties = Me.Sequence
    Case enTitle
        Properties = Me.Title
'    Case enAuthor
'        Properties = Me.
    Case enCode
        Properties = Me.code
    End Select
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DEL.Properties(pSrt)", pSrt
End Property

Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property

Public Function RecalculateLine()
    On Error GoTo errHandler
'First in local currency
    mobjValid.RuleBroken "OVERFLOW", False
    If Me.DEL.ISForeignCurrency Then
        mudtProps.Price = CLng(CDbl(mudtProps.ForeignPrice) / Me.DEL.Supplier.DefaultCurrency.Factor)
    End If
    mLPExt = mudtProps.Price * (mudtProps.QtyFirm + mudtProps.QtySS + mudtProps.QtyShort)
    mLPLessDisc = mudtProps.Price * ((100 - mudtProps.Discount) / 100)
    mLPLessDiscExt = mLPLessDisc * (mudtProps.QtyFirm + mudtProps.QtySS + mudtProps.QtyShort) ' * ((100 - mudtProps.Discount) / 100)
    mLPSHortLessDiscExt = mLPLessDisc * (mudtProps.QtyShort)
    mLPDiscExt = (mudtProps.Price * (mudtProps.QtyFirm + mudtProps.QtySS + mudtProps.QtyShort)) - mLPLessDiscExt
    mLPLessDiscExtVAT = (mLPLessDiscExt - (CCur(mLPLessDiscExt) * 100 / (100 + mudtProps.VATRate)))
    mLPLessDiscExtNonVAT = mLPLessDiscExt - mLPLessDiscExtVAT

'Then in Foreign if appropriate
    If DEL Is Nothing Then Exit Function
    If DEL.ISForeignCurrency Then
    '    SetForeignPriceAndDeposit
        mFPExt = mudtProps.ForeignPrice * (mudtProps.QtyFirm + mudtProps.QtySS + mudtProps.QtyShort)
        mFPLessDisc = mudtProps.ForeignPrice * ((100 - mudtProps.Discount) / 100)
        mFPLessDiscExt = mFPLessDisc * (mudtProps.QtyFirm + mudtProps.QtySS + mudtProps.QtyShort) '* ((100 - mudtProps.Discount) / 100)
       ' mFPLessDiscExt = mFPLessDisc * (mudtProps.QtyFirm + mudtProps.QtySS + mudtProps.QtyShort)
        mFPSHortLessDiscExt = mFPLessDisc * (mudtProps.QtyShort)
        mFPDiscExt = (mudtProps.ForeignPrice * (mudtProps.QtyFirm + mudtProps.QtySS + mudtProps.QtyShort)) - mFPLessDiscExt
        mFPLessDiscExtVAT = (mFPLessDiscExt - (CCur(mFPLessDiscExt) * 100 / (100 + mudtProps.VATRate)))
        mFPLessDiscExtNonVAT = mFPLessDiscExt - mFPLessDiscExtVAT
    End If
    CalcCost
    CalculateMarginFromSP
  '  RaiseEvent ValueChanges
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'ov erflow
        mobjValid.RuleBroken "OVERFLOW", True
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.RecalculateLine"
    'Resume
End Function
Public Property Get PLessDisc(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PLessDisc = mFPLessDisc
    Else
        PLessDisc = mLPLessDisc
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PLessDisc(val)", val
End Property
Public Property Get PLessDiscF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscF = Format(CDbl(mFPLessDisc / DEL.CaptureCurrency.Divisor), DEL.CaptureCurrency.FormatString)
    Else
        PLessDiscF = Format(CDbl(mLPLessDisc / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PLessDiscF(val)", val
End Property


Public Property Get PExt(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PExt = mFPExt
    Else
        PExt = mLPExt
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PExt(val)", val
End Property

Public Property Get PExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PExtF = Format(CDbl(mFPExt / DEL.CaptureCurrency.Divisor), DEL.CaptureCurrency.FormatString)
    Else
        PExtF = Format(CDbl(mLPExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PExtF(val)", val
End Property



Public Property Get PLessDiscExt(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PLessDiscExt = mFPLessDiscExt
    Else
        PLessDiscExt = mLPLessDiscExt
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PLessDiscExt(val)", val
End Property

Public Property Get PLessDiscExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscExtF = Format(CDbl(mFPLessDiscExt / DEL.CaptureCurrency.Divisor), DEL.CaptureCurrency.FormatString)
    Else
        PLessDiscExtF = Format(CDbl(mLPLessDiscExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PLessDiscExtF(val)", val
End Property

Public Property Get PShortLessDiscExt(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PShortLessDiscExt = mFPSHortLessDiscExt
    Else
        PShortLessDiscExt = mLPSHortLessDiscExt
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PSHortLessDiscExt(val)", val
End Property

Public Property Get PSHortLessDiscExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PSHortLessDiscExtF = Format(CDbl(mFPSHortLessDiscExt / DEL.CaptureCurrency.Divisor), DEL.CaptureCurrency.FormatString)
    Else
        PSHortLessDiscExtF = Format(CDbl(mLPSHortLessDiscExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PSHortLessDiscExtF(val)", val
End Property




Public Property Get PDiscExt(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PDiscExt = mFPDiscExt
    Else
        PDiscExt = mLPDiscExt
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PDiscExt(val)", val
End Property
Public Property Get PDiscExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PDiscExtF = Format(CDbl(mFPDiscExt / DEL.CaptureCurrency.Divisor), DEL.CaptureCurrency.FormatString)
    Else
        PDiscExtF = Format(CDbl(mLPDiscExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PDiscExtF(val)", val
End Property

Public Property Get PLessDiscExtVAT(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PLessDiscExtVAT = mFPLessDiscExtVAT
    Else
        PLessDiscExtVAT = mLPLessDiscExtVAT
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PLessDiscExtVAT(val)", val
End Property
Public Property Get PLessDiscExtVATF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscExtVATF = Format(CDbl(mFPLessDiscExtVAT / DEL.CaptureCurrency.Divisor), DEL.CaptureCurrency.FormatString)
    Else
        PLessDiscExtVATF = Format(CDbl(mLPLessDiscExtVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PLessDiscExtVATF(val)", val
End Property

Public Property Get PLessDiscExtNonVAT(val As Boolean) As Long
    On Error GoTo errHandler
    If val Then
        PLessDiscExtNonVAT = mFPLessDiscExtNonVAT
    Else
        PLessDiscExtNonVAT = mLPLessDiscExtNonVAT
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PLessDiscExtNonVAT(val)", val
End Property
Public Property Get PLessDiscExtNonVATF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscExtNonVATF = Format(CDbl(mFPLessDiscExtNonVAT / DEL.CaptureCurrency.Divisor), DEL.CaptureCurrency.FormatString)
    Else
        PLessDiscExtNonVATF = Format(CDbl(mLPLessDiscExtNonVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PLessDiscExtNonVATF(val)", val
End Property


Private Function GetState() As String
    On Error GoTo errHandler
Dim mudtData As DELLData
  
    LSet mudtData = mudtProps
    GetState = mudtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
Dim mudtData As DELLData
  
    mudtData.buffer = buffer
    LSet mudtProps = mudtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
    GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.SetSuperState(buffer)", buffer
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddDELLine Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.ApplyEdit"
End Sub

Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.BeginEdit"
End Sub
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DELL.IsEditing"
End Property

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
    ErrorIn "a_DeliveryLine.CancelEdit"
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    LoadClassorsCollection
    mudtProps.IsNew = True
    mudtProps.PIID = 0
    mudtProps.CodeFForExport = ""
    mudtProps.POLID = 0
    mudtProps.DiscountedPrice = 0
    mudtProps.IsNew = True
    mudtProps.IsNew = True
    mudtProps.IsNew = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Class_Initialize"
End Sub
Public Sub ClearValues()
    mudtProps.ClaimID = 0
    mudtProps.code = ""
    mudtProps.COLID = 0
    mudtProps.CorrectedDiscount = 0
    mudtProps.CorrectedForeignPrice = 0
    mudtProps.CorrectedPrice = 0
    mudtProps.Cost = 0
    mudtProps.Discount = 0
    mudtProps.Note = ""
    mudtProps.POL_Discount = 0
    mudtProps.POL_ForeignPrice = 0
    mudtProps.POL_Price = 0
    mudtProps.POL_QtyFirm = 0
    mudtProps.POL_QtySS = 0
    mudtProps.POLID = 0
    mudtProps.Price = 0
    mudtProps.PriceSell = 0
    mudtProps.ProductTypeID = 0
    mudtProps.QtyFirm = 1
    mudtProps.QtyShort = 0
    mudtProps.QtySS = 0
    mudtProps.QtyTotal = 0
    mudtProps.ReasonID = 0
    mudtProps.Ref = 0
    mudtProps.ReasonID = 0
    mudtProps.Section = ""
    mudtProps.Title = ""
    mudtProps.VATRate = oPC.Configuration.VATRate
End Sub
Friend Sub Initialize(Parent As ch_DeliveryLine)
    On Error GoTo errHandler
    If mudtProps.IsNew Then Set mobjParent = Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Initialize(Parent)", Parent
End Sub

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.IsValid"
End Property

Private Sub Class_Terminate()
    On Error GoTo errHandler
  If mobjParent Is Nothing Then
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
  End If
  Set mcolStack = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.mobjValid_BrokenRule(ors)", oRS
End Sub

'Private Sub mobjValid_NoBrokenRules()
'    On Error GoTo ErrHandler
'    RaiseEvent Valid("")
'    Exit Sub
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_DeliveryLine.mobjValid_NoBrokenRules"
'End Sub
'Private Sub mobjValid_BrokenRule(ors As String)
'    On Error GoTo ErrHandler
'    RaiseEvent Valid(TranslateErrors(ors))
'    Exit Sub
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_DeliveryLine.mobjValid_BrokenRule(ors)", ors
'End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.mobjValid_Status(pMsg)", pMsg
End Sub

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Delete"
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.IsDeleted"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.IsDirty"
End Property
'Public Property Get PriceF() As String
'    PriceF = Format(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get Price() As Long
'    Price = mudtProps.Price
'End Property
'Public Property Let Price(val As Long)
'  If mcolStack.Count = 0 Then Err.Raise 383
'    If val <> mudtProps.Price Then
'        mudtProps.Price = val
'        SetDirty True
'    End If
'End Property
'Public Function SetPrice(val As String) As Boolean
'Dim bOK As Boolean
'Dim bCalcOK As Boolean
'Dim lngTemp As Long
''MsgBox "here use alternatuive foreign"
'    If mDELParent.ISForeignCurrency Then
'        lngTemp = mudtProps.ForeignPrice
'        bOK = SetField_LONG(mudtProps.ForeignPrice, val, "FOREIGNPRICE", mcolStack.Count)
'        bCalcOK = CalculateReceivedPriceFromForeign(mudtProps.ForeignPrice, mudtProps.Price, mDELParent.CaptureCurrency.Factor)
'        bOK = ValidateObject("FOREIGNPRICE")
'        SetPrice = bOK And bCalcOK
'    Else
'        lngTemp = mudtProps.Price
'        bOK = SetField_LONG(mudtProps.Price, val, "PRICE", mcolStack.Count)
'            bOK = ValidateObject("PRICE")
'        SetPrice = bOK
'    End If
'    mDELParent.ValidateObject ""
'
'End Function
'Private Function CalculateReceivedPriceFromForeign(lngForeign As Long, lngLocal As Long, snglFactor As Single) As Boolean
'Dim bOK As Boolean
'    bOK = True
'    lngLocal = lngForeign * snglFactor
'    CalculateReceivedPriceFromForeign = bOK
'End Function
'Private Function CalculateLocalPriceFromForeign()
'
'End Function
'
''
'Public Property Get ForeignPriceF() As String
'    ForeignPriceF = Format(mudtProps.ForeignPrice / mDELParent.CaptureCurrency.Divisor, mDELParent.CaptureCurrency.FormatString)
'End Property
'Public Property Get ForeignPrice() As Long
'    ForeignPrice = mudtProps.ForeignPrice
'End Property
'Public Property Let ForeignPrice(val As Long)
'  If mcolStack.Count = 0 Then Err.Raise 383
'    If val <> mudtProps.ForeignPrice Then
'        mudtProps.ForeignPrice = val
'        SetDirty True
'    End If
'   RecalculateLine
'End Property
'Public Function SetForeignPrice(val As String) As Boolean
'Dim bOK As Boolean
'Dim lngTemp As Long
'    lngTemp = mudtProps.ForeignPrice
'    bOK = SetField_LONG(mudtProps.ForeignPrice, val, "FOREIGNPRICE", mcolStack.Count)
'        bOK = ValidateObject("FOREIGNPRICE")
'    SetForeignPrice = bOK
'   RecalculateLine
'End Function
'
'Public Function SetPrice(val As String) As Boolean
'Dim bOK As Boolean
'Dim lngTemp As Long
'    If Not .CaptureCurrency Is oPC.Configuration.DefaultCurrency Then  'The PO is being captured in a foreign currency and both prices nust be saved
'        lngTemp = mudtProps.ForeignPrice
'        bOK = SetField_strAsCurrencyToLong(mudtProps.ForeignPrice, val, mcolStack.Count, "PRICE")
'        mudtProps.Price = mudtProps.ForeignPrice * PO.CaptureCurrency.Factor
'        If bOK And (lngTemp <> mudtProps.ForeignPrice) Then 'only validate if conversion is OK and the new value is other than the original
'            bOK = ValidateObject("PRICE")
'        End If
'    Else
'        lngTemp = mudtProps.Price
'        bOK = SetField_strAsCurrencyToLong(mudtProps.Price, val, mcolStack.Count, "PRICE")
'        If bOK And (lngTemp <> mudtProps.Price) Then 'only validate if conversion is OK and the new value is other than the original
'            bOK = ValidateObject("PRICE")
'        End If
'    End If
'    SetPrice = bOK
'    RecalculateLine
'End Function

Public Property Get PriceF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PriceF = Format(CDbl(mudtProps.ForeignPrice / DEL.CaptureCurrency.Divisor), DEL.CaptureCurrency.FormatString)
    Else
        PriceF = Format(CDbl(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PriceF(val)", val
End Property

Public Property Get Price(bForeign As Boolean) As Long
    On Error GoTo errHandler
    If bForeign Then
        Price = mudtProps.ForeignPrice
    Else
        Price = mudtProps.Price
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Price(bForeign)", bForeign
End Property

Public Property Get CorrectedPrice(bForeign As Boolean) As Long
    On Error GoTo errHandler
    If bForeign Then
        CorrectedPrice = mudtProps.CorrectedForeignPrice
    Else
        CorrectedPrice = mudtProps.CorrectedPrice
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.CorrectedPrice(bForeign)", bForeign
End Property
Public Function SetCost(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Cost
    bOK = SetField_strAsCurrencyToLong(mudtProps.Cost, val, mcolStack.Count, "COST", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
    If bOK And (lngTemp <> mudtProps.Cost) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("COST")
    End If
    Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.SetCost(Val)", val
End Function
Public Property Get Cost() As Long
    On Error GoTo errHandler
    Cost = mudtProps.Cost
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Cost"
End Property
Public Property Get CostF() As String
    On Error GoTo errHandler
    CostF = Format(CDbl(mudtProps.Cost / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Cost"
End Property

Public Function SetPriceSell(val As String) As Boolean
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.PriceSell
    bOK = SetField_strAsCurrencyToLong(mudtProps.PriceSell, val, mcolStack.Count, "PRICESELL", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
    If bOK And (lngTemp <> mudtProps.PriceSell) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("PRICESELL")
    End If
    CalculateMarginFromSP

End Function

Public Property Get PriceSell() As Long
    On Error GoTo errHandler
    PriceSell = mudtProps.PriceSell
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PriceSell"
End Property
Public Property Get PriceSell_string() As String
    On Error GoTo errHandler
    PriceSell_string = CStr(mudtProps.PriceSell)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PriceSell"
End Property

Public Property Get PriceSellF() As String
    On Error GoTo errHandler
    PriceSellF = Format(CDbl(mudtProps.PriceSell / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
   ' PriceSellF = Format(CDbl(mudtProps.PriceSell), oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PriceSell"
End Property

Public Property Get Discount() As Double
    On Error GoTo errHandler
    Discount = mudtProps.Discount
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Discount"
End Property

Public Property Let Discount(Value As Double)
    On Error GoTo errHandler
    mudtProps.Discount = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Discount(Value)", Value
End Property
Public Property Get TRID() As Long
    On Error GoTo errHandler
    TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.TRID"
End Property

Public Property Let TRID(Value As Long)
    On Error GoTo errHandler
    mudtProps.TRID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.TRID(Value)", Value
End Property
Public Property Get COLID() As Long
    On Error GoTo errHandler
    COLID = mudtProps.COLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.COLID"
End Property
Public Property Get EAN() As String
    On Error GoTo errHandler
    EAN = FNS(mudtProps.EAN)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.EAN"
End Property
Public Property Let EAN(val As String)
    On Error GoTo errHandler
    mudtProps.EAN = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.EAN"
End Property
Public Property Let COLID(Value As Long)
    On Error GoTo errHandler
    mudtProps.COLID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.COLID(Value)", Value
End Property
Public Property Get DELLID() As Long
    On Error GoTo errHandler
    DELLID = mudtProps.DELLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.DELLID"
End Property

Public Property Let DELLID(Value As Long)
    On Error GoTo errHandler
    mudtProps.DELLID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.DELLID(Value)", Value
End Property
    
Public Property Get Title() As String
    On Error GoTo errHandler
    Title = FNS(mudtProps.Title)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Title"
End Property

Public Property Let Title(Value As String)
    On Error GoTo errHandler
    mudtProps.Title = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Title(Value)", Value
End Property

Public Property Get CodeF() As String
    On Error GoTo errHandler
    CodeF = FNS(mudtProps.CodeF)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.CodeF"
End Property

Public Property Let CodeF(Value As String)
    On Error GoTo errHandler
    mudtProps.CodeF = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.CodeF(Value)", Value
End Property

Public Property Get CodeForExport() As String
    On Error GoTo errHandler
    CodeForExport = FNS(mudtProps.CodeFForExport)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.CodeForExport"
End Property

Public Property Let CodeForExport(Value As String)
    On Error GoTo errHandler
    mudtProps.CodeFForExport = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.CodeForExport(Value)", Value
End Property

Public Property Get code() As String
    On Error GoTo errHandler
    code = FNS(mudtProps.code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Code"
End Property
Public Property Get Ref() As String
    On Error GoTo errHandler
    Ref = FNS(mudtProps.Ref)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Ref"
End Property
Public Property Let Ref(val As String)
    On Error GoTo errHandler
    mudtProps.Ref = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Ref(val)", val
End Property
Public Function SetRef(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim strTemp As String
    strTemp = mudtProps.Ref
    mudtProps.Ref = val
    If strTemp <> mudtProps.Ref Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("REF")
    End If
    SetRef = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.SetRef(val)", val
End Function
Public Property Get Note() As String
    On Error GoTo errHandler
    Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Note"
End Property
Public Property Let Note(val As String)
    On Error GoTo errHandler
    mudtProps.Note = stripCRLF(val)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Note(val)", val
End Property
Public Function SetNote(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim strTemp As String
    strTemp = Trim(mudtProps.Note)
    mudtProps.Note = val
    If strTemp <> mudtProps.Note Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("NOTE")
    End If
    SetNote = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.SetNote(val)", val
End Function

Public Property Let code(val As String)
    On Error GoTo errHandler
Dim oProdCode As z_ProdCode
    mudtProps.code = val
    
    Set oProdCode = New z_ProdCode
    If Len(val) = 13 Then
        oProdCode.LoadNew "", val, mudtProps.IsNew, , "", True
    Else
        oProdCode.LoadNew val, "", mudtProps.IsNew, , "", True
    End If
    mudtProps.CodeF = oProdCode.FormatCode(False)
    mudtProps.CodeFForExport = oProdCode.FormatCode(True)
    Set oProdCode = Nothing
    
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.Code(Val)", val
End Property
    
Public Property Get POLID() As Long
    On Error GoTo errHandler
    POLID = mudtProps.POLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.POLID"
End Property

Public Property Let POLID(Value As Long)
    On Error GoTo errHandler
'    If Value = 0 Then
'        mDELParent.POLsOSPersSUPP.FindByPOLID(mudtProps.POLID).Applied = False
'    Else
'        mDELParent.POLsOSPersSUPP.FindByPOLID(Value).Applied = True
'    End If
    mudtProps.POLID = Value
        
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.POLID(Value)", Value
End Property

Public Property Get PID() As String
    On Error GoTo errHandler
    PID = FNS(mudtProps.PID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PID"
End Property

Public Property Let PID(Value As String)
    On Error GoTo errHandler
    mudtProps.PID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.PID(Value)", Value
End Property
Public Property Get VATRate() As Double
    On Error GoTo errHandler
    VATRate = mudtProps.VATRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.VATRate"
End Property
Public Property Get VATRateF() As String
    On Error GoTo errHandler
    VATRateF = Format(mudtProps.VATRate, gPercentFormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.VATRateF"
End Property

Public Property Let VATRate(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.VATRate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.VATRate(val)", val
End Property
Public Property Get POLQtySS() As Long
    On Error GoTo errHandler
    POLQtySS = mudtProps.POL_QtySS
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.POLQtySS"
End Property
Public Property Get POLQtyFirm() As Long
    On Error GoTo errHandler
    POLQtyFirm = mudtProps.POL_QtyFirm
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.POLQtyFirm"
End Property
Public Property Get POLPrice() As Long
    On Error GoTo errHandler
    POLPrice = mudtProps.POL_Price
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.POLPrice"
End Property
Public Property Get POLPriceF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        POLPriceF = Format(mudtProps.POL_ForeignPrice / Me.DEL.CaptureCurrency.Divisor, Me.DEL.CaptureCurrency.FormatString)
    Else
        POLPriceF = Format(mudtProps.POL_Price / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.POLPriceF(bForeign)", bForeign
End Property
Public Property Get POLDiscount() As Double
    On Error GoTo errHandler
    POLDiscount = mudtProps.POL_Discount
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.POLDiscount"
End Property
Public Property Get POLDiscountF() As String
    On Error GoTo errHandler
    POLDiscountF = PBKSPercentF(mudtProps.POL_Discount)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.POLDiscountF"
End Property
Public Property Get POCode() As String
    On Error GoTo errHandler
    POCode = FNS(mudtProps.PO_Code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.POCode"
End Property

Public Function SetQtyShort(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.QtyShort
    bOK = SetField_LONG(mudtProps.QtyShort, val, "QtyShort", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.QtyShort) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QtyShort")
    End If
    SetQtyShort = bOK
    RecalculateLine
    mDELParent.ValidateObject ""
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.SetQtyShort(val)", val
End Function
Public Property Get QtyShort() As Long
    On Error GoTo errHandler
    QtyShort = mudtProps.QtyShort
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.QtyShort"
End Property
Public Property Get QtyShortF() As String
    On Error GoTo errHandler
    QtyShortF = Format(mudtProps.QtyShort, "##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.QtyShortF"
End Property

'=================
Public Function SetCorrectedDiscount(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.CorrectedDiscount
    bOK = SetField_DOUBLE(mudtProps.CorrectedDiscount, val, "CorrectedDiscount", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.CorrectedDiscount) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("CorrectedDiscount")
    End If
    SetCorrectedDiscount = bOK
    RecalculateLine
    mDELParent.ValidateObject ""
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.SetCorrectedDiscount(val)", val
End Function
Public Property Get CorrectedDiscount() As Long
    On Error GoTo errHandler
    CorrectedDiscount = mudtProps.CorrectedDiscount
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.CorrectedDiscount"
End Property
Public Property Get CorrectedDiscountF() As String
    On Error GoTo errHandler
    CorrectedDiscountF = Format(mudtProps.CorrectedDiscount, "##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.CorrectedDiscountF"
End Property
Public Function SetCorrectedPrice(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
Dim tmpPrice As Long
Dim lngRoundTo As Long
    
    bPriceChanged = False
    If val = "" Then val = "0"
    If Not DEL.CaptureCurrency Is oPC.Configuration.DefaultCurrency Then
        lngTemp = mudtProps.CorrectedForeignPrice
        bOK = SetField_strAsCurrencyToLong(mudtProps.CorrectedForeignPrice, val, mcolStack.Count, "PRICE", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
        
        If bOK And (lngTemp <> mudtProps.CorrectedForeignPrice) Then 'only validate if conversion is OK and the new value is other than the original
            bOK = ValidateObject("PRICE")
        End If
    Else
        lngTemp = mudtProps.CorrectedPrice
        bOK = SetField_strAsCurrencyToLong(mudtProps.CorrectedPrice, val, mcolStack.Count, "PRICE", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
        If bOK And (lngTemp <> mudtProps.CorrectedPrice) Then 'only validate if conversion is OK and the new value is other than the original
            bPriceChanged = True
            bOK = ValidateObject("PRICE")
        End If
    End If
    If bOK Then
        SetCorrectedPrice = bOK
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.SetPrice(val)", val

End Function
Public Property Let ReasonID(val As String)
    mudtProps.ReasonID = val
End Property

Public Property Get ReasonID() As String
    ReasonID = FNS(mudtProps.ReasonID)
End Property

Public Property Let ClaimID(val As Long)
    mudtProps.ClaimID = val
End Property

Public Property Get ClaimID() As Long
    ClaimID = FNS(mudtProps.ClaimID)
End Property

Public Function SetQtySS(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.QtySS
    bOK = SetField_LONG(mudtProps.QtySS, val, "QTYSS", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.QtySS) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QTYSS")
    End If
    SetQtySS = bOK
    RecalculateLine
    mDELParent.ValidateObject ""
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.SetQtySS(val)", val
End Function
Public Property Get QtyTotal() As Long
    On Error GoTo errHandler
    QtyTotal = mudtProps.QtyTotal
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.QtyTotal"
End Property

Public Property Get QtySS() As Long
    On Error GoTo errHandler
    QtySS = mudtProps.QtySS
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.QtySS"
End Property
Public Property Get QtySSF() As String
    On Error GoTo errHandler
    QtySSF = Format(mudtProps.QtySS, "##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.QtySSF"
End Property
Public Function SetQtyFirm(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.QtyFirm
    bOK = SetField_LONG(mudtProps.QtyFirm, val, "QTYFIRM", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.QtyFirm) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QTYFIRM")
    End If
    SetQtyFirm = bOK
    RecalculateLine
    mDELParent.ValidateObject ""
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.SetQtyFirm(val)", val
End Function
Public Sub ForceValidation()
     ValidateObject ""

End Sub
Public Property Get QtyFirm() As Long
    On Error GoTo errHandler
    QtyFirm = mudtProps.QtyFirm
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.QtyFirm"
End Property
Public Property Get QtyFirmF() As String
    On Error GoTo errHandler
    QtyFirmF = Format(mudtProps.QtyFirm, "##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.QtyFirmF"
End Property
Public Property Get POCurrID() As Long
    On Error GoTo errHandler
    POCurrID = mudtProps.PO_CURRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.POCurrID"
End Property
Public Property Get POCurrRate() As Double
    On Error GoTo errHandler
    POCurrRate = mudtProps.PO_CURRRATE
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.POCurrRate"
End Property
Public Property Get POLForignPrice() As Long
    On Error GoTo errHandler
    POLForignPrice = mudtProps.POL_ForeignPrice
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.POLForignPrice"
End Property
Public Property Get POLForignPriceF() As String
    On Error GoTo errHandler
    POLForignPriceF = Format(mudtProps.POL_ForeignPrice / Me.DEL.CaptureCurrency.Divisor, Me.DEL.CaptureCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.POLForignPriceF"
End Property

Public Function SetPrice(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
Dim tmpPrice As Long
Dim lngRoundTo As Long
    
    bPriceChanged = False
    If val = "" Then val = "0"
    If Not DEL.CaptureCurrency Is oPC.Configuration.DefaultCurrency Then
        lngTemp = mudtProps.ForeignPrice
        bOK = SetField_strAsCurrencyToLong(mudtProps.ForeignPrice, val, mcolStack.Count, "PRICE", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
        If oPC.SupplierBasedCurrencyConversion Then
            tmpPrice = mudtProps.ForeignPrice * DEL.Supplier.ConversionToLocalFactor
        Else
            tmpPrice = mudtProps.ForeignPrice / DEL.CaptureCurrency.Factor
        End If
        If tmpPrice > 0 Then
            lngRoundTo = oPC.Configuration.RoundingRules.GetRoundTo(tmpPrice)
        Else
            lngRoundTo = 0
        End If
        If lngRoundTo > 0 Then
            mudtProps.Price = (RoundUp(tmpPrice, lngRoundTo))
        Else
            mudtProps.Price = 0
        End If
        
        If bOK Then   'only validate if conversion is OK and the new value is other than the original
        End If
        If bOK And (lngTemp <> mudtProps.ForeignPrice) Then 'only validate if conversion is OK and the new value is other than the original
            bOK = ValidateObject("PRICE")
        End If
        bPriceChanged = True
        SetSPFromRecPrice

    Else
        lngTemp = mudtProps.Price
        bOK = SetField_strAsCurrencyToLong(mudtProps.Price, val, mcolStack.Count, "PRICE", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
        mudtProps.ForeignPrice = mudtProps.Price
        If bOK And (lngTemp <> mudtProps.Price) Then 'only validate if conversion is OK and the new value is other than the original
            bPriceChanged = True
            bOK = ValidateObject("PRICE")
            SetSPFromRecPrice
        End If
    End If
    RaiseEvent ValueChanges
  '  CalcCost
'    If bOK Then
'        SetPrice = bOK
'            SetSPFromRecPrice
'            RecalculateLine
'            ValidateObject "SP"
'    End If
    SetPrice = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.SetPrice(val)", val
End Function
Private Sub SetSPFromRecPrice()
    On Error GoTo errHandler
        
        If mudtProps.PriceSell = 0 Or bPriceChanged Then
            mudtProps.PriceSell = mudtProps.Price
        End If
        ApplyPTAdjustment
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.SetSPFromRecPrice"
End Sub
Private Sub ApplyPTAdjustment()
Dim Res As Long

        If mudtProps.PriceSell > 0 Then
        With Product
            If .B1Max <> 0 Then
                Res = PTAdjustment(mudtProps.Price, .Discount, .Rounded, .B1Max, .B2Max, .B3Max, .B1MU, .B2MU, .B3MU, oPC.RoundPriceTo)
                If .B1MU > 0 Or .B2MU > 0 Or .B3MU > 0 Or mudtProps.Price <> Res Then
                    mudtProps.PriceSell = Res
                End If
            End If
        End With
        End If

End Sub
Public Function SetDiscount(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim dblTEMP As Double
    If val = "" Then val = "0"
    dblTEMP = mudtProps.Discount
    bOK = SetField_DOUBLE(mudtProps.Discount, val, "DISCOUNT", mcolStack.Count)
    If bOK And (dblTEMP <> mudtProps.Discount) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("DISCOUNT")
    End If
    SetDiscount = bOK
 '   CalcCost
    RecalculateLine
    mDELParent.ValidateObject ""
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.SetDiscount(val)", val
End Function
Private Sub CalcCost()
    'THis is really only tempoary for calculating margin during capture - it is recalculated adding the extra charges when the Delivery is issued.
    mudtProps.Cost = mudtProps.Price * ((CDbl(100) - mudtProps.Discount) / 100)
End Sub
Public Function SetMargin(val As String) As Boolean
    If IsNumeric(val) Then
        dblMargin = val
        CalculateSPfromMargin
        SetMargin = True
    Else
        SetMargin = False
        Exit Function
    End If
End Function
Public Property Get MarginF() As String
    MarginF = Format(dblMargin, "0.00")
End Property
Public Property Get Margin_string() As String
    Margin_string = CStr(dblMargin)
End Property
Private Sub CalculateSPfromMargin()

    If mDELParent.ISForeignCurrency Then
        mudtProps.PriceSell = mudtProps.Cost / ((100 - dblMargin) / 100)
    Else
        mudtProps.PriceSell = mudtProps.Cost / ((100 - dblMargin) / 100)
    End If
    RaiseEvent ValueChanges
End Sub
Private Sub CalculateMarginFromSP()
    If mDELParent.ISForeignCurrency Then
        If mudtProps.PriceSell > 0 Then
            dblMargin = ((mudtProps.PriceSell - mudtProps.Cost) / mudtProps.PriceSell) * 100
        Else
            dblMargin = 0
        End If
    Else
        If mudtProps.PriceSell > 0 Then
            dblMargin = ((mudtProps.PriceSell - mudtProps.Cost) / mudtProps.PriceSell) * 100
        Else
            dblMargin = 0
        End If
    End If
    RaiseEvent ValueChanges
End Sub
Public Function DiscountF() As String
    On Error GoTo errHandler
    If mudtProps.Discount = 0 Then
        DiscountF = ""
    Else
        DiscountF = PBKSPercentF(Round(mudtProps.Discount, 2))
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.DiscountF"
End Function

Friend Sub SetParentDEL(pParentDEL As a_Delivery)
    On Error GoTo errHandler
    Set mDELParent = pParentDEL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.SetParentDEL(pParentDEL)", pParentDEL
End Sub
''---------------------------------
'Public Property Get ExtensionwithDiscount() As Long
'    ExtensionwithDiscount = CLng(DiscountedPrice * (mudtProps.QtyFirm + mudtProps.QtySS))
'End Property
'Public Property Get ExtensionwithDiscountF() As String
'    ExtensionwithDiscountF = Format(ExtensionwithDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get Extension() As Long
'    Extension = (CLng(CLng(mudtProps.Price * (mudtProps.QtySS + mudtProps.QtyFirm)) * (100 - mudtProps.Discount) / 100))
'End Property
'Public Property Get ExtensionF() As String
'    ExtensionF = Format(Extension / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get ExtensionSimple() As Long
'    ExtensionSimple = CLng(CLng(mudtProps.Price * (mudtProps.QtyFirm + mudtProps.QtySS)))
'End Property
'Public Property Get ExtensionSimpleF() As String
'    ExtensionSimpleF = Format(ExtensionSimple / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get DiscountedPrice() As Currency
'    DiscountedPrice = mudtProps.Price * ((100 - mudtProps.Discount) / 100)
'End Property
'Public Property Get DiscountedPriceF() As String
'    DiscountedPriceF = Format(DiscountedPrice, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
''---------------------------------
'Public Property Get ForeignExtensionwithDiscount() As Long
'    ForeignExtensionwithDiscount = CLng(ForeignDiscountedPrice * (mudtProps.QtyFirm + mudtProps.QtySS))
'End Property
'Public Property Get ForeignExtensionwithDiscountF() As String
'    ForeignExtensionwithDiscountF = Format(ForeignExtensionwithDiscount / mDELParent.CaptureCurrency.Divisor, mDELParent.CaptureCurrency.FormatString)
'End Property
'Public Property Get ForeignExtension() As Long
'    ForeignExtension = (CLng(CLng(mudtProps.ForeignPrice * (mudtProps.QtySS + mudtProps.QtyFirm)) * (100 - mudtProps.Discount) / 100))
'End Property
'Public Property Get ForeignExtensionF() As String
'    ForeignExtensionF = Format(ForeignExtension / mDELParent.CaptureCurrency.Divisor, mDELParent.CaptureCurrency.FormatString)
'End Property
'Public Property Get ForeignExtensionSimple() As Long
'    ForeignExtensionSimple = CLng(CLng(mudtProps.ForeignPrice * (mudtProps.QtyFirm + mudtProps.QtySS)))
'End Property
'Public Property Get ForeignExtensionSimpleF() As String
'    ForeignExtensionSimpleF = Format(ForeignExtensionSimple / mDELParent.CaptureCurrency.Divisor, mDELParent.CaptureCurrency.FormatString)
'End Property
'Public Property Get ForeignDiscountedPrice() As Currency
'    ForeignDiscountedPrice = mudtProps.ForeignPrice * ((100 - mudtProps.Discount) / 100)
'End Property
'Public Property Get ForeignDiscountedPriceF() As String
'    ForeignDiscountedPriceF = Format(ForeignDiscountedPrice, mDELParent.CaptureCurrency.FormatString)
'End Property






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
    ErrorIn "a_DeliveryLine.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Invalid price", "PRICE"
    colClassors.Add "Invalid quantity", "QTY"
    colClassors.Add "Invalid discount", "DISCOUNT"
    colClassors.Add "A numeric value is too great for calculation", "OVERFLOW"
    colClassors.Add "Margin too low", "PRICESELL"
    colClassors.Add "Line not matched to purchase order or declared unmatched", "POLID"
    colClassors.Add "Zero S.P.", "SP"
    colClassors.Add "Title is missing, edit product before going on", "TITLE"
    colClassors.Add "Categories or product types not assigned", "CAT"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.LoadClassorsCollection"
End Sub
Private Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean

    bValid = True
    If mudtProps.POLID > -1 Then
        mobjValid.RuleBroken "POLID", False
    Else
        If pFld = "POLID" Then bValid = False
        mobjValid.RuleBroken "POLID", True
    End If
    If mudtProps.Price >= 0 Then
        mobjValid.RuleBroken "PRICE", False
    Else
        If pFld = "PRICE" Then bValid = False
        mobjValid.RuleBroken "PRICE", True
    End If
    If FNS(mudtProps.Title) > "" Then
        mobjValid.RuleBroken "TITLE", False
    Else
        If pFld = "TITLE" Then bValid = False
        mobjValid.RuleBroken "TITLE", True
    End If
    If Markup(mudtProps.PriceSell, mudtProps.Cost) >= oPC.Configuration.MinMU Then
        mobjValid.RuleBroken "PRICESELL", False
    Else
        If pFld = "PRICESELL" Then bValid = False
        mobjValid.RuleBroken "PRICESELL", True
    End If
    If mudtProps.PriceSell > 0 Then
        mobjValid.RuleBroken "SP", False
    Else
        If pFld = "SP" Then bValid = False
        mobjValid.RuleBroken "SP", True
    End If

    If (mudtProps.QtySS > 0 Or mudtProps.QtyFirm > 0 Or mudtProps.QtyShort > 0) And (Not mudtProps.QtySS < 0) And (Not mudtProps.QtyFirm < 0) And (Not mudtProps.QtyShort < 0) Then
        mobjValid.RuleBroken "QTY", False
    Else
        If pFld = "QTY" Or pFld = "QTYFIRM" Then bValid = False
        mobjValid.RuleBroken "QTY", True
    End If
    If (mudtProps.Discount < 0 Or mudtProps.Discount > 100) Then
        If pFld = "DISCOUNT" Then bValid = False
        mobjValid.RuleBroken "DISCOUNT", True
    Else
        mobjValid.RuleBroken "DISCOUNT", False
    End If
    If oPC.Configuration.EnforceSections = True Then
        If Me.Product Is Nothing Then
            mobjValid.RuleBroken "CAT", True
        Else
            If Me.Product.ProductSections.Count = 0 Or Me.ProductTypeID = oPC.Configuration.DefaultPT Then
                mobjValid.RuleBroken "CAT", True
            Else
                mobjValid.RuleBroken "CAT", False
            End If
        End If
    Else
        mobjValid.RuleBroken "CAT", False
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
    ErrorIn "a_DeliveryLine.ValidateObject(pFld)", pFld
End Function
Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.SetDirty(pVal)", pVal
End Sub
Public Property Get Product() As a_Product
    Set Product = moProd
End Property
Public Sub GetStatus()
    On Error GoTo errHandler
    mobjValid.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.GetSTatus"
End Sub
Public Property Get DEL() As a_Delivery
    Set DEL = mDELParent
End Property
Public Function SetLineProduct(Optional PID As String, Optional pCode As String, _
                                Optional IsEditingRow As Boolean = False) As Boolean
10        On Error GoTo errHandler
      Dim tmp As Double
20        Set moProd = New a_Product
30        With moProd
40            If .Load(PID, 0, FNS(pCode)) <> 99 Then   'product found
50                Me.Title = .TitleAuthorPublisher
60                Me.PID = .PID
70                If Not Me.IsEditing Then ClearValues
80                If (Not IsEditingRow) Or (mDELParent.DeliveryLines.FindLineByEAN(Me.EAN) Is Nothing) Then
90                    If mudtProps.Price = 0 Then
100                       mudtProps.Price = moProd.RRP
110                   End If
120                   If mudtProps.PriceSell = 0 Then
130                       mudtProps.PriceSell = moProd.SP
140                   End If
150               End If
160               Me.SetSection .Section
170               Me.ProductTypeID = .ProductTypeID
180               Me.code = .code
190               Me.CodeF = .CodeF
200               Me.EAN = .EAN
210               If .PID = "" Then
220                   LogSaveToFile "PID is nothing: in a_Deliveryline:SetLineProduct" & vbCrLf & "pCode = " & pCode
230                   MsgBox "LEAVE THIS MESSAGE ON THE SCREEN AND CALL GERARD OR DAVID at Papyrus Support" & vbCrLf & "PID is nothing: in a_Deliveryline:SetLineProduct" & vbCrLf & "pCode = " & pCode
240               End If
250               SetLineProduct = True
260           Else
270               SetLineProduct = False   'Not found
280           End If
290       End With
300       mobjValid.GetStatus
310       Exit Function
errHandler:
320       If ErrMustStop Then Debug.Assert False: Resume
330       ErrorIn "a_DeliveryLine.SetLineProduct(PID,pCODE)", Array(PID, pCode), , , "line number", Array(Erl())
End Function

Public Property Get Section() As String
    Section = FNS(mudtProps.Section)
End Property
Public Function SetSection(pVal As String) As Boolean
    mudtProps.Section = pVal
End Function

Public Property Get ProductTypeID() As Long
    On Error GoTo errHandler
    ProductTypeID = FNN(mudtProps.ProductTypeID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.ProductTypeID"
End Property
Public Property Let ProductTypeID(Value As Long)
    On Error GoTo errHandler
    mudtProps.ProductTypeID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine.ProductTypeID(Value)", Value
End Property

Public Property Get MBCode() As String
    MBCode = FNS(mudtProps.MBCode)
End Property
Public Property Let MBCode(val As String)
    mudtProps.MBCode = FNS(val)
End Property

