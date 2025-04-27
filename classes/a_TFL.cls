VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_TFL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event Valid(msg As String)
Event ValueChanges()
Private moProd As a_Product
Private mudtProps As TFLProps
Private mcolStack As Collection
Private mContainingCollection As ch_TFL
Private colClassors As Collection
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private mKey As String
Private mTFParent As a_TF

Private mExt As Long
Private mExtExVAT As Long
Private mDiscExVAT As Long
Private mExtLessDisc As Long
Private mExtLessDiscExVAT As Long
Private mCostExt As Long

Public Property Get Product() As a_Product
    Set Product = moProd
End Property
Public Property Get testParent() As Boolean
    testParent = (mContainingCollection Is Nothing)
End Property
Public Function RecalculateLine()
    On Error GoTo errHandler
    mExt = mudtProps.Price * (mudtProps.Qty)
    mExtLessDisc = Round(mExt * ((100 - mudtProps.Discount) / 100), 0)
    mExtLessDiscExVAT = Round(mExtLessDisc * (100 / (100 + mudtProps.VATRate)), 0)
    mExtExVAT = Round(mExt * (100 / (100 + mudtProps.VATRate)), 0)
    mDiscExVAT = mExtExVAT - mExtLessDiscExVAT
    mCostExt = mudtProps.Cost * (mudtProps.Qty)
    
    RaiseEvent ValueChanges
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'overflow
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.RecalculateLine"
End Function

Public Property Get LineCost() As Long
    LineCost = mCostExt
End Property
Public Property Get LineCostF() As String
    LineCostF = Format(CDbl(LineCost / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get DiscExVAT() As Long
    DiscExVAT = mDiscExVAT
End Property
Public Property Get DiscExVATF() As String
    DiscExVATF = Format(CDbl(DiscExVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get ExtExVAT() As Long
    ExtExVAT = mExtExVAT
End Property
Public Property Get ExtExVATF() As String
    ExtExVATF = Format(CDbl(ExtExVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get ExtLessDiscExVAT() As Long
        ExtLessDiscExVAT = mExtLessDiscExVAT
End Property
Public Property Get ExtLessDiscExVATF() As String
    ExtLessDiscExVATF = Format(CDbl(ExtLessDiscExVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get ExtLessDisc() As Long
        ExtLessDisc = mExtLessDisc
End Property
Public Property Get ExtLessDiscF() As String
    ExtLessDiscF = Format(CDbl(ExtLessDisc / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
End Property
'------------------------------


Public Property Get PriceF() As String
    On Error GoTo errHandler
    PriceF = Format(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.PriceF"
End Property
Public Property Get CostF() As String
    On Error GoTo errHandler
    CostF = Format(mudtProps.Cost / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.CostF"
End Property



Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property
Friend Sub SetParentTF(pParentTF As a_TF)
    On Error GoTo errHandler
    Set mTFParent = pParentTF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.SetParentTF(pParentTF)", pParentTF
End Sub

Private Function GetState() As String
    On Error GoTo errHandler
  Dim mudtData As TFLData
  LSet mudtData = mudtProps
  GetState = mudtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim mudtData As TFLData
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.SetSuperState(buffer)", buffer
End Sub
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
    ErrorIn "a_TFL.GetSTatus"
End Function


Public Sub ApplyEdit()
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 445
   
    If Not mContainingCollection Is Nothing Then
          mContainingCollection.AddTFLine Me
          Set mContainingCollection = Nothing
    End If
    mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.ApplyEdit"
End Sub
Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.BeginEdit"
End Sub
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.IsEditing"
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
    ErrorIn "a_TFL.CancelEdit"
End Sub
Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    LoadClassorsCollection
    mudtProps.Qty = 1
    mudtProps.IsNew = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Class_Initialize"
End Sub
Friend Sub Initialize(Parent As ch_TFL)
    On Error GoTo errHandler
  If mudtProps.IsNew Then Set mContainingCollection = Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Initialize(Parent)", Parent
End Sub

Friend Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.IsValid"
End Property

Private Sub Class_Terminate()
    On Error GoTo errHandler
  If mContainingCollection Is Nothing Then
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
  End If
  Set mcolStack = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(oRS)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.mobjValid_BrokenRule(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.mobjValid_NoBrokenRules"
End Sub

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Delete"
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.IsDeleted"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.IsDirty"
End Property







Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = mudtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.ID"
End Property
Public Property Get Qty() As Long
    On Error GoTo errHandler
    Qty = mudtProps.Qty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Qty"
End Property
Public Property Let Qty(val As Long)
    On Error GoTo errHandler
    mudtProps.Qty = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Qty(val)", val
End Property
Public Function SetQty(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Qty
    bOK = SetField_LONG(mudtProps.Qty, val, "QTY", mcolStack.Count)
    If bOK And mudtProps.Qty <> lngTemp Then   'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QTY")
    End If
    SetQty = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.SetQty(val)", val
End Function
Public Function SetPrice(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Price
    bOK = SetField_LONG(mudtProps.Price, val, "PRICE", mcolStack.Count)
    If bOK And mudtProps.Price <> lngTemp Then   'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("PRICE")
    End If
    SetPrice = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.SetPrice(val)", val
End Function
Public Function SetDiscount(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Double
    lngTemp = mudtProps.Discount
    bOK = SetField_DOUBLE(mudtProps.Discount, val, "DISCOUNT", mcolStack.Count)
    If bOK And mudtProps.Discount <> lngTemp Then   'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("DISCOUNT")
    End If
    SetDiscount = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.SetDiscount(val)", val
End Function

Public Property Get Note() As String
    On Error GoTo errHandler
    Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Note"
End Property
Public Property Let Note(val As String)
    On Error GoTo errHandler
    mudtProps.Note = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Note(val)", val
End Property
Public Property Get Author() As String
    On Error GoTo errHandler
    Author = FNS(mudtProps.Author)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Author"
End Property
Public Property Let Author(val As String)
    On Error GoTo errHandler
    mudtProps.Author = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Author(val)", val
End Property

Public Property Get BindingCode() As String
    On Error GoTo errHandler
    BindingCode = FNS(mudtProps.Binding)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.BindingCode"
End Property

Public Property Get Title() As String
    On Error GoTo errHandler
    Title = FNS(mudtProps.Title)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Title"
End Property
Public Property Let Title(val As String)
    On Error GoTo errHandler
    mudtProps.Title = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Title(val)", val
End Property
Public Property Let CodeF(val As String)
    On Error GoTo errHandler
    mudtProps.CodeF = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.CodeF(val)", val
End Property
Public Property Let curPrice(val As Long)
    On Error GoTo errHandler
    mudtProps.Price = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.curPrice(val)", val
End Property
Public Property Get Cost() As Long
    On Error GoTo errHandler
    Cost = mudtProps.Cost
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Cost"
End Property
Public Property Let Cost(val As Long)
    On Error GoTo errHandler
    mudtProps.Cost = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Cost(val)", val
End Property

Public Property Get PID() As String
    On Error GoTo errHandler
    PID = FNS(mudtProps.PID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.PID"
End Property
Public Property Let PID(val As String)
    On Error GoTo errHandler
    mudtProps.PID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.PID(val)", val
End Property
Public Property Get EAN() As String
    On Error GoTo errHandler
    EAN = FNS(mudtProps.EAN)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.EAN"
End Property
Public Property Let EAN(val As String)
    On Error GoTo errHandler
    If Len(val) > Len(mudtProps.EAN) Then Err.Raise vbObjectError + 1001, "String value too long"
    mudtProps.EAN = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.EAN(val)", val
End Property
Public Property Get CodeF() As String
    On Error GoTo errHandler
    CodeF = FNS(mudtProps.CodeF)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.CodeF"
End Property
Public Property Get code() As String
    On Error GoTo errHandler
    code = FNS(mudtProps.code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Code"
End Property
Public Property Let code(val As String)
    On Error GoTo errHandler
    mudtProps.code = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Code(val)", val
End Property
Public Property Let Discount(val As Double)
    On Error GoTo errHandler
    mudtProps.Discount = val 'IIf(val = 0, 38, val)  'we use 38% as nominal value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Discount(val)", val
End Property
Public Property Get Discount() As Double
    On Error GoTo errHandler
    Discount = Round(FNDBL(mudtProps.Discount), 2)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Discount"
End Property
Public Property Get DiscountF() As String
    On Error GoTo errHandler
    DiscountF = PBKSPercentF(Round(FNDBL(mudtProps.Discount), 2))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.DiscountF"
End Property

Public Property Let VATRate(val As Double)
    On Error GoTo errHandler
    mudtProps.VATRate = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.VATRate(val)", val
End Property
Public Property Get VATRate() As Double
    On Error GoTo errHandler
    VATRate = FNDBL(mudtProps.VATRate)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.VATRate"
End Property


Public Property Let Section(val As String)
    On Error GoTo errHandler
    mudtProps.Section = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Section(val)", val
End Property
Public Property Get Price() As Long
    On Error GoTo errHandler
    Price = mudtProps.Price
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Price"
End Property
Public Property Let Price(val As Long)
    On Error GoTo errHandler
    mudtProps.Price = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.Price(val)", val
End Property

Public Property Get TRID() As Long
    On Error GoTo errHandler
    TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.TRID"
End Property
Public Property Let TRID(val As Long)
    On Error GoTo errHandler
    mudtProps.TRID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.TRID(val)", val
End Property
'Public Property Let DateTime(val As String)
'
'    mobjValid.RuleBroken "DAteTime", (Not IsDate(val))
'    If (IsDate(val)) Or (val = "0") Then mudtProps.DateTime = CVDate(val)
'    SetDirty True
'End Property
'Public Property Get DateTime() As String
'    If CLng(mudtProps.DateTime) = 0 Then
'        DateTime = ""
'    Else
'        DateTime = Format(mudtProps.DateTime, "dd-mm-yyyy hh:mm:ss")
'    End If
'End Property
'Public Property Get dteDateTime() As Date
'    If HasData(mudtProps.DateTime) Then
'        dteDateTime = mudtProps.DateTime
'    End If
'End Property
'Public Property Get CSLParent() As c_CSL
'    Set CSLParent = mCSLParent
'End Property
'Public Property Set CSLParent(ByVal NewParent As c_CSL)
'    If CSLParent Is Nothing Then
'        ' Assign the initial value.
'        Set mCSLParent = NewParent
'    Else
'        Err.Raise Number:=vbObjectError + 32144, Description:="Parent property is read-only"
'    End If
'End Property
Private Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler


Dim bOK As Boolean
    bOK = True
    If mudtProps.Price > 0 Then
        mobjValid.RuleBroken "PRICE", False
    Else
        mobjValid.RuleBroken "PRICE", True
    End If
    If mudtProps.Qty >= 0 Then
        mobjValid.RuleBroken "QTY", False
    Else
        mobjValid.RuleBroken "QTY", True
    End If
    mobjValid.GetStatus
    ValidateObject = bOK
    SetDirty bOK
EXIT_Handler:
'H:
'    MsgBox "CheckOK: " & or
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.ValidateObject(pFld)", pFld
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
    ErrorIn "a_TFL.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Invalid price", "PRICE"
    colClassors.Add "Invalid quantity", "QTY"
    colClassors.Add "Invalid discount", "DISCOUNT"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.LoadClassorsCollection"
End Sub


Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.SetDirty(pVal)", pVal
End Sub

Public Function SetLineProduct(Optional PID As String, Optional pCode As String) As Boolean
    On Error GoTo errHandler
    
    Set moProd = New a_Product
    With moProd
        If .Load(PID, 0, FNS(pCode)) <> 99 Then   'product found
            Me.Title = .TitleAuthorPublisher
            Me.PID = .PID
            mudtProps.Binding = .BindingCode
            
            If .SP > 0 Then
                SetPrice .SP
            Else
                SetPrice .RRP
            End If
            Me.Cost = .Cost
            Me.SetSection .Section
            If oPC.Configuration.TFRDiscount > 0 Then
                Me.Discount = oPC.Configuration.TFRDiscount
'            Else
'                Me.Discount = oPC.Configuration.TFRDiscountAdj + .DealDiscount
            End If
            Me.VATRate = .VATRateToUse
            Me.Qty = 1
            Me.code = .code
            Me.CodeF = .CodeF
            SetLineProduct = True
        Else
            SetLineProduct = False   'Not found
        End If
    End With
    mobjValid.GetStatus
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.SetLineProduct(PID,pCODE)", Array(PID, pCode)
End Function
'Public Function SetSectionALL(pVal As String) As Boolean
'    On Error GoTo ErrHandler
'Dim tmp As String
'Dim bOK As Boolean
'
'    tmp = pVal
'    bOK = oPC.Configuration.Sections.ValidateEntireSet(tmp)
'    If bOK Then
'        mudtProps.Section = tmp
'    End If
'    SetSectionALL = bOK
'    ValidateObject ("SECTION")
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_TFL.SetSectionALL(pVal)", pVal
'End Function
'Public Function SetSection(pVal As String) As Boolean
'    On Error GoTo ErrHandler
'    SetSection = oPC.Configuration.Sections.ValidateNewSetMember(mudtProps.Section, pVal)
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_TFL.SetSection(pVal)", pVal
'End Function

Public Property Get Section() As String
    Section = FNS(mudtProps.Section)
End Property
Public Function SetSection(pVal As String) As Boolean
    mudtProps.Section = pVal
End Function
Public Property Get VATRateF() As String
    On Error GoTo errHandler
    VATRateF = Format(mudtProps.VATRate, gPercentFormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.VATRateF"
End Property
Public Property Get SpecialVATMsg() As String
    On Error GoTo errHandler
    If Me.VATRate <> oPC.Configuration.VATRate And oPC.Configuration.IsVATRegion Then
         SpecialVATMsg = "(" & VATRateF & " VAT) "
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL.SpecialVATMsg"
End Property
''
Public Property Get Properties(pSrt As enSortField) As Variant
    On Error GoTo errHandler
    Select Case pSrt
    Case enSequence
        Properties = Me.ID
    Case enTitle
        Properties = Me.Title
    Case enAuthor
        Properties = Me.Author
    Case enCode
        Properties = Me.code
    End Select
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DEL.Properties(pSrt)", pSrt
End Property

