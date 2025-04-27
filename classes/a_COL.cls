VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_COL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public oProd As a_Product

Event Valid(msg As String)
Private mudtProps As COLProps
Private mcolStack As Collection
Private mobjParent As ch_COL
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassors As Collection
Private mCOParent As a_CO
'XX Private mchCOFFs As ch_COFF

Private mKey As String

Private mExt As Long
Private mExtS As Long
Private mDisc As Long
Private mFVAT As Long
Private mVAT As Long
Private mExt_Ex As Long


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
    mExt = CLng(CDbl(mudtProps.Price * mudtProps.Qty) * CDbl((100 - (mudtProps.Discount))) / 100)
    mDisc = (mudtProps.Price * mudtProps.Qty) - mExt
    mVAT = (mExt - (CDbl(mExt) * 100 / (100 + mudtProps.VATRate)))
    mExt_Ex = mExt - mVAT

    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'ov erflow
        mobjValid.RuleBroken "OVERFLOW", True
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.CalculateLine"
End Function
Public Property Get Ext() As Long
    On Error GoTo errHandler

        Ext = mExt

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Ext"
End Property

Public Property Get ExtF() As String
    On Error GoTo errHandler

        ExtF = Format(CDbl(mExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtF"
End Property

Public Property Get Disc() As Long
    On Error GoTo errHandler

        Disc = mDisc

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Disc"
End Property
Public Property Get DiscF() As String
    On Error GoTo errHandler

        DiscF = Format(CDbl(mDisc / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.DiscF"
End Property
Public Property Get ExtSF() As String
    On Error GoTo errHandler

        ExtSF = Format(CDbl(mExtS / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtSF"
End Property
Public Property Get ExtS() As Long
    On Error GoTo errHandler

        ExtS = mExtS

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtS"
End Property

Public Property Get Ext_Ex() As Long
    On Error GoTo errHandler

        Ext_Ex = mExt_Ex

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Ext_Ex"
End Property
Public Property Get VAT() As Long
    On Error GoTo errHandler
        VAT = mVAT
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.VAT"
End Property
Public Property Get VATF() As String
    On Error GoTo errHandler
        VATF = Format(CDbl(mVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.VATF"
End Property

Public Property Get CO() As a_CO
    Set CO = mCOParent
End Property
Private Function GetState() As String
    On Error GoTo errHandler
  Dim mudtData As COLData
  
  LSet mudtData = mudtProps
  GetState = mudtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim mudtData As COLData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
    GetSuperState = GetState
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler

    SetState buffer
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetSuperState(buffer)", buffer
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddCOLine Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ApplyEdit"
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
    ErrorIn "a_COL.BeginEdit"
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
    ErrorIn "a_COL.CancelEdit"
End Sub
Public Property Get ServiceItem() As Boolean
    On Error GoTo errHandler
    ServiceItem = mudtProps.ServiceItem
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ServiceItem"
End Property
Public Property Let ServiceItem(val As Boolean)
    mudtProps.ServiceItem = val
End Property
Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    LoadClassorsCollection
    mobjValid.RuleBroken "Price", True
    mobjValid.RuleBroken "ETA", False
    mobjValid.RuleBroken "REF", True
    mudtProps.ETA = DateAdd("m", 1, Date)
   'XX  Set mchCOFFs = New ch_COFF
    mudtProps.Discount = 0
    mudtProps.Deposit = 0
    mudtProps.DepositStatus = "P"
    If oPC.AllowsSSInvoicing Then
        mudtProps.Qty = 0
        mudtProps.QtyFirm = 0
        mudtProps.QtySS = 0
    Else
        mudtProps.Qty = 1
        mudtProps.QtyFirm = 1
        mudtProps.QtySS = 0
    End If
    mudtProps.IsNew = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Class_Initialize"
End Sub
Friend Sub Initialize(Parent As ch_COL)
    On Error GoTo errHandler
    If mudtProps.IsNew Then
        Set mobjParent = Parent
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Initialize(Parent)", Parent
End Sub
Friend Sub SetParentCO(pParentCO As a_CO)
    On Error GoTo errHandler
    Set mCOParent = pParentCO
    If mCOParent.OrderType = enWant Then
        mobjValid.RuleBroken "ETA", False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetParentCO(pParentCO)", pParentCO
End Sub
Friend Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.IsValid"
End Property

Private Sub Class_Terminate()
    On Error GoTo errHandler
  'XX   Set mchCOFFs = Nothing
    Set oProd = Nothing
  If mobjParent Is Nothing Then
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
  End If
  Set mcolStack = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.mobjValid_Status(pMsg)", pMsg
End Sub
Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Delete"
End Sub
Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty 'XX Or mchCOFFs.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.IsDirty"
End Property

Public Property Get BottomOfDocument() As Boolean
    On Error GoTo errHandler
    BottomOfDocument = (Right(code, 2) = "_B")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.BottomOfDocument"
End Property

Public Property Get COLineID() As Long
    On Error GoTo errHandler
    COLineID = mudtProps.COLineID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.COLineID"
End Property

Public Property Let COLineID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COLineID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.COLineID(val)", val
End Property
Public Property Get TRID() As Long
    On Error GoTo errHandler
  TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.TRID"
End Property
Public Property Let TRID(val As Long)
    On Error GoTo errHandler
    mudtProps.TRID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.TRID(val)", val
End Property
Public Property Get SupplierID() As Long
    On Error GoTo errHandler
  SupplierID = mudtProps.SupplierID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SupplierID"
End Property
Public Property Get HasSupplierID() As String
    On Error GoTo errHandler
    If oPC.IncludeSupplierFeatures Then
        HasSupplierID = IIf(mudtProps.SupplierID > 0, "@", "")
    Else
        HasSupplierID = ""
    End If
    
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.HasSupplierID"
End Property

Public Property Get PID() As String
    On Error GoTo errHandler
    PID = mudtProps.ProdID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.PID"
End Property
Public Property Let PID(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    SetDirty (mudtProps.ProdID <> val)
    mudtProps.ProdID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.PID(val)", val
End Property
Public Function LoadProduct(PID As String)
    On Error GoTo errHandler
    Set oProd = Nothing
    Set oProd = New a_Product
    
    If oProd.Load(PID, 0) <> 99 Then   'product found
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.LoadProduct(PID)", PID
End Function
Public Function SetLineProduct(Optional PID As String, Optional pCode As String) As Boolean
    On Error GoTo errHandler
    Set oProd = Nothing
    Set oProd = New a_Product
    With oProd
        If .Load(PID, 0, Trim$(pCode)) <> 99 Then   'product found
            Me.Title = .TitleAuthorPublisher
            Me.PID = .PID
            Me.Price = .SP
            Me.VATRate = .VATRate
            Me.ServiceItem = .IsServiceItem
            Me.code = .code
            Me.CodeF = .EAN
            Me.EAN = .EAN
            Me.SetDiscount CO.Customer.DefaultDiscount
            Me.SetETAFromDeliveryDays .DefaultDeliveryDays
            
            SetLineProduct = True
        Else
            SetLineProduct = False   'Not found
        End If
    End With
    mobjValid.GetStatus
 '   Set oProd = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetLineProduct(PID,pCODE)", Array(PID, pCode)
End Function
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
    ErrorIn "a_COL.SetLineExtraProduct(pCODE)", pCode
End Function

Public Function SetETA(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim dteTemp As Date
    If mCOParent.OrderType = enWant Then
        SetETA = True
        Exit Function
    End If
    dteTemp = mudtProps.ETA
    If IsDate(val) Then
        bOK = SetField_DATE(mudtProps.ETA, val, "ETA", mcolStack.Count)
    Else
        bOK = SetField_DIARYPERIODS(mudtProps.ETA, val, "ETA", mcolStack.Count)
    End If
    If bOK And (dteTemp <> mudtProps.ETA) Then 'only validate if conversion is OK and the new value is other than the original
        mudtProps.ETACode = val
        bOK = ValidateObject("ETA")
    End If
    SetETA = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetETA(val)", val
End Function
Public Function SetETAFromDeliveryDays(pDefaultDeliveryDays As Integer)
    On Error GoTo errHandler
    mudtProps.ETA = DateAdd("d", pDefaultDeliveryDays, Date)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetETAFromDeliveryDays(pDefaultDeliveryDays)", pDefaultDeliveryDays
End Function
Public Property Get ETACode() As String
    On Error GoTo errHandler
    ETACode = mudtProps.ETACode
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ETACode"
End Property
Public Property Get ETA() As Date
    On Error GoTo errHandler
    ETA = mudtProps.ETA
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ETA"
End Property
Public Property Get ETAF() As String
    On Error GoTo errHandler
    ETAF = Format(mudtProps.ETA, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ETAF"
End Property
Public Property Get EAN() As String
    On Error GoTo errHandler
    EAN = mudtProps.EAN
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.EAN"
End Property
Public Property Let EAN(val As String)
    On Error GoTo errHandler
    mudtProps.EAN = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.EAN(val)", val
End Property
Public Property Get WantDate() As Date
    On Error GoTo errHandler
    WantDate = mudtProps.WantDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.WantDate"
End Property
Public Property Get WantDateF() As String
    On Error GoTo errHandler
    WantDateF = Format(mudtProps.WantDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.WantDateF"
End Property
Public Function SetWantDate(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim dteTemp As Date
    dteTemp = mudtProps.WantDate
    bOK = ConvertToDate(val, mudtProps.WantDate) ', val, "WANTDATE", mcolStack.Count)
    If bOK And (dteTemp <> mudtProps.WantDate) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("WANTDATE")
    End If
    SetWantDate = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetWantDate(val)", val
End Function
Public Property Get Note() As String
    On Error GoTo errHandler
    Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Note"
End Property
Public Property Get LastAction() As String
    On Error GoTo errHandler
    LastAction = FNS(mudtProps.LastAction)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.LastAction"
End Property
Public Property Get LastActionAndDate() As String
    On Error GoTo errHandler
Dim bDummy As Boolean
Dim strReport As String
Dim strDiarizePeriod As String
Dim strDia As String
    strDiarizePeriod = ConvertCOLActionCodes(FNS(mudtProps.LastAction), bDummy, strDia, strReport)
    LastActionAndDate = strReport & " (" & Format(FNDF(mudtProps.LastActionDate), "dd/mm/yyyy") & ")"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.LastActionAndDate"
End Property
Public Property Let Note(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If mudtProps.Note <> val Then
        mudtProps.Note = stripCRLF(val)
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Note(val)", val
End Property

'Public Property Get COLineCode() As String
'    COLineCode = FNS(mudtProps.COLineCode)
'End Property
'Public Property Let COLineCode(val As String)
'  If mcolStack.Count = 0 Then Err.Raise 383
'    If mudtProps.COLineCode <> val Then
'        mudtProps.COLineCode = val
'        SetDirty True
'    End If
'End Property
Public Property Get POLID() As Long
    On Error GoTo errHandler
    POLID = FNN(mudtProps.POLID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.POLID"
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
Public Property Get code() As String
    On Error GoTo errHandler
    code = FNS(mudtProps.code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Code"
End Property
Public Property Let code(val As String)
    On Error GoTo errHandler
    mudtProps.code = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Code(val)", val
End Property
Public Property Get DeliveryDocument() As String
    On Error GoTo errHandler
    DeliveryDocument = FNS(mudtProps.DeliveryDocument)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.DeliveryDocument"
End Property
Public Property Let DeliveryDocument(val As String)
    On Error GoTo errHandler
    mudtProps.DeliveryDocument = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.DeliveryDocument(val)", val
End Property

Public Property Get ExtraCode() As String
    On Error GoTo errHandler
    ExtraCode = FNS(mudtProps.ExtraCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtraCode"
End Property
Public Property Let ExtraCode(val As String)
    On Error GoTo errHandler
    mudtProps.ExtraCode = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtraCode(val)", val
End Property

Public Property Get ExtraPID() As String
    On Error GoTo errHandler
    ExtraPID = FNS(mudtProps.ExtraPID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtraPID"
End Property
Public Property Let ExtraPID(val As String)
    On Error GoTo errHandler
    mudtProps.ExtraPID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtraPID(val)", val
End Property


Sub TestEvent()
    On Error GoTo errHandler
RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.TestEvent"
End Sub
'Public Property Let Code(val As String)
'Dim oCode As z_ProdCode
'  If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.Code = val
'    Set oCode = New z_ProdCode
'    oCode.Load val
'    mudtProps.CodeF = oCode.FormatCode
'    mudtProps.IsDirty = True
'    Set oCode = Nothing
'End Property

Public Property Get CodeF() As String
    On Error GoTo errHandler
    CodeF = FNS(mudtProps.CodeF)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.CodeF"
End Property
Public Property Let CodeF(val As String)
    On Error GoTo errHandler
    mudtProps.CodeF = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.CodeF(val)", val
End Property

'Public Property Get CodeForExport() As String
'  CodeForExport = FNS(mudtProps.CodeForExport)
'End Property
'Public Property Let CodeForExport(Val As String)
'  If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.CodeForExport = Val
'End Property
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
    ErrorIn "a_COL.TitleAuthorPublisher"
End Property
'Public Property Get COLType() As String
'    On Error GoTo ErrHandler
'    COLType = FNS(mudtProps.COLType)
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_COL.COLType"
'End Property
'Public Property Get COLParent() As Long
'    On Error GoTo ErrHandler
'    COLParent = FNS(mudtProps.COLParent)
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_COL.COLParent"
'End Property

Public Property Get Publisher() As String
    On Error GoTo errHandler
    Publisher = FNS(mudtProps.Publisher)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Publisher"
End Property
Public Property Get Title() As String
    On Error GoTo errHandler
    Title = stripCRLF(FNS(mudtProps.Title))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Title"
End Property
Public Property Let Title(val As String)
    On Error GoTo errHandler
    If val <> mudtProps.Title Then
        mudtProps.Title = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Title(val)", val
End Property
Public Property Get ExtraChargeDescription() As String
    On Error GoTo errHandler
    ExtraChargeDescription = stripCRLF(FNS(mudtProps.ExtraChargeDescription))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtraChargeDescription"
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
    ErrorIn "a_COL.ExtraChargeDescription(val)", val
End Property

Public Property Get Fulfilled() As String
    On Error GoTo errHandler
    Fulfilled = FNS(mudtProps.Fulfilled)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Fulfilled"
End Property
Public Sub SetFulfilled(val As String)
    On Error GoTo errHandler
    mudtProps.Fulfilled = val
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetFulfilled(val)", val
End Sub
Public Property Get Ref() As String
    On Error GoTo errHandler
    Ref = FNS(mudtProps.Ref)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Ref"
End Property
Public Function SetRef(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
    If val <> mudtProps.Ref Then
        mudtProps.Ref = Left(val, 50)
        bOK = ValidateObject("REF")
        SetRef = bOK
        SetDirty True
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetRef(val)", val
End Function

Public Property Get MainAuthor() As String
    On Error GoTo errHandler
    MainAuthor = stripCRLF(FNS(mudtProps.Author))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.MainAuthor"
End Property
Public Property Let MainAuthor(val As String)
    On Error GoTo errHandler
    If val <> mudtProps.Author Then
        mudtProps.Author = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.MainAuthor(val)", val
End Property
'---------------------------------
Public Property Get ExtensionInclDeposit() As Long
    On Error GoTo errHandler
    If oPC.AllowsSSInvoicing Then
        ExtensionInclDeposit = CLng(CDbl(mudtProps.Price * (mudtProps.QtyFirm + mudtProps.QtySS)) * CDbl((100 - (mudtProps.Discount))) / 100)
    Else
        ExtensionInclDeposit = CLng(CDbl(mudtProps.Price * mudtProps.Qty) * CDbl((100 - (mudtProps.Discount))) / 100)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtensionInclDeposit"
End Property
Public Property Get ExtensionInclDepositF() As String
    On Error GoTo errHandler
    ExtensionInclDepositF = Format(ExtensionInclDeposit / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtensionInclDepositF"
End Property
Public Property Get Extension() As Long
    On Error GoTo errHandler
    If oPC.AllowsSSInvoicing Then
        Extension = (CLng(CLng(mudtProps.Price * (mudtProps.QtyFirm + mudtProps.QtySS)) * (100 - mudtProps.Discount) / 100))
    Else
        Extension = (CLng(CLng(mudtProps.Price * mudtProps.Qty) * (100 - mudtProps.Discount) / 100))
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Extension"
End Property
Public Property Get ExtensionF() As String
    On Error GoTo errHandler
    ExtensionF = Format(Extension / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtensionF"
End Property
Public Property Get ExtensionSimple() As Long
    On Error GoTo errHandler
    If oPC.AllowsSSInvoicing Then
        ExtensionSimple = CLng(CLng(mudtProps.Price * (mudtProps.QtyFirm + mudtProps.QtySS)))
    Else
        ExtensionSimple = CLng(CLng(mudtProps.Price * mudtProps.Qty))
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtensionSimple"
End Property
Public Property Get ExtensionSimpleF() As String
    On Error GoTo errHandler
    ExtensionSimpleF = Format(ExtensionSimple / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtensionSimpleF"
End Property

Public Property Get Discount() As Double
    On Error GoTo errHandler
    Discount = mudtProps.Discount
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Discount"
End Property
Public Property Let Discount(val As Double)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
        mudtProps.Discount = val
        If (val < 0) Then
            mobjValid.RuleBroken "Discount", True
        Else
            mobjValid.RuleBroken "Discount", False
        End If
        mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Discount(val)", val
End Property
Public Function SetDiscount(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim dblTEMP As Double
    dblTEMP = mudtProps.Discount
    bOK = SetField_DOUBLE(mudtProps.Discount, val, "DISCOUNT", mcolStack.Count)
    If bOK And (dblTEMP <> mudtProps.Discount) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("DISCOUNT")
    End If
    SetDiscount = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetDiscount(val)", val
End Function
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
    ErrorIn "a_COL.DiscountF"
End Function
Public Property Get DiscountAmount() As Long
    On Error GoTo errHandler
    DiscountAmount = CLng(ExtensionSimple - ExtensionInclDeposit)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.DiscountAmount"
End Property
Public Property Get DiscountAmountF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        DiscountAmountF = Format(CDbl(DiscountAmount / mCOParent.ForeignCurrency.Divisor), mCOParent.ForeignCurrency.FormatString)
    Else
        DiscountAmountF = Format(CDbl(DiscountAmount / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.DiscountAmountF(bForeign)", bForeign
End Property

Public Property Get DepositF() As String
    On Error GoTo errHandler
    DepositF = Format(mudtProps.Deposit / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.DepositF"
End Property
Public Property Get Deposit() As Long
    On Error GoTo errHandler
    Deposit = mudtProps.Deposit
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Deposit"
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
    ErrorIn "a_COL.Deposit(val)", val
End Property
Public Function SetDeposit(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.Deposit
    bOK = SetField_LONG(mudtProps.Deposit, val, "DEPOSIT", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.Deposit) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("DEPOSIT")
    End If
    SetDeposit = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetDeposit(val)", val
End Function
Public Property Get VATAmount() As Long
    On Error GoTo errHandler
    VATAmount = CLng(CDbl((VATRate / (1 + VATRate))) * ExtensionInclDeposit)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.VATAmount"
End Property

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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetNote(val)", val
End Function
Public Property Get PriceF() As String
    On Error GoTo errHandler
    PriceF = Format(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.PriceF"
End Property
Public Property Get PriceExVAT() As Long
    On Error GoTo errHandler
    PriceExVAT = Round(mudtProps.Price * (100 / (100 + IIf(mudtProps.VATRate = 0, oPC.Configuration.VATRate, mudtProps.VATRate))))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL.PriceExVAT"
End Property

Public Function PriceExVatF() As String
    On Error GoTo errHandler
        PriceExVatF = Format(CDbl(Me.PriceExVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.PriceExVatF"
End Function
Public Property Get Ext_ExFF() As String
    On Error GoTo errHandler
        Ext_ExFF = Format(CDbl(mExt_Ex / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Ext_ExFF"
End Property

Public Property Get Price() As Long
    On Error GoTo errHandler
    Price = mudtProps.Price
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Price"
End Property
Public Property Let Price(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.Price Then
        mudtProps.Price = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Price(val)", val
End Property
Public Function SetPrice(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    If mCOParent.OrderType = enWant Then
        SetPrice = True
        Exit Function
    End If
    lngTemp = mudtProps.Price
    bOK = SetField_LONG(mudtProps.Price, val, "PRICE", mcolStack.Count)
        bOK = ValidateObject("PRICE")
    SetPrice = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetPrice(val)", val
End Function
Public Property Get ExtraChargeF() As String
    On Error GoTo errHandler
    ExtraChargeF = Format(mudtProps.ExtraCharge / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtraChargeF"
End Property

Public Property Get ExtraCharge() As Long
    On Error GoTo errHandler
    ExtraCharge = mudtProps.ExtraCharge
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtraCharge"
End Property
Public Function SetExtraCharge(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    If mCOParent.OrderType = enWant Then
        SetExtraCharge = True
        Exit Function
    End If
    lngTemp = mudtProps.ExtraCharge
    bOK = SetField_LONG(mudtProps.ExtraCharge, val, "EXTRACHARGE", mcolStack.Count)
        bOK = ValidateObject("EXTRACHARGE")
    SetExtraCharge = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetExtraCharge(val)", val
End Function

Public Property Get ExtraVATRate() As Double
    On Error GoTo errHandler
    ExtraVATRate = mudtProps.ExtraVATRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtraVATRate"
End Property
Public Property Get ExtraVATRateF() As String
    On Error GoTo errHandler
    ExtraVATRateF = Format(mudtProps.ExtraVATRate, gPercentFormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtraVATRateF"
End Property

Public Property Let ExtraVATRate(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ExtraVATRate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ExtraVATRate(val)", val
End Property





Public Property Let Qty(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.Qty Then
        mudtProps.Qty = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Qty(val)", val
End Property
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
    ErrorIn "a_COL.SetQty(val)", val
End Function
Public Property Get Qty() As Long
    On Error GoTo errHandler
    Qty = mudtProps.Qty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Qty"
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
    ErrorIn "a_COL.QtyF"
End Property

'=============
Public Property Let QtyFirm(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.QtyFirm Then
        mudtProps.QtyFirm = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.QtyFirm(val)", val
End Property
Public Function SetQtyFirm(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.QtyFirm
    bOK = SetField_LONG(mudtProps.QtyFirm, val, "QtyFirm", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.QtyFirm) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QtyFirm")
    End If
    SetQtyFirm = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetQtyFirm(val)", val
End Function
Public Property Get QtyFirm() As Long
    On Error GoTo errHandler
    QtyFirm = mudtProps.QtyFirm
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.QtyFirm"
End Property
Public Property Get QtyFirmF() As String
    On Error GoTo errHandler
    QtyFirmF = Format(mudtProps.QtyFirm, "##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.QtyFirmF"
End Property
'=================
Public Property Let QtySS(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.QtySS Then
        mudtProps.QtySS = val
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.QtySS(val)", val
End Property
Public Function SetQtySS(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.QtySS
    bOK = SetField_LONG(mudtProps.QtySS, val, "QtySS", mcolStack.Count)
    If bOK And (lngTemp <> mudtProps.QtySS) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QtySS")
    End If
    SetQtySS = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetQtySS(val)", val
End Function
Public Property Get QtySS() As Long
    On Error GoTo errHandler
    QtySS = mudtProps.QtySS
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.QtySS"
End Property
Public Property Get QtySSF() As String
    On Error GoTo errHandler
    QtySSF = Format(mudtProps.QtySS, "##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.QtySSF"
End Property
'===================













Public Property Get DepositStatus() As String
    On Error GoTo errHandler
    DepositStatus = mudtProps.DepositStatus
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.DepositStatus"
End Property
Public Property Let DepositStatus(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    If val = "C" Or val = "P" Or val = "R" Or val = "O" Then   'Credit to account or Paid back or Redeemed or Outstanding
        mudtProps.DepositStatus = val
        mudtProps.IsDirty = True
    Else
        Err.Raise 383
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.DepositStatus(val)", val
End Property

Public Property Get QtyDispatched() As Long
    On Error GoTo errHandler
    QtyDispatched = mudtProps.QtyDispatched
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.QtyDispatched"
End Property
Public Property Get QtyDispatchedF() As String
    On Error GoTo errHandler
    QtyDispatchedF = Format(mudtProps.QtyDispatched, "##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.QtyDispatchedF"
End Property
Public Property Get ActionTaken() As enumActionTaken
    On Error GoTo errHandler
    ActionTaken = mudtProps.ActionTaken
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ActionTaken"
End Property
Public Function SetActionTaken(val As enumActionTaken)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    If mudtProps.ActionTaken <> val Then
        mudtProps.ActionTaken = val
        SetDirty True
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetActionTaken(val)", val
End Function
Public Property Get QtyOrdered_QtyDispatched() As String
    On Error GoTo errHandler
    If oPC.AllowsSSInvoicing Then
        QtyOrdered_QtyDispatched = QtyFirmF & "/" & QtySSF & " (" & QtyDispatchedF & ")"
    Else
        QtyOrdered_QtyDispatched = QtyF & " (" & QtyDispatchedF & ")"
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.QtyOrdered_QtyDispatched"
End Property
Public Property Get VATRate() As Double
    On Error GoTo errHandler
    VATRate = mudtProps.VATRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.VATRate"
End Property
Public Property Get VATRateF() As String
    On Error GoTo errHandler
    VATRateF = Format(mudtProps.VATRate, gPercentFormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.VATRateF"
End Property

Public Property Let VATRate(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.VATRate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.VATRate(val)", val
End Property

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
    ErrorIn "a_COL.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Invalid price", "PRICE"
    colClassors.Add "Invalid quantity", "QTY"
    colClassors.Add "Invalid ETA", "ETA"
    colClassors.Add "Invalid discount", "DISCOUNT"
    colClassors.Add "Invalid deposit", "DEPOSIT"
    colClassors.Add "Invalid extra charge", "EXTRACHARGE"
    colClassors.Add "Missing line reference", "REF"
    colClassors.Add "A numeric value is too great for calculation", "OVERFLOW"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.LoadClassorsCollection"
End Sub
Private Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean

    bValid = True
    If FNS(mudtProps.Ref) = "" And oPC.Configuration.EnforceCOLRef And Me.ServiceItem = False Then
        mobjValid.RuleBroken "REF", True
        If pFld = "REF" Then bValid = False
    Else
        mobjValid.RuleBroken "REF", False
    End If
    
    If mudtProps.Price > 0 Or oPC.AllowZeropricedPOLines Then
        mobjValid.RuleBroken "PRICE", False
    Else
        If pFld = "PRICE" Then bValid = False
        mobjValid.RuleBroken "PRICE", True
    End If
    
    If mudtProps.ExtraCharge <= 0 And FNS(mudtProps.ExtraCode) > "" Then
        If pFld = "EXTRACHARGE" Then bValid = False
        mobjValid.RuleBroken "EXTRACHARGE", True
    Else
        mobjValid.RuleBroken "EXTRACHARGE", False
    End If
    
    If mudtProps.Qty >= 0 Or mudtProps.QtySS > 0 Or mudtProps.QtyFirm > 0 Then
        mobjValid.RuleBroken "QTY", False
    Else
        If pFld = "QTY" Then bValid = False
        mobjValid.RuleBroken "QTY", True
    End If
    
    If (mudtProps.Discount < 0 Or mudtProps.Discount > 100) Then
        If pFld = "DISCOUNT" Then bValid = False
        mobjValid.RuleBroken "DISCOUNT", True
    Else
        mobjValid.RuleBroken "DISCOUNT", False
    End If
    
    If (mudtProps.ETA < Date Or mudtProps.ETA = CDate(0)) And mCOParent.OrderType = enNormalCO Then
        If pFld = "ETA" Then bValid = False
        mobjValid.RuleBroken "ETA", True
    Else
        mobjValid.RuleBroken "ETA", False
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
    ErrorIn "a_COL.ValidateObject(pFld)", pFld
End Function
Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetDirty(pVal)", pVal
End Sub


Public Sub GetStatus()
    On Error GoTo errHandler
    mobjValid.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.GetSTatus"
End Sub
Public Function CancelLine()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
'    oSM.CancelCOL Me.COLineID ' Me.PID, CLng(Me.Qty)
    lngResult = oSM.MarkCOLStatus(COLineID, "CAN")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.CancelLine"
End Function

Public Function FulfilLine()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    lngResult = oSM.MarkCOLStatus(COLineID, "FUL")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.FulfilLine"
End Function

Public Sub SetReplacementForLineID(val As Long)
    On Error GoTo errHandler
    
    mudtProps.Replacementfor = val

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.SetReplacementForLineID(val)", val
End Sub


Public Property Get ForeignPrice() As Long
    On Error GoTo errHandler
    ForeignPrice = mudtProps.ForeignPrice
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ForeignPrice"
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
    ErrorIn "a_COL.ForeignPrice(val)", val
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
    ErrorIn "a_COL.SetForeignPrice(val)", val
End Function
Public Property Get ForeignPriceF() As String
    On Error GoTo errHandler
Dim oFC As New a_Currency
    If mudtProps.FCID > 0 Then
        Set oFC = oPC.Configuration.Currencies.FindCurrencyByID(mudtProps.FCID)
        ForeignPriceF = Format(CDbl(mudtProps.ForeignPrice / oFC.Divisor), oFC.FormatString)
    Else
        ForeignPriceF = ""
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.ForeignPriceF"
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
    ErrorIn "a_COL.SetFCFactor(val)", val
End Function

Public Property Get FCFactor() As Double
    On Error GoTo errHandler
    FCFactor = mudtProps.FCFactor
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.FCFactor"
End Property
Public Property Get FCFactorF() As String
    On Error GoTo errHandler
    FCFactorF = Format(FCFactor, "##0.####")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.FCFactorF"
End Property
Public Property Get FCFactorInvF() As String
    FCFactorInvF = Format(Round(1# / FCFactor, 3), "##0.###")
End Property

Public Property Get FCID() As Long
    On Error GoTo errHandler
    FCID = mudtProps.FCID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.FCID"
End Property
Public Property Let FCID(val As Long)
    mudtProps.FCID = val
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
    ErrorIn "a_COL.Properties(pSrt)", pSrt
End Property

Public Property Get Sequence() As Long
    On Error GoTo errHandler
    Sequence = mudtProps.Sequence
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.Sequence"
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
    ErrorIn "a_COL.Sequence(val)", val
End Property

