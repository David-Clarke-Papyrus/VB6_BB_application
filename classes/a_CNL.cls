VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_CNL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Valid(msg As String)
Private mudtProps As CNLProps
Private mcolStack As Collection
Private mobjParent As ch_CNL
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassors As Collection
Private mCNParent As a_CN
Private mKey As String
Private moProd As a_Product

Private mFPLessDisc As Long
Private mLPLessDisc As Long

Private mFPLessDiscExt As Long
Private mLPLessDiscExt As Long

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

Public Property Get CNParent() As a_CN
    Set CNParent = mCNParent
End Property
Public Property Get CodeForEditing() As String
    If Not IsISBN10(mudtProps.ProductCode) And FNS(mudtProps.ProductCode) > "" Then
        CodeForEditing = FNS(mudtProps.ProductCode)
    Else
        If IsISBN13(FNS(mudtProps.EAN), True) Then
            CodeForEditing = FNS(mudtProps.EAN)
        Else
            If FNS(mudtProps.ProductCode) > "" Then
                CodeForEditing = FNS(mudtProps.ProductCode)
            Else
                CodeForEditing = FNS(mudtProps.EAN)
            End If
        End If
    End If
    
End Property
Private Function GetState() As String
    On Error GoTo errHandler
  Dim mudtData As CNLData
  
  LSet mudtData = mudtProps
  GetState = mudtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim mudtData As CNLData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.SetSuperState(buffer)", buffer
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddCNLine Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.ApplyEdit"
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
    ErrorIn "a_CNL.BeginEdit"
End Sub
'Public Sub SetAsNEW()
'    mudtProps.IsNew = True
'End Sub
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
    ErrorIn "a_CNL.CancelEdit"
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    LoadClassorsCollection
    mobjValid.RuleBroken "Price", True
    mudtProps.DiscountRate = 0
    mudtProps.Qty = 1
    mudtProps.IsNew = True
    Set moProd = New a_Product
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.Class_Initialize"
End Sub
Friend Sub Initialize(Parent As ch_CNL)
    On Error GoTo errHandler
    If mudtProps.IsNew Then
        Set mobjParent = Parent
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.Initialize(Parent)", Parent
End Sub
Friend Sub SetParentCN(pParentCN As a_CN)
    On Error GoTo errHandler
    Set mCNParent = pParentCN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.SetParentCN(pParentCN)", pParentCN
End Sub
Friend Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.IsValid"
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
    ErrorIn "a_CNL.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.mobjValid_Status(pMsg)", pMsg
End Sub
Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.Delete"
End Sub
Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.IsDirty"
End Property
Public Property Get INVLineID() As Long
    On Error GoTo errHandler
    INVLineID = mudtProps.INVLineID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.INVLineID"
End Property

Public Property Let INVLineID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.INVLineID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.INVLineID(val)", val
End Property
Public Property Get INVLineCode() As String
    On Error GoTo errHandler
    INVLineCode = FNS(mudtProps.INVLineCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.INVLineCode"
End Property

Public Property Let INVLineCode(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.INVLineCode = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.INVLineCode(val)", val
End Property
Public Property Get ProductCode() As String
    ProductCode = FNS(mudtProps.ProductCode)
End Property
Public Property Get EAN() As String
    EAN = FNS(mudtProps.EAN)
End Property
Public Property Let EAN(val As String)
    mudtProps.EAN = val
End Property

Public Property Get CNLineID() As Long
    On Error GoTo errHandler
    CNLineID = mudtProps.CNLineID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.CNLineID"
End Property

Public Property Let CNLineID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CNLineID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.CNLineID(val)", val
End Property
Public Property Get TRID() As Long
    On Error GoTo errHandler
  TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.TRID"
End Property
Public Property Let TRID(val As Long)
    On Error GoTo errHandler
    mudtProps.TRID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.TRID(val)", val
End Property
Public Property Get PID() As String
    On Error GoTo errHandler
    PID = FNS(mudtProps.PID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.PID"
End Property
Public Property Get InvLineRef() As String
    InvLineRef = FNS(mudtProps.InvLineRef)
End Property
Public Property Let PID(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    SetDirty (mudtProps.PID <> val)
    mudtProps.PID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.PID(val)", val
End Property
Public Property Get CopyID() As Long
    On Error GoTo errHandler
    CopyID = mudtProps.PIID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.CopyID"
End Property
Public Property Let CopyID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    SetDirty (mudtProps.PIID <> val)
    mudtProps.PIID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.CopyID(val)", val
End Property
Public Property Get Product() As a_Product
    Set Product = moProd
End Property
Public Function SetLineProduct(Optional PID As String, Optional pCode As String) As Boolean
    On Error GoTo errHandler
    If Not moProd.IsNew Then
        moProd.PrepareForReload
    End If
    With moProd
        If .Load(PID, 0, Trim$(pCode)) <> 99 Then   'product found
            Me.Title = .TitleAuthorPublisher
            Me.PID = .PID
          '  Me.Price = .RRP
            If Not .DefaultCopy Is Nothing Then
                Me.CopyID = .DefaultCopy.ID
            End If
            Me.ProductCode = .code
            Me.ProductCodeF = .CodeF
            Me.EAN = .EAN
            SetLineProduct = True
        Else
            SetLineProduct = False   'Not found
        End If
    End With

    mobjValid.GetStatus
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.SetLineProduct(PID,pCODE)", Array(PID, pCode)
End Function
Public Property Get Note() As String
    On Error GoTo errHandler
    Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.Note"
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
    ErrorIn "a_CNL.Note(val)", val
End Property

'Public Property Get CNLineCode() As String
'    CNLineCode = FNS(mudtProps.CNLineCode)
'End Property
'Public Property Let CNLineCode(Val As String)
'  If mcolStack.Count = 0 Then Err.Raise 383
'    If mudtProps.CNLineCode <> Val Then
'        mudtProps.CNLineCode = Val
'        SetDirty True
'    End If
'End Property

'Public Property Get ProductCode() As String
'    On Error GoTo ErrHandler
'    ProductCode = FNS(mudtProps.ProductCode)
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_CNL.ProductCode"
'End Property
Public Property Let ProductCode(val As String)
    On Error GoTo errHandler
    mudtProps.ProductCode = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.ProductCode(val)", val
End Property
Sub TestEvent()
    On Error GoTo errHandler
RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.TestEvent"
End Sub
'Public Property Let ProductCode(val As String)
'Dim oCode As z_ProdCode
'  If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.ProductCode = val
'    Set oCode = New z_ProdCode
'    oCode.Load val
'    mudtProps.ProductCodeF = oCode.FormatCode
'    mudtProps.IsDirty = True
'    Set oCode = Nothing
'End Property

Public Property Get ProductCodeF() As String
    On Error GoTo errHandler
    ProductCodeF = FNS(mudtProps.ProductCodeF)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.ProductCodeF"
End Property
Public Property Let ProductCodeF(val As String)
    On Error GoTo errHandler
    mudtProps.ProductCodeF = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.ProductCodeF(val)", val
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
    TitleAuthorPublisher = tmp

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.TitleAuthorPublisher"
End Property
Public Property Get Publisher() As String
    On Error GoTo errHandler
    Publisher = FNS(mudtProps.Publisher)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.Publisher"
End Property
Public Property Get Title() As String
    On Error GoTo errHandler
    Title = FNS(mudtProps.Title)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.Title"
End Property
Public Property Get TitleF(pLen As Integer) As String
    On Error GoTo errHandler
    TitleF = Left(Title, pLen) & IIf(Len(Title) > pLen, "...", "")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.TitleF(pLen)", pLen
End Property
Public Property Let Title(val As String)
    On Error GoTo errHandler
    mudtProps.Title = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.Title(val)", val
End Property

Public Property Get MainAuthor() As String
    On Error GoTo errHandler
    MainAuthor = FNS(mudtProps.Author)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.MainAuthor"
End Property
Public Property Let MainAuthor(val As String)
    On Error GoTo errHandler
    mudtProps.Author = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.MainAuthor(val)", val
End Property
'---------------------------------
Public Property Get VATRate() As Double
    On Error GoTo errHandler
    VATRate = mudtProps.VATRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.VATRate"
End Property
Public Property Get VATRateF() As String
    On Error GoTo errHandler
    VATRateF = Format(mudtProps.VATRate, gPercentFormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.VATRateF"
End Property
Public Property Let VATRate(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.VATRate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.VATRate(val)", val
End Property
Public Function SetVATRate(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim dblTEMP As Double
    dblTEMP = mudtProps.VATRate
    bOK = SetField_DOUBLE(mudtProps.VATRate, val, "VATRATE", mcolStack.Count)
    If bOK And (dblTEMP <> mudtProps.VATRate) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("VATRATE")
    End If
    SetVATRate = bOK
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.SetVATRate(val)", val
End Function


'----------------------
Public Function CalculateLine()
    On Error GoTo errHandler
'First in local currency
    mobjValid.RuleBroken "OVERFLOW", False
    mLPLessDiscExt = Round(CDbl(mudtProps.InvPrice * mudtProps.Qty) * (CDbl(100) - mudtProps.DiscountRate) / 100)
    mLPDiscExt = (CDbl(mudtProps.InvPrice * mudtProps.Qty)) - mLPLessDiscExt
    mLPLessDiscExtVAT = Round(mLPLessDiscExt - (CDbl(mLPLessDiscExt) * 100 / (CDbl(100) + mudtProps.VATRate)))
'=======
    mLPLessDiscExt = Round17(CDbl(mudtProps.InvPrice * mudtProps.Qty) * (CDbl(100) - mudtProps.DiscountRate) / 100)
    mLPDiscExt = (CDbl(mudtProps.InvPrice * mudtProps.Qty)) - mLPLessDiscExt
    mLPLessDiscExtVAT = Round17(mLPLessDiscExt - (CDbl(mLPLessDiscExt) * 100 / (CDbl(100) + mudtProps.VATRate)))
'=======
    mLPLessDiscExt = Round17(CDbl(mudtProps.InvPrice * mudtProps.Qty) * (CDbl(100) - mudtProps.DiscountRate) / 100)
    mLPDiscExt = (CDbl(mudtProps.InvPrice * mudtProps.Qty)) - mLPLessDiscExt
    mLPLessDiscExtVAT = Round17(mLPLessDiscExt - (CDbl(mLPLessDiscExt) * 100 / (CDbl(100) + mudtProps.VATRate)))
    mLPLessDiscExtNonVAT = mLPLessDiscExt - mLPLessDiscExtVAT
    
'Then in Foreign if appropriate
    If mCNParent Is Nothing Then Exit Function
    If Not mCNParent.ForeignCurrency Is Nothing Then
        mFPLessDiscExt = Round(CDbl(mudtProps.InvForeignPrice * mudtProps.Qty) * ((CDbl(100) - mudtProps.DiscountRate) / 100))
        mFPDiscExt = (CDbl(mudtProps.InvForeignPrice * mudtProps.Qty)) - mFPLessDiscExt
        mFPLessDiscExtVAT = Round(mFPLessDiscExt - (mFPLessDiscExt * 100 / (CDbl(100) + mudtProps.VATRate)))
''=======
'        mFPLessDiscExt = Round17(CDbl(mudtProps.InvForeignPrice * mudtProps.Qty) * ((CDbl(100) - mudtProps.DiscountRate) / 100))
'        mFPDiscExt = (CDbl(mudtProps.InvForeignPrice * mudtProps.Qty)) - mFPLessDiscExt
'        mFPLessDiscExtVAT = Round17(mFPLessDiscExt - (mFPLessDiscExt * 100 / (CDbl(100) + mudtProps.VATRate)))
''=======
'        mFPLessDiscExt = Round17(CDbl(mudtProps.InvForeignPrice * mudtProps.Qty) * ((CDbl(100) - mudtProps.DiscountRate) / 100))
'        mFPDiscExt = (CDbl(mudtProps.InvForeignPrice * mudtProps.Qty)) - mFPLessDiscExt
'        mFPLessDiscExtVAT = Round17(mFPLessDiscExt - (mFPLessDiscExt * 100 / (CDbl(100) + mudtProps.VATRate)))
        mFPLessDiscExtNonVAT = mFPLessDiscExt - mFPLessDiscExtVAT
    End If
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'overflow
        mobjValid.RuleBroken "OVERFLOW", True
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.CalculateLine"
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
    ErrorIn "a_CNL.PLessDisc(val)", val
End Property
Public Property Get PLessDiscF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscF = Format(CDbl(mFPLessDisc / mCNParent.ForeignCurrency.Divisor), mCNParent.ForeignCurrency.FormatString)
    Else
        PLessDiscF = Format(CDbl(mLPLessDisc / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.PLessDiscF(val)", val
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
    ErrorIn "a_CNL.PLessDiscExt(val)", val
End Property

Public Property Get PLessDiscExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscExtF = Format(CDbl(mFPLessDiscExt / mCNParent.ForeignCurrency.Divisor), mCNParent.ForeignCurrency.FormatString)
    Else
        PLessDiscExtF = Format(CDbl(mLPLessDiscExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.PLessDiscExtF(val)", val
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
    ErrorIn "a_CNL.PDiscExt(val)", val
End Property
Public Property Get PDiscExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PDiscExtF = Format(CDbl(mFPDiscExt / mCNParent.ForeignCurrency.Divisor), mCNParent.ForeignCurrency.FormatString)
    Else
        PDiscExtF = Format(CDbl(mLPDiscExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.PDiscExtF(val)", val
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
    ErrorIn "a_CNL.PLessDiscExtVAT(val)", val
End Property
Public Property Get PLessDiscExtVATF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscExtVATF = Format(CDbl(mFPLessDiscExtVAT / mCNParent.ForeignCurrency.Divisor), mCNParent.ForeignCurrency.FormatString)
    Else
        PLessDiscExtVATF = Format(CDbl(mLPLessDiscExtVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.PLessDiscExtVATF(val)", val
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
    ErrorIn "a_CNL.PLessDiscExtNonVAT(val)", val
End Property
Public Property Get PLessDiscExtNonVATF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscExtNonVATF = Format(CDbl(mFPLessDiscExtNonVAT / mCNParent.ForeignCurrency.Divisor), mCNParent.ForeignCurrency.FormatString)
    Else
        PLessDiscExtNonVATF = Format(CDbl(mLPLessDiscExtNonVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.PLessDiscExtNonVATF(val)", val
End Property

'Public Property Get VATValue() As Long
'    VATValue = (VATRate / (100 + VATRate)) * ExtensionInclDeposit
'End Property
'Public Property Get ExtensionExVAT() As Long
'    ExtensionExVAT = (CLng((mudtProps.InvPrice * mudtProps.Qty) * ((100 - mudtProps.Discount) / 100))) * (100 / (100 + mudtProps.VATRate))
'End Property
'Public Property Get SalesValue() As Long
'    SalesValue = (1 - (VATRate / (100 + VATRate))) * ExtensionInclDeposit
'End Property
'Public Property Get ExtensionInclDeposit() As Long
'    ExtensionInclDeposit = CLng(CLng(mudtProps.InvPrice * mudtProps.Qty) * ((100 - mudtProps.Discount) / 100))
'End Property
'Public Property Get ExtensionInclDepositF() As String
'    ExtensionInclDepositF = Format(ExtensionInclDeposit / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get Extension() As Long
'    Extension = CLng(CLng(mudtProps.InvPrice * mudtProps.Qty) * (100 - mudtProps.Discount) / 100)
'End Property
'Public Property Get ExtensionF() As String
'    ExtensionF = Format(Extension / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get ExtensionSimple() As Long
'    ExtensionSimple = CLng(CLng(mudtProps.InvPrice * mudtProps.Qty))
'End Property
'Public Property Get ExtensionSimpleF() As String
'    ExtensionSimpleF = Format(ExtensionSimple / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get DiscountValue() As Long
'    DiscountValue = (mudtProps.InvPrice * mudtProps.Discount) / 100
'End Property


Public Property Get Discount() As Double
    On Error GoTo errHandler
    Discount = mudtProps.DiscountRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.Discount"
End Property
Public Property Get InvLineDate() As Date
    On Error GoTo errHandler
    InvLineDate = mudtProps.InvLineDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.InvLineDate"
End Property
Public Property Let InvLineDate(val As Date)
    On Error GoTo errHandler
    mudtProps.InvLineDate = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.InvLineDate(val)", val
End Property
Public Property Let Discount(val As Double)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
        mudtProps.DiscountRate = val
        If (val < 0) Then
            mobjValid.RuleBroken "Discount", True
        Else
            mobjValid.RuleBroken "Discount", False
        End If
        mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.Discount(val)", val
End Property
Public Function SetDiscount(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim dblTEMP As Double
    dblTEMP = mudtProps.DiscountRate
    bOK = SetField_DOUBLE(mudtProps.DiscountRate, val, "DISCOUNT", mcolStack.Count)
    If bOK And (dblTEMP <> mudtProps.DiscountRate) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("DISCOUNT")
    End If
    SetDiscount = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.SetDiscount(val)", val
End Function
Public Function DiscountPercentF() As String
    On Error GoTo errHandler
    If mudtProps.DiscountRate = 0 Then
        DiscountPercentF = ""
    Else
        DiscountPercentF = Format(mudtProps.DiscountRate, "0.00\%")
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.DiscountPercentF"
End Function
'Public Property Get DepositF() As String
'    DepositF = Format(mudtProps.Deposit / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get Deposit() As Long
'    Deposit = mudtProps.Deposit
'End Property
'Public Property Let Deposit(Val As Long)
'  If mcolStack.Count = 0 Then Err.Raise 383
'    If Val <> mudtProps.Deposit Then
'        mudtProps.Deposit = Val
'        SetDirty True
'    End If
'End Property
'Public Function SetDeposit(Val As String) As Boolean
'Dim bOK As Boolean
'Dim lngTEMP As Long
'    lngTEMP = mudtProps.Deposit
'    bOK = SetField_LONG(mudtProps.Deposit, Val, "DEPOSIT")
'    If bOK And (lngTEMP <> mudtProps.Deposit) Then 'only validate if conversion is OK and the new value is other than the original
'        bOK = ValidateObject("DEPOSIT")
'    End If
'    SetDeposit = bOK
'End Function
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
Public Function PriceF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        PriceF = Format(CDbl(mudtProps.InvForeignPrice / mCNParent.ForeignCurrency.Divisor), mCNParent.ForeignCurrency.FormatString)
    Else
        PriceF = Format(CDbl(mudtProps.InvPrice / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.PriceF(bForeign)", bForeign
End Function
Public Property Get PriceExVAT() As Long
    On Error GoTo errHandler
    PriceExVAT = Round(Me.Price * (100 / (100 + IIf(mudtProps.VATRate = 0, oPC.Configuration.VATRate, mudtProps.VATRate))))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.PriceExVAT"
End Property

Public Function PriceExVatF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        'surely this wont be used
       ' PriceExVatF = Format(CDbl(mudtProps.ForeignPrice * (100 / 114) / mCNParent.ForeignCurrency.Divisor), mCNParent.ForeignCurrency.FormatString)
    Else
        PriceExVatF = Format(CDbl(Me.PriceExVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.PriceExVatF(bForeign)", bForeign
End Function

Public Property Get Price() As Long
    On Error GoTo errHandler
    Price = mudtProps.InvPrice
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.Price"
End Property
Public Property Get InvForeignPrice() As Long
    InvForeignPrice = mudtProps.InvForeignPrice
End Property
Public Property Let InvPrice(val As Long)
    On Error GoTo errHandler
Dim bOK As Boolean
  If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.InvPrice Then
        mudtProps.InvPrice = val
        bOK = ValidateObject("PRICE")
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.InvPrice(val)", val
End Property

Public Function SetPrice(val As String) As Boolean
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.InvPrice
    bOK = SetField_strAsCurrencyToLong(mudtProps.InvPrice, val, mcolStack.Count, "PRICE", oPC.Configuration.CaptureDecimal, 100)
    If bOK And (lngTemp <> mudtProps.InvPrice) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("PRICE")
    End If
    SetPrice = bOK
End Function
Public Property Get ServiceItem() As Boolean
    On Error GoTo errHandler
    ServiceItem = mudtProps.ServiceItem
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.ServiceItem"
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
    ErrorIn "a_CNL.Qty(val)", val
End Property
Public Function SetQty(val As String, pMax As Long) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim bOK2 As Boolean
Dim lngTemp As Long

    lngTemp = mudtProps.Qty
    
    bOK = SetField_LONG(mudtProps.Qty, val, "QTY", mcolStack.Count)
    If bOK And mudtProps.Qty > pMax Then
        bOK = False
    End If
    If bOK And (lngTemp <> mudtProps.Qty) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QTY")
    End If
    SetQty = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.SetQty(val)", val
End Function
Public Function SetQtyDam(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim bOK2 As Boolean
Dim lngTemp As Long

    lngTemp = mudtProps.QtyDAM
    
    bOK = SetField_LONG(mudtProps.QtyDAM, val, "QTYDAM", mcolStack.Count)
    If bOK And mudtProps.QtyDAM > mudtProps.Qty Then
        bOK = False
    End If
    If bOK And (lngTemp <> mudtProps.QtyDAM) Then  'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QTYDAM")
    End If
    SetQtyDam = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.SetQtyDam(val)", val
End Function

Public Property Get Qty() As Long
    On Error GoTo errHandler
    Qty = mudtProps.Qty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.Qty"
End Property
Public Property Get QtyDAM() As Long
    On Error GoTo errHandler
    QtyDAM = mudtProps.QtyDAM
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.QtyDam"
End Property
Public Property Get QtyF() As String
    On Error GoTo errHandler
        QtyF = Format(Qty, "##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.QtyF"
End Property
Public Property Get QtyDamF() As String
    On Error GoTo errHandler
        QtyDamF = Format(QtyDAM, "##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.QtyDamF"
End Property
Public Property Get QtyComboF() As String
    On Error GoTo errHandler
        If QtyDAM <> Qty And QtyDAM <> 0 Then
            QtyComboF = Format(Qty, "##0") & "/" & Format(QtyDAM, "##0")
        Else
            QtyComboF = Format(Qty, "##0")
        End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.QtyComboF"
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
    ErrorIn "a_CNL.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Invalid price", "PRICE"
    colClassors.Add "Invalid quantity", "QTY"
    colClassors.Add "Invalid discount", "DISCOUNT"
    colClassors.Add "Numeric overflow, value captured too large", "OVERFLOW"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.LoadClassorsCollection"
End Sub
Private Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean
    bValid = True
   ' Discount = True
    If mudtProps.InvPrice >= 0 Then
        mobjValid.RuleBroken "PRICE", False
    Else
        If pFld = "PRICE" Then bValid = False
        mobjValid.RuleBroken "PRICE", True
    End If
    If mudtProps.Qty >= 0 Then
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
'H:
'    MsgBox "CheckOK: " & or
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.ValidateObject(pFld)", pFld
End Function
Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.SetDirty(pVal)", pVal
End Sub


Public Sub GetStatus()
    On Error GoTo errHandler
    ValidateObject ""
    mobjValid.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL.GetSTatus"
End Sub
Public Property Get Properties(pSrt As enSortField) As Variant
    On Error GoTo errHandler
    Select Case pSrt
    Case enTitle
        Properties = Me.Title
    Case enCode
        Properties = Me.ProductCode
    End Select
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.Properties(pSrt)", pSrt
End Property

