VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_POL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Valid(msg As String)
Event ValueChanges()
Private mudtProps As POLProps
Private mcolStack As Collection
Private mobjParent As ch_POL
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassors As Collection
Private mPOParent As a_PO
Private mKey As String
Private moProd As a_Product
'Private mcCOLsPerPID As c_COLSPerPID
Private mFPExt As Long
Private mLPExt As Long

Private mFPLessDisc As Long
Private mLPLessDisc As Long
Private mLPLessDiscVAT As Long
Private mLPLessDiscNONVAT As Long
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
        Properties = Me.ProductCode
    Case enRef
        Properties = Me.Ref
    End Select
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Properties(pSrt)", pSrt
End Property

Public Property Get Sequence() As Long
    On Error GoTo errHandler
    Sequence = mudtProps.Sequence
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Sequence"
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
    ErrorIn "a_POL.Sequence(val)", val
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

Public Property Get CodeF() As String
    On Error GoTo errHandler
    CodeF = FNS(mudtProps.ProductCodeF)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.CodeF"
End Property
Public Property Get Replacementfor() As Long
    Replacementfor = FND(mudtProps.Replacementfor)
End Property

Private Function GetState() As String
    On Error GoTo errHandler
  Dim mudtData As POLData
  
  LSet mudtData = mudtProps
  GetState = mudtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim mudtData As POLData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
    SetState buffer
    mobjValid.RuleBroken "QtySS", False
    mobjValid.RuleBroken "QtyFirm", False
    mobjValid.RuleBroken "Price", False
    mobjValid.RuleBroken "PID", False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.SetSuperState(buffer)", buffer
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
    ErrorIn "a_POL.GetSTatus"
End Function

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddSOLine Me
    Set mobjParent = Nothing
  End If
  PO.CalculateTotals
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.ApplyEdit"
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
    ErrorIn "a_POL.BeginEdit"
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
    ErrorIn "a_POL.CancelEdit"
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    Set moProd = New a_Product
    Set colClassors = New Collection
    LoadClassorsCollection
    mobjValid.RuleBroken "QtySS", True
    mobjValid.RuleBroken "QtyFirm", True
    mobjValid.RuleBroken "Price", IIf(oPC.AllowZeropricedPOLines, False, True)
    mobjValid.RuleBroken "PID", True
    mudtProps.Discount = 0
    mudtProps.QtyFirm = 1
    mudtProps.QtySS = 0
    mudtProps.VATRate = oPC.Configuration.VATRate
    mudtProps.IsNew = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Class_Initialize"
End Sub
Friend Sub Initialize(Parent As ch_POL)
    On Error GoTo errHandler
    If mudtProps.IsNew Then
        Set mobjParent = Parent
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Initialize(Parent)", Parent
End Sub
Friend Sub SetParentPO(pParentPO As a_PO)
    On Error GoTo errHandler
    Set mPOParent = pParentPO
    If FND(mudtProps.ETA) < Date Then
        mudtProps.ETA = DateAdd("d", mPOParent.Supplier.DefaultETA, Date)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.SetParentPO(pParentPO)", pParentPO
End Sub
Public Property Get PO() As a_PO
    On Error GoTo errHandler
    Set PO = mPOParent
    Exit Property
errHandler:
    ErrorIn "a_POL.PO"
End Property
Friend Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.IsValid"
End Property

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set colClassors = Nothing
    Set moProd = Nothing
    If mobjParent Is Nothing Then
        If mcolStack.Count > 0 Then _
            Err.Raise vbObjectError + 1001, , "State stack is not empty"
    End If
  Set mcolStack = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.mobjValid_Status(pMsg)", pMsg
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
    ErrorIn "a_POL.Delete"
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.IsDirty"
End Property


Public Property Get POLID() As Long
    On Error GoTo errHandler
    POLID = mudtProps.POLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.POLID"
End Property
Public Property Let POLID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.POLID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.POLID(val)", val
End Property

Public Property Let TRID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TRID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.TRID(val)", val
End Property
Public Property Get TRID() As Long
    On Error GoTo errHandler
  TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.TRID"
End Property

Public Property Get PID() As String
    On Error GoTo errHandler
    PID = FNS(mudtProps.PID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.PID"
End Property
Public Property Get Fulfilled() As String
    On Error GoTo errHandler
    Fulfilled = FNS(mudtProps.Fulfilled)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Fulfilled"
End Property
Public Property Let Fulfilled(val As String)
    mudtProps.Fulfilled = FNS(val)
    SetDirty True
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
    ErrorIn "a_POL.PID(val)", val
End Property
Public Property Get Product() As a_Product
    On Error GoTo errHandler
    Set Product = moProd
    Exit Property
errHandler:
    ErrorIn "a_POL.Product"
End Property
Public Property Let Product(val As a_Product)
    Set moProd = val
End Property

Public Property Get QtyFirm() As Long
    On Error GoTo errHandler
    QtyFirm = mudtProps.QtyFirm
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.QtyFirm"
End Property
Public Property Get QtyFirmF() As String
    On Error GoTo errHandler
    QtyFirmF = Format(mudtProps.QtyFirm, "##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.QtyFirmF"
End Property
Public Property Let QtyFirm(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.QtyFirm = val
    RecalculateLine
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.QtyFirm(val)", val
End Property
Public Property Get QtySS() As Long
    On Error GoTo errHandler
    QtySS = mudtProps.QtySS
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.QtySS"
End Property
Public Property Get QtySSF() As String
    On Error GoTo errHandler
    QtySSF = Format(mudtProps.QtySS, "##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.QtySSF"
End Property
Public Property Let QtySS(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.QtySS = val
    RecalculateLine
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.QtySS(val)", val
End Property
Public Property Get QtyReceivedSoFar() As Long
    On Error GoTo errHandler
    QtyReceivedSoFar = mudtProps.QtyReceivedSoFar
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.QtyReceivedSoFar"
End Property
Public Property Get ProductTypeID() As Long
    On Error GoTo errHandler
    ProductTypeID = mudtProps.ProductTypeID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.ProductTypeID"
End Property
Public Property Let ProductTypeID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ProductTypeID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.ProductTypeID(val)", val
End Property

Public Property Get ProductCodeF() As String
    On Error GoTo errHandler
    ProductCodeF = FNS(mudtProps.ProductCodeF)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.ProductCodeF"
End Property
Public Property Get ProductCodeForExport() As String
    On Error GoTo errHandler
    ProductCodeForExport = FNS(mudtProps.ProductCodeForExport)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.ProductCodeForExport"
End Property
Public Property Get EAN() As String
    On Error GoTo errHandler
    EAN = mudtProps.EAN
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.EAN"
End Property
Public Property Let EAN(val As String)
    On Error GoTo errHandler
    mudtProps.EAN = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.EAN"
End Property

Public Property Let ProductCodeF(val As String)
    On Error GoTo errHandler
    mudtProps.ProductCodeF = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.ProductCodeF(val)", val
End Property
Public Property Get ProductCode() As String
    On Error GoTo errHandler
    ProductCode = FNS(mudtProps.ProductCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.ProductCode"
End Property
Public Property Let ProductCode(val As String)
    On Error GoTo errHandler
    mudtProps.ProductCode = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.ProductCode(val)", val
End Property
Public Function SetLineProduct(Optional PID As String, Optional pCode As String, _
                Optional pETA As Integer, Optional IsEditingRow As Boolean = False) As Boolean
    On Error GoTo errHandler
Dim tmp As Double
    Set moProd = New a_Product
    With moProd
        If .Load(PID, 0, FNS(pCode)) <> 99 Then   'product found
            Me.Title = .TitleAuthorPublisher
            If Not IsEditingRow Then
                Me.ProductTypeID = .ProductTypeID
            End If
            Me.PID = .PID
            Me.VATRate = .VATRateToUse
            If PO.CaptureCurrency Is oPC.Configuration.DefaultCurrency Then
 '               If Me.Discount = 0 Then
 '                   Me.LocalPrice = .Cost
 '               Else
                    Me.LocalPrice = .RRP
 '               End If
            ElseIf PO.CaptureCurrency.ID = moProd.ForeignOrderedCURRID Then
                If Not IsEditingRow Then
                    Me.ForeignPrice = moProd.ForeignOrderedPrice
                End If
            ElseIf UCase(Me.PO.CaptureCurrency.SYSNAME) = "GBP" Then
                If Not IsEditingRow Then
                    ConvertToDBL .UKPriceF, tmp
                    Me.ForeignPrice = tmp * PO.CaptureCurrency.Divisor
                End If
            ElseIf UCase(Me.PO.CaptureCurrency.SYSNAME) = "USD" Then
                If Not IsEditingRow Then
                    ConvertToDBL .USPriceF, tmp
                    Me.ForeignPrice = tmp * PO.CaptureCurrency.Divisor
                End If
            ElseIf UCase(Me.PO.CaptureCurrency.SYSNAME) = "EUR" Then
                If Not IsEditingRow Then
                    ConvertToDBL .EUPriceF, tmp
                    Me.ForeignPrice = tmp * PO.CaptureCurrency.Divisor
                End If
            End If
            If Not IsEditingRow Then
                If .Seesafe Then
                    SetQtySS 1
                    SetQtyFirm 0
                Else
                    SetQtyFirm 1
                    SetQtySS 0
                End If
            End If
            If Me.DealID = 0 Then Me.DealID = moProd.DealID
            Me.SetSection .Section
            Me.SetDiscount 0
            Me.ProductCode = .code
            Me.ProductCodeF = .CodeF
            Me.EAN = .EAN
            SetLineProduct = True
        Else
            SetLineProduct = False   'Not found
        End If
    End With
    If pETA > 0 And mPOParent.OrderType <> "NS" Then
        SetETA DateAdd("d", pETA, Date)
    End If
   ValidateObject "TXTCODE"
   ' mobjValid.GetSTatus
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.SetLineProduct(PID,pCODE,pETA)", Array(PID, pCode, pETA)
End Function

'Public Property Get COLsPerPID() As c_COLSPerPID
'    On Error GoTo ErrHandler
'    Set COLsPerPID = mcCOLsPerPID
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_POL.COLsPerPID"
'End Property

'Public Property Get Publisher() As String
'    Publisher = FNS(mudtProps.Publisher)
'End Property
'Public Property Let Publisher(Val As String)
'  If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.Publisher = Val
'    mudtProps.IsDirty = True
'End Property

'Public Property Get Description() As String
'    Description = Trim$(mudtProps.Description)
'End Property
'Public Property Let Description(Val As String)
'  If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.Description = Val
'    mudtProps.IsDirty = True
'End Property
Public Property Get LastAction() As String
    On Error GoTo errHandler
    LastAction = FNS(mudtProps.LastAction)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.LastAction"
End Property
Public Property Let LastAction(val As String)
    mudtProps.LastAction = FNS(val)
End Property
Public Property Get LastActionAndDate() As String
    On Error GoTo errHandler
Dim bDummy As Boolean
    LastActionAndDate = ConvertPOLActionCodes(FNS(mudtProps.LastAction), bDummy) & " " & Format(FNDF(mudtProps.LastActionDate), "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.LastActionAndDate"
End Property

Public Property Get Note() As String
    On Error GoTo errHandler
    Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Note"
End Property
Public Property Let Note(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Note = stripCRLF(val)
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Note(val)", val
End Property

Public Property Get Ref() As String
    On Error GoTo errHandler
    Ref = FNS(mudtProps.Ref)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Ref"
End Property
Public Property Let Ref(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Ref = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Ref(val)", val
End Property
Public Property Get COLID() As Long
    On Error GoTo errHandler
    COLID = FNN(mudtProps.COLID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.COLID"
End Property
Public Property Let COLID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COLID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.COLID(val)", val
End Property
Public Property Get COQty() As Long
    On Error GoTo errHandler
    COQty = FNN(mudtProps.COQty)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.COQty"
End Property
'Public Property Get Code() As String
'    Code = FNS(mudtProps.ProductCode)
'End Property
'Sub TestEvent()
'RaiseEvent Valid("")
'End Sub
'Public Property Let Code(val As String)
'Dim oCode As z_ProdCode
'  If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.ProductCode = val
'    mudtProps.IsDirty = True
'End Property

'Public Property Get CodeF() As String
'    CodeF = FNS(mudtProps.CodeF)
'End Property
'Public Property Let CodeF(val As String)
'  If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.CodeF = val
'    mudtProps.IsDirty = True
'End Property

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
    ErrorIn "a_POL.TitleAuthor"
End Property
Public Property Get Title() As String
    On Error GoTo errHandler
    Title = FNS(mudtProps.Title)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Title"
End Property
Public Property Let Title(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Title = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Title(val)", val
End Property
Public Property Get Publisher() As String
    On Error GoTo errHandler
    Publisher = FNS(mudtProps.Publisher)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Publisher"
End Property
Public Property Let Publisher(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Publisher = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Publisher(val)", val
End Property
Public Property Get PublicationDate() As String
    On Error GoTo errHandler
    PublicationDate = FNS(mudtProps.PublicationDate)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.PublicationDate"
End Property
Public Property Let PublicationDate(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PublicationDate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.PublicationDate(val)", val
End Property
Public Property Get Edition() As String
    On Error GoTo errHandler
    Edition = FNS(mudtProps.Edition)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.PublicationDate"
End Property
Public Property Let Edition(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Edition = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Edition(val)", val
End Property
Public Property Get MainAuthor() As String
    On Error GoTo errHandler
    MainAuthor = FNS(mudtProps.MainAuthor)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.MainAuthor"
End Property
Public Property Let MainAuthor(val As String)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.MainAuthor = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.MainAuthor(val)", val
End Property
Public Property Get DealID() As Long
    On Error GoTo errHandler
    DealID = FNN(mudtProps.DealID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.DEALID"
End Property
Public Property Let DealID(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DealID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.DEALID(val)", val
End Property

Public Property Get Discount() As Double
    On Error GoTo errHandler
    Discount = mudtProps.Discount
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Discount"
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
    RecalculateLine
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Discount(val)", val
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
    RecalculateLine
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.SetDiscount(val)", val
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
    ErrorIn "a_POL.DiscountF"
End Function
Public Property Get VATRate() As Double
    On Error GoTo errHandler
    VATRate = mudtProps.VATRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.VATRate"
End Property
Public Property Let VATRate(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.VATRate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.VATRate(val)", val
End Property
Public Function RecalculateLine()
    On Error GoTo errHandler
'First in local currency
    If Me.PO.ISForeignCurrency Then
        mudtProps.Price = CLng(mudtProps.ForeignPrice / Me.PO.Supplier.DefaultCurrency.Factor)
     '   MsgBox "Code missing"
    End If
    mLPExt = mudtProps.Price * (mudtProps.QtyFirm + mudtProps.QtySS)
    mLPLessDisc = mudtProps.Price * ((100 - mudtProps.Discount) / 100)
    mLPLessDiscVAT = (mLPLessDisc - (mLPLessDiscExt * 100 / (100 + mudtProps.VATRate)))
    mLPLessDiscNONVAT = mLPLessDisc - mLPLessDiscVAT
    mLPLessDiscExt = mLPLessDisc * (mudtProps.QtyFirm + mudtProps.QtySS)
    mLPDiscExt = (mudtProps.Price * (mudtProps.QtyFirm + mudtProps.QtySS)) - mLPLessDiscExt
    mLPLessDiscExtVAT = (mLPLessDiscExt - (mLPLessDiscExt * 100 / (100 + mudtProps.VATRate)))
    mLPLessDiscExtNonVAT = mLPLessDiscExt - mLPLessDiscExtVAT

'Then in Foreign if appropriate
    If mPOParent Is Nothing Then Exit Function
    If Not mPOParent Is Nothing Then
    '    SetForeignPriceAndDeposit
        mFPExt = mudtProps.ForeignPrice * (mudtProps.QtyFirm + mudtProps.QtySS)
        mFPLessDisc = mudtProps.ForeignPrice * ((100 - mudtProps.Discount) / 100)
        mFPLessDiscExt = mFPLessDisc * (mudtProps.QtyFirm + mudtProps.QtySS)
        mFPDiscExt = (mudtProps.ForeignPrice * (mudtProps.QtyFirm + mudtProps.QtySS)) - mFPLessDiscExt
        mFPLessDiscExtVAT = (mFPLessDiscExt - (mFPLessDiscExt * 100 / (100 + mudtProps.VATRate)))
        mFPLessDiscExtNonVAT = mFPLessDiscExt - mFPLessDiscExtVAT
    End If
    RaiseEvent ValueChanges
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'overflow
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.RecalculateLine"
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
    ErrorIn "a_POL.PLessDisc(val)", val
End Property
Public Property Get PLessDiscF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscF = Format(CDbl(mFPLessDisc / mPOParent.CaptureCurrency.Divisor), mPOParent.CaptureCurrency.FormatString)
    Else
        PLessDiscF = Format(CDbl(mLPLessDisc / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.PLessDiscF(val)", val
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
    ErrorIn "a_POL.PExt(val)", val
End Property

Public Property Get PExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PExtF = Format(CDbl(mFPExt / mPOParent.CaptureCurrency.Divisor), mPOParent.CaptureCurrency.FormatString)
    Else
        PExtF = Format(CDbl(mLPExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.PExtF(val)", val
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
    ErrorIn "a_POL.PLessDiscExt(val)", val
End Property

Public Property Get PLessDiscExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscExtF = Format(CDbl(mFPLessDiscExt / mPOParent.CaptureCurrency.Divisor), mPOParent.CaptureCurrency.FormatString)
    Else
        PLessDiscExtF = Format(CDbl(mLPLessDiscExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.PLessDiscExtF(val)", val
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
    ErrorIn "a_POL.PDiscExt(val)", val
End Property
Public Property Get PDiscExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PDiscExtF = Format(CDbl(mFPDiscExt / mPOParent.CaptureCurrency.Divisor), mPOParent.CaptureCurrency.FormatString)
    Else
        PDiscExtF = Format(CDbl(mLPDiscExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.PDiscExtF(val)", val
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
    ErrorIn "a_POL.PLessDiscExtVAT(val)", val
End Property
Public Property Get PLessDiscExtVATF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscExtVATF = Format(CDbl(mFPLessDiscExtVAT / mPOParent.CaptureCurrency.Divisor), mPOParent.CaptureCurrency.FormatString)
    Else
        PLessDiscExtVATF = Format(CDbl(mLPLessDiscExtVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.PLessDiscExtVATF(val)", val
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
    ErrorIn "a_POL.PLessDiscExtNonVAT(val)", val
End Property
Public Property Get PLessDiscExtNonVATF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscExtNonVATF = Format(CDbl(mFPLessDiscExtNonVAT / mPOParent.CaptureCurrency.Divisor), mPOParent.CaptureCurrency.FormatString)
    Else
        PLessDiscExtNonVATF = Format(CDbl(mLPLessDiscExtNonVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.PLessDiscExtNonVATF(val)", val
End Property

Public Function SetNote(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Note) Then
        Err.Raise 384
    End If
    mudtProps.Note = val
    mudtProps.IsDirty = True
    SetNote = True
    Exit Function
End Function
'------------------------------
'
Public Property Get PriceF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PriceF = Format(CDbl(mudtProps.ForeignPrice / mPOParent.CaptureCurrency.Divisor), mPOParent.CaptureCurrency.FormatString)
    Else
        PriceF = Format(CDbl(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.PriceF(val)", val
End Property

Public Property Get Price(bForeign As Boolean)
    On Error GoTo errHandler
    If bForeign Then
        Price = mudtProps.ForeignPrice
    Else
        Price = mudtProps.Price
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Price(bForeign)", bForeign
End Property
Public Function SetPriceLong(val As Long, bForeign As Boolean)
    On Error GoTo errHandler
    If bForeign Then
        mudtProps.ForeignPrice = val
    Else
        mudtProps.Price = val
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.Price(Val,bForeign)", Array(val, bForeign)
End Function

Public Property Get PriceExVAT()
    PriceExVAT = mLPLessDiscNONVAT
End Property
Public Property Get ForeignPrice() As Long
    ForeignPrice = FNN(mudtProps.ForeignPrice)
End Property
Public Property Let ForeignPrice(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ForeignPrice = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.ForeignPrice(val)", val
End Property

Public Property Let LocalPrice(val As Long)
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Price = val
    mudtProps.IsDirty = True
    ValidateObject "PRICE"
    RecalculateLine
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.LocalPrice(val)", val
End Property
Public Function SetPrice(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    If Not PO.CaptureCurrency Is oPC.Configuration.DefaultCurrency Then  'The PO is being captured in a foreign currency and both prices nust be saved
        lngTemp = mudtProps.ForeignPrice
        bOK = SetField_strAsCurrencyToLong(mudtProps.ForeignPrice, val, mcolStack.Count, "PRICE", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
        mudtProps.Price = mudtProps.ForeignPrice * PO.CaptureCurrency.Factor
        If bOK And (lngTemp <> mudtProps.ForeignPrice) Then 'only validate if conversion is OK and the new value is other than the original
            bOK = ValidateObject("PRICE")
        End If
    Else
        lngTemp = mudtProps.Price
        bOK = SetField_strAsCurrencyToLong(mudtProps.Price, val, mcolStack.Count, "PRICE", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
        If bOK And (lngTemp <> mudtProps.Price) Then 'only validate if conversion is OK and the new value is other than the original
            bOK = ValidateObject("PRICE")
        End If
    End If
    SetPrice = bOK
    RecalculateLine
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.SetPrice(val)", val
End Function
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.SetQtyFirm(val)", val
End Function
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
    RecalculateLine
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.SetQtySS(val)", val
End Function
Public Sub SetReplacementForLineID(val As Long)
    
    mudtProps.Replacementfor = val

End Sub
Public Sub SetDateReplaced(val As Date)
    mudtProps.DateReplaced = val
End Sub
Public Property Get DateReplaced() As Date
    DateReplaced = FND(mudtProps.DateReplaced)
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
'H:
'    MsgBox "A_POL: TranslateErrors, " & or
  '  Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    colClassors.Add "Invalid price", "PRICE"
    colClassors.Add "Invalid firm quantity", "QTYFIRM"
    colClassors.Add "Invalid seesafe quantity", "QtySS"
    colClassors.Add "Invalid ETA date (missing or not future date)", "ETA"
    colClassors.Add "Invalid discount", "DISCOUNT"
    colClassors.Add "Invalid product ID", "PID"
    colClassors.Add "Invalid section codes", "SECTION"
    colClassors.Add "Total qty must be > 0", "QTY"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.LoadClassorsCollection"
End Sub
Public Sub ForceValidation()
    ValidateObject ""
End Sub
Private Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean
    bValid = True
    If mudtProps.Price > 0 Or oPC.AllowZeropricedPOLines Then
        mobjValid.RuleBroken "PRICE", False
    Else
        If pFld = "PRICE" Then bValid = False
        mobjValid.RuleBroken "PRICE", True
    End If
    If mudtProps.ETA >= Date Then
        mobjValid.RuleBroken "ETA", False
    Else
        If pFld = "ETA" Then bValid = False
        mobjValid.RuleBroken "ETA", True
    End If
    If mudtProps.QtySS >= 0 Then
        mobjValid.RuleBroken "QtySS", False
    Else
        If pFld = "QtySS" Then bValid = False
        mobjValid.RuleBroken "QtySS", True
    End If
    If mudtProps.QtyFirm >= 0 Then
        mobjValid.RuleBroken "QTYFIRM", False
    Else
        If pFld = "QTYFIRM" Then bValid = False
        mobjValid.RuleBroken "QTYFIRM", True
    End If
    If mudtProps.QtyFirm + mudtProps.QtySS <= 0 Then
        mobjValid.RuleBroken "QTY", True
    Else
        mobjValid.RuleBroken "QTY", False
    End If
    If (mudtProps.Discount < 0 Or mudtProps.Discount > 100) Then
        If pFld = "DISCOUNT" Then bValid = False
        mobjValid.RuleBroken "DISCOUNT", True
    Else
        mobjValid.RuleBroken "DISCOUNT", False
    End If
    If FNS(mudtProps.Section) = "" And oPC.Configuration.EnforceSections Then
        If pFld = "SECTION" Then bValid = False
        mobjValid.RuleBroken "SECTION", True
    Else
        mobjValid.RuleBroken "SECTION", False
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
    ErrorIn "a_POL.ValidateObject(pFld)", pFld
End Function

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.SetDirty(pVal)", pVal
End Sub

'Private Function SetField_CURRENCY(fld As Currency, val As String, pValidationName As String)
'Dim cTemp As Currency
'Dim bTemp As Boolean
'    If mcolStack.Count = 0 Then Err.Raise 383
'    SetField_CURRENCY = True
'    If Trim$(val) = "" Then
'        cTemp = 0
'    ElseIf Not ConvertToCurr(val, cTemp) Then
'        SetField_CURRENCY = False
'        Exit Function
'    End If
'    fld = cTemp
'End Function
'Private Function SetField_strAsCurrencyToLong(fld As Long, val As String, pValidationName As String)
'Dim cTemp As Currency
'Dim bTemp As Boolean
'    If mcolStack.Count = 0 Then Err.Raise 383
'    SetField_strAsCurrencyToLong = True
'    If Trim$(val) = "" Then
'        cTemp = 0
'    ElseIf Not ConvertToCurr(val, cTemp) Then
'        SetField_strAsCurrencyToLong = False
'        Exit Function
'    End If
'    fld = cTemp * oPC.Configuration.DefaultCurrency.Divisor
'End Function
'Private Function SetField_LONG(fld As Long, val As String, pValidationName As String)
'Dim lngTemp As Long
'Dim bTemp As Boolean
'
'    If mcolStack.Count = 0 Then Err.Raise 383
'    SetField_LONG = True
'    If Trim$(val) = "" Then
'        lngTemp = 0
'    ElseIf Not ConvertToLng(val, lngTemp) Then
'        SetField_LONG = False
'        Exit Function
'    End If
'    fld = lngTemp
'End Function
'Private Function SetField_DATE(fld As Date, val As String, pValidationName As String)
'Dim dteTemp As Date
'Dim bTemp As Boolean
'
'    If mcolStack.Count = 0 Then Err.Raise 383
'    SetField_DATE = True
'    If Trim$(val) = "" Then
'        dteTemp = 0
'    ElseIf Not ConvertToDate(val, dteTemp) Then
'        SetField_DATE = False
'        Exit Function
'    End If
'    fld = dteTemp
'End Function
'Private Function SetField_STRING(fld As String, val As String, pValidationName As String)
'Dim strTemp As String
'
'    If mcolStack.Count = 0 Then Err.Raise 383
'    SetField_STRING = True
'    strTemp = val
'    If Len(strTemp) > Len(fld) Then
'        Err.Raise vbObjectError + 1001, "String value too long"
'    End If
'    fld = strTemp
'End Function
'Private Function SetField_DOUBLE(fld As Double, val As String, pValidationName As String)
'Dim dblTEMP As Double
'Dim bTemp As Boolean
'
'    If mcolStack.Count = 0 Then Err.Raise 383
'    SetField_DOUBLE = True
'    If Trim$(val) = "" Then
'        dblTEMP = 0
'    ElseIf Not ConvertToDBL(val, dblTEMP) Then
'        SetField_DOUBLE = False
'        Exit Function
'    End If
'    fld = dblTEMP
'End Function
Public Function CancelLine()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.MarkPOLStatus POLID, "CAN"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.CancelLine"
End Function
Public Function SetETA(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim dteTemp As Date
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
    ErrorIn "a_POL.SetETA(val)", val
End Function
Public Property Get ETACode() As String
    On Error GoTo errHandler
    ETACode = mudtProps.ETACode
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.ETACode"
End Property
Public Property Get ETA() As Date
    On Error GoTo errHandler
    ETA = mudtProps.ETA
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.ETA"
End Property
Public Property Get ETAF() As String
    On Error GoTo errHandler
    ETAF = Format(mudtProps.ETA, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.ETAF"
End Property
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
'    ErrorIn "a_POL.SetSectionALL(pVal)", pVal
'End Function
'Public Property Get Section() As String
'    On Error GoTo ErrHandler
'    Section = FNS(mudtProps.Section)
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_POL.Section"
'End Property
'Public Function SetSection(pVal As String) As Boolean
'    On Error GoTo ErrHandler
'    SetSection = oPC.Configuration.Sections.ValidateNewSetMember(mudtProps.Section, pVal)
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_POL.SetSection(pVal)", pVal
'End Function
Public Property Get Section() As String
    Section = FNS(mudtProps.Section)
End Property
Public Function SetSection(pVal As String) As Boolean
    mudtProps.Section = pVal
End Function

Public Function MarkLineasFulfilled()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.MarkPOLStatus POLID, "FUL"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL.MarkLineasFulfilled"
End Function

