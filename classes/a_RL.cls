VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_RL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event Valid(msg As String)
Event ValueChanges()
Private moProd As a_Product
Private mudtProps As RLProps
Private mcolStack As Collection
Private mobjParent As ch_RL
Private colClassors As Collection
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private mKey As String
Private mRParent As a_R
Private mFPExt As Long
Private mLPExt As Long
Dim bAllowNoMatchToDelivery As Boolean
Private mLPLessDisc As Long
Private mFPLessDisc As Long
Private mLPDiscExt As Long
Private mFPDiscExt As Long
Private mLPLessDiscExtNonVAT As Long
Private mLPLessDiscExtVAT As Long
Private mLPDisc_Unit As Long
Private mLPDisc_Ext As Long
Private mFPLessDiscExt As Long
Private mLPLessDiscExt As Long
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
    End Select
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Properties(pSrt)", pSrt
End Property

Public Property Get Sequence() As Long
    On Error GoTo errHandler
    Sequence = mudtProps.Sequence
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Sequence"
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
    ErrorIn "a_RL.Sequence(val)", val
End Property

Public Function RecalculateLine()
Dim lngTmp As Long

    On Error GoTo errHandler
    If mRParent.Status = stCOMPLETE Then
        lngTmp = mudtProps.QtyReturned - mudtProps.QtyRejected
    Else
        lngTmp = mudtProps.QtyRequested
    End If
    mLPExt = mudtProps.Price * (lngTmp)
    mFPExt = mudtProps.ForeignPrice * (lngTmp)
    mLPLessDisc = mudtProps.Price * ((100 - Me.Discount) / 100)
    mFPLessDisc = mudtProps.ForeignPrice * ((100 - Me.Discount) / 100)
    
    mLPLessDiscExt = mLPLessDisc * (lngTmp)
    mFPLessDiscExt = mFPLessDisc * (lngTmp)
    mLPDiscExt = (mudtProps.Price * (lngTmp)) - mLPLessDiscExt
    mFPDiscExt = (mudtProps.ForeignPrice * (lngTmp)) - mFPLessDiscExt
'    mLPLessDiscExtVAT = (mLPLessDiscExt - (mLPLessDiscExt * 100 / (100 + mudtProps.VATRate)))
'    mLPLessDiscExtNonVAT = mLPLessDiscExt - mLPLessDiscExtVAT
    RaiseEvent ValueChanges
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.RecalculateLine"
End Function
Public Property Get PLessDisc() As Long
    On Error GoTo errHandler
        PLessDisc = mLPLessDisc
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.PLessDisc"
End Property
Public Property Get PLessDiscF() As String
    On Error GoTo errHandler
        PLessDiscF = Format(CDbl(mLPLessDisc / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.PLessDiscF"
End Property
'Public Property Get PLessDiscExtNonVAT() As Long
'    PLessDiscExtNonVAT = mLPLessDiscExtNonVAT
'End Property
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
    ErrorIn "a_RL.PExt"
End Property

Public Property Get PExtF() As String
    On Error GoTo errHandler
        PExtF = Format(CDbl(mLPExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.PExtF"
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
    ErrorIn "a_RL.PLessDiscExt"
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
    ErrorIn "a_RL.PDiscExt"
End Property
Public Property Get PDiscExtF() As String
    On Error GoTo errHandler
        PDiscExtF = Format(CDbl(mLPDiscExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.PDiscExtF"
End Property

'Public Property Get PLessDiscExtVAT() As Long
'        PLessDiscExtVAT = mLPLessDiscExtVAT
'End Property
'Public Property Get PLessDiscExtVATF() As String
'        PLessDiscExtVATF = Format(CDbl(mLPLessDiscExtVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'
'Public Property Get PLessDiscExtNonVATF() As String
'        PLessDiscExtNonVATF = Format(CDbl(mLPLessDiscExtNonVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
'End Property

'------------------------------

Public Property Get Key() As String
    On Error GoTo errHandler
    Key = mKey
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Key"
End Property
Public Property Let Key(val As String)
    mKey = val
End Property
Friend Sub SetParentR(pParentR As a_R)
    On Error GoTo errHandler
    Set mRParent = pParentR
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.SetParentR(pParentR)", pParentR
End Sub

Private Function GetState() As String
    On Error GoTo errHandler
  Dim mudtData As RLData
  LSet mudtData = mudtProps
  GetState = mudtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim mudtData As RLData
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.SetSuperState(buffer)", buffer
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
    ErrorIn "a_RL.GetSTatus"
End Function


Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddRLine Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.ApplyEdit"
End Sub
Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.BeginEdit"
End Sub
Public Function CancelLine()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.CancelRL Me.ID, Me.PID, CLng(Me.QtyReturned)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.CancelLine"
End Function
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.IsEditing"
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
    ErrorIn "a_RL.CancelEdit"
End Sub
Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    LoadClassorsCollection
    mobjValid.RuleBroken "QTY", True
 '   mobjValid.RuleBroken "DELLID", True
    mudtProps.IsNew = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Class_Initialize"
End Sub
Friend Sub Initialize(Parent As ch_RL)
    On Error GoTo errHandler
  If mudtProps.IsNew Then Set mobjParent = Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Initialize(Parent)", Parent
End Sub

Friend Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.IsValid"
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
    ErrorIn "a_RL.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.mobjValid_BrokenRule(ors)", oRS
End Sub

'Private Sub mobjValid_NoBrokenRules()
'    On Error GoTo ErrHandler
'    RaiseEvent Valid("")
'    Exit Sub
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_RL.mobjValid_NoBrokenRules"
'End Sub

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Delete"
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.IsDeleted"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.IsDirty"
End Property

Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = mudtProps.RLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.ID"
End Property
Public Property Get QtyRequested() As Long
    On Error GoTo errHandler
    QtyRequested = mudtProps.QtyRequested
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.QtyRequested"
End Property
Public Property Let QtyRequested(val As Long)
    On Error GoTo errHandler
    mudtProps.QtyRequested = val
    mobjValid.RuleBroken "QTY", Not val > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.QtyRequested(val)", val
End Property
Public Function SetQtyRequested(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.QtyRequested
    bOK = SetField_LONG(mudtProps.QtyRequested, val, "QTY", mcolStack.Count)
    If bOK And mudtProps.QtyRequested <> lngTemp Then   'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QTY")
    End If
    SetQtyRequested = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.SetQtyRequested(val)", val
End Function

Public Property Get QtyReturned() As Long
    On Error GoTo errHandler
    QtyReturned = mudtProps.QtyReturned
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.QtyReturned"
End Property
Public Property Let QtyReturned(val As Long)
    On Error GoTo errHandler
    mudtProps.QtyReturned = val
    mobjValid.RuleBroken "QTY", Not val > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.QtyReturned(val)", val
End Property
Public Function SetQtyReturned(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.QtyReturned
    bOK = SetField_LONG(mudtProps.QtyReturned, val, "QTY", mcolStack.Count)
    If bOK And mudtProps.QtyReturned <> lngTemp Then   'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QTY")
    End If
    SetQtyReturned = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.SetQtyReturned(val)", val
End Function

Public Property Get QtyApproved() As Long
    On Error GoTo errHandler
    QtyApproved = mudtProps.QtyApproved
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.QtyApproved"
End Property
Public Property Let QtyApproved(val As Long)
    On Error GoTo errHandler
    mudtProps.QtyApproved = val
    mobjValid.RuleBroken "QTY", Not val > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.QtyApproved(val)", val
End Property
Public Function SetQtyApproved(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.QtyApproved
    bOK = SetField_LONG(mudtProps.QtyApproved, val, "QTY", mcolStack.Count)
    If bOK And mudtProps.QtyApproved <> lngTemp Then   'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QTY")
    End If
    SetQtyApproved = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.SetQtyApproved(val)", val
End Function

Public Property Get QtyRejected() As Long
    On Error GoTo errHandler
    QtyRejected = mudtProps.QtyRejected
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.QtyRejected"
End Property
Public Property Let QtyRejected(val As Long)
    On Error GoTo errHandler
    mudtProps.QtyRejected = val
    mobjValid.RuleBroken "QTY", Not val > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.QtyRejected(val)", val
End Property
Public Function SetQtyRejected(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim lngTemp As Long
    lngTemp = mudtProps.QtyRejected
    bOK = SetField_LONG(mudtProps.QtyRejected, val, "QTY", mcolStack.Count)
    If bOK And mudtProps.QtyRejected <> lngTemp Then   'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("QTY")
    End If
    SetQtyRejected = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.SetQtyRejected(val)", val
End Function


Public Function SetPrice(val As String) As Boolean
    On Error GoTo errHandler
'Dim bOK As Boolean
'Dim lngTemp As Long
'    lngTemp = mudtProps.Price
'    bOK = SetField_LONG(mudtProps.Price, val, "PRICE", mcolStack.Count)
'    If bOK And mudtProps.Price <> lngTemp Then   'only validate if conversion is OK and the new value is other than the original
'        bOK = ValidateObject("PRICE")
'    End If
'    SetPrice = bOK
'
'
'
    '''''''''''''''''''''''''''''''''''''''''
Dim bOK As Boolean
Dim lngTemp As Long
Dim tmpPrice As Long
Dim lngRoundTo As Long

    SetPrice = False
    If Not mRParent.Supplier.DefaultCurrency Is oPC.Configuration.DefaultCurrency Then
        lngTemp = mudtProps.ForeignPrice
        bOK = SetField_strAsCurrencyToLong(mudtProps.ForeignPrice, val, mcolStack.Count, "PRICE", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
        tmpPrice = mudtProps.ForeignPrice / mRParent.Supplier.DefaultCurrency.Factor
        lngRoundTo = oPC.Configuration.RoundingRules.GetRoundTo(tmpPrice)
        mudtProps.Price = (RoundUp(tmpPrice, lngRoundTo))
        
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
    If bOK Then
        SetPrice = bOK
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.SetPrice(val)", val
End Function

Public Property Get Title() As String
    On Error GoTo errHandler
    Title = FNS(mudtProps.Title)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Title"
End Property
Public Property Let Title(val As String)
    On Error GoTo errHandler
    mudtProps.Title = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Title(val)", val
End Property
Public Property Get Pubcode() As String
    On Error GoTo errHandler
    Pubcode = FNS(mudtProps.Pubcode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Pubcode"
End Property
Public Property Get Sections() As String
    On Error GoTo errHandler
    Sections = FNS(mudtProps.Section)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Sections"
End Property

Public Property Let CodeF(val As String)
    On Error GoTo errHandler
    mudtProps.CodeF = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.CodeF(val)", val
End Property
Public Property Let curPrice(val As Long)
    On Error GoTo errHandler
    mudtProps.Price = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.curPrice(val)", val
End Property
Public Property Get PID() As String
    On Error GoTo errHandler
    PID = FNS(mudtProps.PID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.PID"
End Property
Public Property Let PID(val As String)
    On Error GoTo errHandler
    mudtProps.PID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.PID(val)", val
End Property
Public Property Get SINVRef() As String
    On Error GoTo errHandler
    SINVRef = FNS(mudtProps.SINVRef)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.SINVRef"
End Property
Public Property Let SINVRef(val As String)
    On Error GoTo errHandler
    mudtProps.SINVRef = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.SINVRef(val)", val
End Property
Public Property Get DocRef() As String
    On Error GoTo errHandler
    DocRef = FNS(mudtProps.DocRef)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.DocRef"
End Property
Public Property Let DocRef(val As String)
    On Error GoTo errHandler
    mudtProps.DocRef = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.DocRef(val)", val
End Property
Public Property Get DOCDate() As Date
    On Error GoTo errHandler
    DOCDate = FND(mudtProps.DOCDate)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.DocDate"
End Property
Public Property Let DOCDate(val As Date)
    On Error GoTo errHandler
    mudtProps.DOCDate = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.DocDate(val)", val
End Property

Public Property Get PriceF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PriceF = Format(CDbl(mudtProps.ForeignPrice / mRParent.CaptureCurrency.Divisor), mRParent.CaptureCurrency.FormatString) & " " & mRParent.CaptureCurrency.SYSNAME
    Else
        PriceF = Format(CDbl(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.PriceF(val)", val
End Property
Public Property Get PLessDiscExtF(val As Boolean) As String
    On Error GoTo errHandler
    If val Then
        PLessDiscExtF = Format(CDbl(mFPLessDiscExt / mRParent.CaptureCurrency.Divisor), mRParent.CaptureCurrency.FormatString) & " " & mRParent.CaptureCurrency.SYSNAME
    Else
        PLessDiscExtF = Format(CDbl(mLPLessDiscExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.PLessDiscExtF(val)", val
End Property

'Public Property Get EAN() As String
'    EAN = Trim$(mudtProps.EAN)
'End Property
'Public Property Let EAN(val As String)
'    If Len(val) > Len(mudtProps.EAN) Then Err.Raise vbObjectError + 1001, "String value too long"
'    mudtProps.EAN = val
'    SetDirty True
'End Property
Public Property Get Status() As String
    On Error GoTo errHandler
    Status = FNS(mudtProps.Status)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Status"
End Property
Public Property Let Status(val As String)
    On Error GoTo errHandler
    mudtProps.Status = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Status(val)", val
End Property
Public Property Get EAN() As String
    On Error GoTo errHandler
    EAN = FNS(mudtProps.EAN)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.EAN"
End Property
Public Property Let EAN(val As String)
    On Error GoTo errHandler
    mudtProps.EAN = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.EAN(val)", val
End Property

Public Property Get CodeF() As String
    On Error GoTo errHandler
    CodeF = FNS(mudtProps.CodeF)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.CodeF"
End Property
Public Property Get code() As String
    On Error GoTo errHandler
    code = FNS(mudtProps.code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Code"
End Property
Public Property Let code(val As String)
    On Error GoTo errHandler
    mudtProps.code = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Code(val)", val
End Property
Public Property Let Discount(val As Double)
    On Error GoTo errHandler
    mudtProps.Discount = val 'IIf(val = 0, 38, val)  'we use 38% as nominal value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Discount(val)", val
End Property
Public Property Get Discount() As Double
    On Error GoTo errHandler
    Discount = Round(FNDBL(mudtProps.Discount), 2)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Discount"
End Property
Public Property Get DiscountF() As String
    On Error GoTo errHandler
    DiscountF = PBKSPercentF(Round(FNDBL(mudtProps.Discount), 2))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.DiscountF"
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
    ErrorIn "a_RL.Price(bForeign)", bForeign
End Property
Public Property Let LocalPrice(val As Long)
    On Error GoTo errHandler
    mudtProps.Price = val
    mobjValid.RuleBroken "PRICE", Not (val > 0 Or mudtProps.ForeignPrice > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.LocalPrice(val)", val
End Property
Public Property Let ForeignPrice(val As Long)
    On Error GoTo errHandler
    mudtProps.ForeignPrice = val
    mobjValid.RuleBroken "PRICE", Not (val > 0 Or mudtProps.Price > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.ForeignPrice(val)", val
End Property
Public Property Get TRID() As Long
    On Error GoTo errHandler
    TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.TRID"
End Property
Public Property Let TRID(val As Long)
    On Error GoTo errHandler
    mudtProps.TRID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.TRID(val)", val
End Property
Public Property Get DELLID() As Long
    On Error GoTo errHandler
    DELLID = mudtProps.DELLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.DELLID"
End Property
Public Property Let DELLID(val As Long)
    On Error GoTo errHandler
    mudtProps.DELLID = val
  '  mobjValid.RuleBroken "DELLID", Not mudtProps.DELLID > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.DELLID(val)", val
End Property
Public Sub AllowNoMatchingDelivery()
    bAllowNoMatchToDelivery = True
     ValidateObject ""
End Sub
Public Property Get DELID() As Long
    On Error GoTo errHandler
    DELID = mudtProps.DELID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.DELID"
End Property
Public Property Let DELID(val As Long)
    On Error GoTo errHandler
    mudtProps.DELID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.DELID(val)", val
End Property

Public Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler


Dim bOK As Boolean
    bOK = True
    If mudtProps.Price > 0 Then
        mobjValid.RuleBroken "PRICE", False
    Else
        mobjValid.RuleBroken "PRICE", True
    End If
    If mudtProps.QtyRequested >= 0 Then
        mobjValid.RuleBroken "QTY", False
    Else
        mobjValid.RuleBroken "QTY", True
    End If
    If mudtProps.DELLID > 0 Or bAllowNoMatchToDelivery = True Then
        mobjValid.RuleBroken "DELLID", False
        LogSaveToFile ("clear DELLID")
    Else
        mobjValid.RuleBroken "DELLID", True
        LogSaveToFile ("set DELLID")
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
    ErrorIn "a_RL.ValidateObject(pFld)", pFld
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
    ErrorIn "a_RL.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Invalid price", "PRICE"
    colClassors.Add "Invalid quantity", "QTY"
    colClassors.Add "No delivery line selected", "DELLID"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.LoadClassorsCollection"
End Sub


Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.SetDirty(pVal)", pVal
End Sub

Public Function SetLineProduct(Optional PID As String, Optional pCode As String) As Boolean
    On Error GoTo errHandler
    
    Set moProd = New a_Product
    With moProd
        If .Load(PID, 0, FNS(pCode)) <> 99 Then   'product found
            Me.Title = .Title '.TitleAuthorPublisher
            Me.PID = .PID
            Me.code = .code
            Me.CodeF = .CodeF
            Me.QtyRequested = 1
            SetLineProduct = True
        Else
            SetLineProduct = False   'Not found
        End If
    End With
    mobjValid.GetStatus
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.SetLineProduct(PID,pCODE)", Array(PID, pCode)
End Function

Public Property Get Note() As String
    On Error GoTo errHandler
    Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Note"
End Property
Public Property Let Note(val As String)
    On Error GoTo errHandler
'  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Note = stripCRLF(val)
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.Note(val)", val
End Property
Public Property Get MainAuthor() As String
    On Error GoTo errHandler
    MainAuthor = FNS(mudtProps.MainAuthor)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.MainAuthor"
End Property



Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.mobjValid_NoBrokenRules"
End Sub

Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_Status(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL.mobjValid_Status(ors)", oRS
End Sub


