VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_CSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event Valid(msg As String)
Event ValueChanges()

Private mudtProps As CSLProps
Private mcolStack As Collection
Private mobjParent As ch_CSL
Private colClassors As Collection
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private mKey As String
Private mCSParent As a_CS
Private mFPExt As Long
Private mLPExt As Long

Private mLPLessDisc As Long
Private mLPLessDiscExt As Long
Private mLPDiscExt As Long
Private mLPLessDiscExtNonVAT As Long
Private mLPLessDiscExtVAT As Long
Private mLPDisc_Unit As Long
Private mLPDisc_Ext As Long

Public Function RecalculateLine()
    On Error GoTo errHandler
    mLPExt = mudtProps.Price * (mudtProps.Qty)
    mLPLessDisc = mudtProps.Price * ((100 - mudtProps.Discount) / 100)
    mLPLessDiscExt = mLPLessDisc * (mudtProps.Qty)
    mLPDiscExt = (mudtProps.Price * (mudtProps.Qty)) - mLPLessDiscExt
    mLPLessDiscExtVAT = (mLPLessDiscExt - (mLPLessDiscExt * 100 / (100 + mudtProps.VATRate)))
    mLPLessDiscExtNonVAT = mLPLessDiscExt - mLPLessDiscExtVAT
    RaiseEvent ValueChanges
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'overflow
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.RecalculateLine"
End Function
Public Property Get PLessDisc() As Long
        PLessDisc = mLPLessDisc
End Property
Public Property Get PLessDiscF() As String
    On Error GoTo errHandler
        PLessDiscF = Format(CDbl(mLPLessDisc / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.PLessDiscF"
End Property

Public Property Get PExt() As Long
        PExt = mLPExt
End Property

Public Property Get PExtF() As String
    On Error GoTo errHandler
        PExtF = Format(CDbl(mLPExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.PExtF"
End Property

Public Property Get PLessDiscExt() As Long
        PLessDiscExt = mLPLessDiscExt
End Property

Public Property Get PLessDiscExtF() As String
    On Error GoTo errHandler
        PLessDiscExtF = Format(CDbl(mLPLessDiscExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.PLessDiscExtF"
End Property

Public Property Get PDiscExt() As Long
        PDiscExt = mLPDiscExt
End Property
Public Property Get PDiscExtF() As String
    On Error GoTo errHandler
        PDiscExtF = Format(CDbl(mLPDiscExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.PDiscExtF"
End Property

Public Property Get PLessDiscExtVAT() As Long
        PLessDiscExtVAT = mLPLessDiscExtVAT
End Property
Public Property Get PLessDiscExtVATF() As String
    On Error GoTo errHandler
        PLessDiscExtVATF = Format(CDbl(mLPLessDiscExtVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.PLessDiscExtVATF"
End Property

Public Property Get PLessDiscExtNonVAT() As Long
        PLessDiscExtNonVAT = mLPLessDiscExtNonVAT
End Property
Public Property Get PLessDiscExtNonVATF() As String
    On Error GoTo errHandler
        PLessDiscExtNonVATF = Format(CDbl(mLPLessDiscExtNonVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.PLessDiscExtNonVATF"
End Property

'Public Function SetNote(val As String) As Boolean
'    If mcolStack.Count = 0 Then Err.Raise 383
'    If Len(val) > Len(mudtProps.Note) Then
'        Err.Raise 384
'    End If
'    mudtProps.Note = val
'    mudtProps.IsDirty = True
'    SetNote = True
'End Function
'------------------------------
Public Property Get PriceF() As String
    On Error GoTo errHandler
    PriceF = Format(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.PriceF"
End Property



Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property
Friend Sub SetParentCS(pParentCS As a_CS)
    On Error GoTo errHandler
    Set mCSParent = pParentCS
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.SetParentCS(pParentCS)", pParentCS
End Sub

Private Function GetState() As String
    On Error GoTo errHandler
  Dim mudtData As CSLData
  LSet mudtData = mudtProps
  GetState = mudtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim mudtData As CSLData
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.SetSuperState(buffer)", buffer
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
    ErrorIn "a_CSL.GetSTatus"
End Function


Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddCSLine Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.ApplyEdit"
End Sub
Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.BeginEdit"
End Sub
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.IsEditing"
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
    ErrorIn "a_CSL.CancelEdit"
End Sub
Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    LoadClassorsCollection
    mobjValid.RuleBroken "Price", True
    mudtProps.Qty = 1
    mudtProps.IsNew = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.Class_Initialize"
End Sub
Friend Sub Initialize(Parent As ch_CSL)
    On Error GoTo errHandler
  If mudtProps.IsNew Then Set mobjParent = Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.Initialize(Parent)", Parent
End Sub

Friend Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.IsValid"
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
    ErrorIn "a_CSL.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(oRS)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.mobjValid_BrokenRule(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.mobjValid_NoBrokenRules"
End Sub

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.Delete"
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.IsDeleted"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.IsDirty"
End Property





Public Property Let ID(val As Long)
    On Error GoTo errHandler
    mudtProps.ID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.ID(val)", val
End Property


Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = mudtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.ID"
End Property
Public Property Get Qty() As Long
    On Error GoTo errHandler
    Qty = mudtProps.Qty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.Qty"
End Property
Public Property Let Qty(val As Long)
    On Error GoTo errHandler
    mudtProps.Qty = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.Qty(val)", val
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
    ErrorIn "a_CSL.SetQty(val)", val
End Function
Public Property Get Title() As String
    On Error GoTo errHandler
    Title = FNS(mudtProps.Title)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.Title"
End Property
Public Property Let Title(val As String)
    On Error GoTo errHandler
    mudtProps.Title = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.Title(val)", val
End Property
Public Property Let CodeF(val As String)
    On Error GoTo errHandler
    mudtProps.CodeF = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.CodeF(val)", val
End Property
Public Property Let curPrice(val As Long)
    On Error GoTo errHandler
    mudtProps.Price = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.curPrice(val)", val
End Property
Public Property Get PID() As String
    On Error GoTo errHandler
    PID = FNS(mudtProps.PID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.PID"
End Property
Public Property Let PID(val As String)
    On Error GoTo errHandler
    mudtProps.PID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.PID(val)", val
End Property
Public Property Get EAN() As String
    On Error GoTo errHandler
    EAN = Trim$(mudtProps.EAN)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.EAN"
End Property
Public Property Let EAN(val As String)
    On Error GoTo errHandler
    If Len(val) > Len(mudtProps.EAN) Then Err.Raise vbObjectError + 1001, "String value too long"
    mudtProps.EAN = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.EAN(val)", val
End Property
Public Property Get CodeF() As String
    On Error GoTo errHandler
    CodeF = FNS(mudtProps.CodeF)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.CodeF"
End Property
Public Property Get code() As String
    On Error GoTo errHandler
    code = FNS(mudtProps.code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.Code"
End Property
Public Property Get DiscountF() As String
    On Error GoTo errHandler
    If mudtProps.Discount = 0 Then
        DiscountF = ""
    Else
        DiscountF = PBKSPercentF(CDbl(mudtProps.Discount))
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.DiscountF"
End Property
Public Property Get Price() As Long
    On Error GoTo errHandler
    Price = mudtProps.Price
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.Price"
End Property
Public Property Let Price(val As Long)
    On Error GoTo errHandler
    mudtProps.Price = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.Price(val)", val
End Property
Public Property Get TRID() As Long
    On Error GoTo errHandler
    TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.TRID"
End Property
Public Property Let TRID(val As Long)
    On Error GoTo errHandler
    mudtProps.TRID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.TRID(val)", val
End Property
Public Property Let DateTime(val As String)
    On Error GoTo errHandler
    
    mobjValid.RuleBroken "DAteTime", (Not IsDate(val))
    If (IsDate(val)) Or (val = "0") Then mudtProps.DateTime = CVDate(val)
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.DateTime(val)", val
End Property
Public Property Get DateTime() As String
    On Error GoTo errHandler
    If CLng(mudtProps.DateTime) = 0 Then
        DateTime = ""
    Else
        DateTime = Format(mudtProps.DateTime, "ddd   h:Nn::ss A/P")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.DateTime"
End Property
Public Property Get dteDateTime() As Date
    On Error GoTo errHandler
    If HasData(mudtProps.DateTime) Then
        dteDateTime = mudtProps.DateTime
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.dteDateTime"
End Property
Public Property Get DateTimeForSort() As String
    On Error GoTo errHandler
    If CLng(mudtProps.DateTime) = 0 Then
        DateTimeForSort = ""
    Else
        DateTimeForSort = Format(mudtProps.DateTime, "yyyy:mm:dd Hh:Nn:Ss")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.DateTimeForSort"
End Property

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
    ErrorIn "a_CSL.ValidateObject(pFld)", pFld
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
    ErrorIn "a_CSL.TranslateErrors(pRawors)", pRawors
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
    ErrorIn "a_CSL.LoadClassorsCollection"
End Sub


Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL.SetDirty(pVal)", pVal
End Sub

