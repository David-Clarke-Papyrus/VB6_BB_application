VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_APPRL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event Valid(msg As String)
Private mudtProps As APPRLProps
Private mcolStack As Collection
Private mobjParent As ch_APPRL
Private mAPPParent As a_APPR
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private mKey As String
Private colClassors As Collection
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
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property
Private Function GetState() As String
    On Error GoTo errHandler
  Dim mudtData As APPRLData
  
  LSet mudtData = mudtProps
  GetState = mudtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim mudtData As APPRLData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
  'mudtProps.DiscountedPrice = (mudtProps.Price * mudtProps.Qty) * (1 - mudtProps.Discount)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.SetSuperState(buffer)", buffer
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddApproReturnLine Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.ApplyEdit"
End Sub

Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.BeginEdit"
End Sub
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.IsEditing"
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
    ErrorIn "a_APPRL.CancelEdit"
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set mcolStack = New Collection
  Set mobjValid = New z_BrokenRules
  mudtProps.IsNew = True
  LoadClassorsCollection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.Class_Initialize"
End Sub
Friend Sub Initialize(Parent As ch_APPRL)
    On Error GoTo errHandler
  If mudtProps.IsNew Then Set mobjParent = Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.Initialize(Parent)", Parent
End Sub

Friend Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.IsValid"
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
    ErrorIn "a_APPRL.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.mobjValid_BrokenRule(ors)", oRS
End Sub

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.Delete"
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.IsDeleted"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.IsDirty"
End Property
    
  Public Property Get AppLineID() As Long
    On Error GoTo errHandler
    AppLineID = mudtProps.APPLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.AppLineID"
  End Property
  
  Public Property Let AppLineID(Value As Long)
    On Error GoTo errHandler
    mudtProps.APPLID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.AppLineID(Value)", Value
  End Property
  
  Public Property Get TRID() As Long
    On Error GoTo errHandler
    TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.TRID"
  End Property
  
  Public Property Let TRID(Value As Long)
    On Error GoTo errHandler
    mudtProps.TRID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.TRID(Value)", Value
  End Property
  
  Public Property Get Qty() As Long
    On Error GoTo errHandler
    Qty = mudtProps.Qty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.Qty"
  End Property
  Public Property Get QtyIssued() As Long
    On Error GoTo errHandler
    QtyIssued = FNN(mudtProps.QtyIssued)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.QtyIssued"
  End Property
    Public Property Let QtyIssued(val As Long)
    On Error GoTo errHandler
    mudtProps.QtyIssued = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.QtyIssued"
  End Property

  Public Property Get QtyReturned() As Long
    On Error GoTo errHandler
    QtyReturned = FNN(mudtProps.QtyReturned)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.QtyReturned"
  End Property
  Public Property Let QtyReturned(val As Long)
    On Error GoTo errHandler
    mudtProps.QtyReturned = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.QtyReturned"
  End Property
  
  Public Property Let Qty(Value As Long)
    On Error GoTo errHandler
    mudtProps.Qty = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.Qty(Value)", Value
  End Property
  
  Public Property Get PID() As String
    On Error GoTo errHandler
    PID = FNS(mudtProps.PID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.PID"
  End Property
  
  Public Property Let PID(Value As String)
    On Error GoTo errHandler
    mudtProps.PID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.PID(Value)", Value
  End Property
  
  Public Property Get APPLID() As Long
    On Error GoTo errHandler
    APPLID = mudtProps.APPLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.APPLID"
  End Property

Public Property Let APPLID(Value As Long)
    On Error GoTo errHandler
    mudtProps.APPLID = Value
    If (Value = 0) Then
        mobjValid.RuleBroken "APPLID", True
    Else
        mobjValid.RuleBroken "APPLID", False
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.APPLID(Value)", Value
End Property

Public Property Get Title() As String
    On Error GoTo errHandler
    Title = stripCRLF(FNS(mudtProps.Title))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.Title"
End Property
Public Property Get Fulfilled() As String
    On Error GoTo errHandler
    Fulfilled = FNS(mudtProps.Fulfilled)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.Fulfilled"
End Property

Public Property Let Title(Value As String)
    On Error GoTo errHandler
    mudtProps.Title = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.Title(Value)", Value
End Property
Public Property Get EAN() As String
    On Error GoTo errHandler
    EAN = FNS(mudtProps.EAN)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.EAN"
End Property
Public Property Let EAN(val As String)
    On Error GoTo errHandler
    mudtProps.EAN = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.EAN"
End Property
Public Property Let CodeF(val As String)
    On Error GoTo errHandler
    mudtProps.CodeF = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.CodeF"
End Property

Public Property Get CodeF() As String
    On Error GoTo errHandler
    CodeF = FNS(mudtProps.CodeF)
    If CodeF = "" Then
        CodeF = FNS(mudtProps.code)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.CodeF"
End Property
Public Property Get code() As String
    On Error GoTo errHandler
    code = FNS(mudtProps.code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.Code"
End Property

Public Property Let code(Value As String)
    On Error GoTo errHandler
    mudtProps.code = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.Code(Value)", Value
End Property
Public Property Get Note() As String
    On Error GoTo errHandler
    Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.Note"
End Property

Public Property Let Note(Value As String)
    On Error GoTo errHandler
    mudtProps.Note = stripCRLF(Value)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.Note(Value)", Value
End Property

Public Property Get APPRLID() As Long
    On Error GoTo errHandler
    APPRLID = mudtProps.APPRLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.APPRLID"
End Property

Public Property Let APPRLID(Value As Long)
    On Error GoTo errHandler
    mudtProps.APPRLID = Value
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.APPRLID(Value)", Value
End Property
Public Property Get ApproCode() As String
    On Error GoTo errHandler
    ApproCode = FNS(mudtProps.ApproCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.ApproCode"
End Property
Public Property Let ApproCode(val As String)
    On Error GoTo errHandler
    mudtProps.ApproCode = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.ApproCode(val)", val
End Property
Public Property Get ApproDate() As Date
    On Error GoTo errHandler
    ApproDate = mudtProps.ApproDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.ApproDate"
End Property
Public Property Get ApproDateF() As Date
    On Error GoTo errHandler
    ApproDateF = Format(mudtProps.ApproDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.ApproDateF"
End Property

Public Function SetQty(val As String) As Boolean
    On Error GoTo errHandler
Dim lngQty As Long
  If mcolStack.Count = 0 Then Err.Raise 383
    SetQty = True
    If Trim$(val) = "" Then
        lngQty = 0
    ElseIf Not ConvertToLng(val, lngQty) Then
        SetQty = False
        Exit Function
    End If
    If (lngQty < 1 Or lngQty > (Me.QtyIssued - Me.QtyReturned)) Then
        mobjValid.RuleBroken "Qty", True
    Else
        mobjValid.RuleBroken "Qty", False
    End If
    mudtProps.Qty = lngQty
    mudtProps.IsDirty = True
    ValidateObject "QTY"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.SetQty(val)", val
End Function
Friend Sub SetParentAPPR(pParentAPPR As a_APPR)
    On Error GoTo errHandler
    Set mAPPParent = pParentAPPR
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.SetParentAPPR(pParentAPPR)", pParentAPPR
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
    ErrorIn "a_APPRL.GetSTatus"
End Function
Private Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean
    bValid = True
    If mudtProps.Qty > 0 And mudtProps.Qty <= (Me.QtyIssued - Me.QtyReturned) Then
        mobjValid.RuleBroken "QTY", False
    Else
        If pFld = "QTY" Then bValid = False
        mobjValid.RuleBroken "QTY", True
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
    ErrorIn "a_APPRL.ValidateObject(pFld)", pFld
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
    ErrorIn "a_APPRL.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Invalid quantity", "QTY"
    colClassors.Add "No appro match", "APPLID"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.LoadClassorsCollection"
End Sub

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.SetDirty(pVal)", pVal
End Sub


Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL.mobjValid_Status(pMsg)", pMsg
End Sub

Public Function CancelLine()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.CancelAPPRL Me.APPRLID
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL.CancelLine"
End Function

