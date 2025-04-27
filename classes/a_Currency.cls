VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Currency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
Dim mudtProps As CurrencyProps
Private mcolStack As Collection
Private mobjParent As ch_Currencies
Event Valid(IsValid As Boolean)
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassors As Collection
Private mKey As String

Public Sub BeginEdit()
    On Error GoTo errHandler
  ' save object state
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.BeginEdit"
End Sub
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.IsEditing"
End Property

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = False
  ' restore object state
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddCurrency Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.ApplyEdit"
End Sub
Friend Sub Initialize(Parent As ch_Currencies)
    On Error GoTo errHandler
  If mudtProps.IsNew Then Set mobjParent = Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.Initialize(Parent)", Parent
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set mcolStack = New Collection
  Set mobjValid = New z_BrokenRules
  mudtProps.IsNew = True
    LoadClassorsCollection
  ' if we know any rules are broken on startup
  ' then add code here to initialize the list
  ' of broken rules
  '
  ' mobjValid.RuleBroken "RuleName", True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.Class_Initialize"
End Sub

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
    ErrorIn "a_Currency.Class_Terminate"
End Sub

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.IsValid"
End Property

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.mobjValid_Status(pMsg)", pMsg
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.IsDeleted"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.IsDirty"
End Property

Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As CurrencyData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As CurrencyData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.SetSuperState(buffer)", buffer
End Sub

Public Sub Load(CURRID As Long)
    On Error GoTo errHandler
  
  If mcolStack.Count > 0 Then Err.Raise 445
  If Not mudtProps.IsNew Then Err.Raise 445
  
  mudtProps.IsNew = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.Load(CURRID)", CURRID
End Sub

Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.Delete"
End Sub

Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = mudtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.ID"
End Property
Public Property Let ID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.ID(val)", val
End Property
Public Property Get SYSNAME() As String
    SYSNAME = FNS(mudtProps.SYS)
End Property
Public Property Let SYSNAME(val As String)
    mudtProps.SYS = FNS(val)
End Property
Public Property Get Description() As String
    On Error GoTo errHandler
    Description = FNS(mudtProps.Description)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.Description"
End Property
Public Property Let Description(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Description = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.Description(val)", val
End Property
Public Property Get Symbol() As String
    On Error GoTo errHandler
    Symbol = FNS(mudtProps.Symbol)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.Symbol"
End Property
Public Property Let Symbol(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Symbol = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.Symbol(val)", val
End Property
Public Property Get FormatString() As String
    On Error GoTo errHandler
    FormatString = FNS(mudtProps.Format)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.FormatString"
End Property
Public Property Let FormatString(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Format = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.FormatString(val)", val
End Property
'Public Property Get Divisor() As Double
'    Divisor = Trim(mudtProps.Divisor)
'End Property
'Public Property Let Divisor(val As Double)
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.Divisor = val
'    mudtProps.IsDirty = True
'End Property
Public Property Get Factor() As Double
    On Error GoTo errHandler
    Factor = FNDBL(mudtProps.Factor)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.Factor"
End Property
Public Property Let Factor(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Factor = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.Factor(val)", val
End Property
Public Property Get FactorInv() As Double
    On Error GoTo errHandler
    FactorInv = CStr(Round(1# / mudtProps.Factor, 6))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.FactorINV"
End Property
Public Property Get FactorF() As String
    On Error GoTo errHandler
    If mudtProps.Factor > 1 Then
        FactorF = Format(Round(mudtProps.Factor, 2), "###0.##")
    Else
        FactorF = Format(Round(mudtProps.Factor, 4), "0.####")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.FactorFormatted"
End Property
Public Property Get FactorINVF() As String
    On Error GoTo errHandler
    If FactorInv > 1 Then
        FactorINVF = Format(Round(FactorInv, 2), "###0.##")
    Else
        FactorINVF = Format(Round(FactorInv, 4), "##0.####")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.FactorFormatted"
End Property
    'txtFactor2 = Round(CStr(1# / CDbl(txtFactor)), 6)

Public Function SetFactor(val As String) As Boolean
    On Error GoTo errHandler
Dim dblTmp As Double
    SetFactor = True
    If Not ConvertToDBL(val, dblTmp) Then
        SetFactor = False
        Exit Function
    End If
    mudtProps.Factor = dblTmp
    mudtProps.IsDirty = True

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.SetFactor(val)", val
End Function
Public Property Get Divisor() As Integer
    On Error GoTo errHandler
    Divisor = FNN(mudtProps.Divisor)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.Divisor"
End Property
Public Property Let Divisor(val As Integer)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Divisor = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.Divisor(val)", val
End Property
Public Property Get DivisorF() As String
    On Error GoTo errHandler
    DivisorF = Format(mudtProps.Divisor, "##0")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.DivisorF"
End Property
Public Function SetDivisor(val As String) As Boolean
    On Error GoTo errHandler
Dim iTmp As Integer
    SetDivisor = True
    If Not ConvertToInt(val, iTmp) Then
        SetDivisor = False
        Exit Function
    End If
    mudtProps.Divisor = iTmp
    mudtProps.IsDirty = True

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.SetDivisor(val)", val
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
    ErrorIn "a_Currency.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Missing customer", "TP"
    colClassors.Add "Note too short", "NOTE"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency.LoadClassorsCollection"
End Sub

