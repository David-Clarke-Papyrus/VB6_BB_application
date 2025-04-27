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
Attribute VB_Exposed = True
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
10        On Error GoTo errHandler
        ' save object state
20      mcolStack.Add GetState
30        Exit Sub
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.BeginEdit"
End Sub
Public Property Get Key() As String
10        Key = mKey
End Property
Public Property Let Key(val As String)
10        mKey = val
End Property
Public Property Get IsEditing() As Boolean
10        On Error GoTo errHandler
20      IsEditing = (mcolStack.Count > 0)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.IsEditing"
End Property

Public Sub CancelEdit()
10        On Error GoTo errHandler
20      If mcolStack.Count = 0 Then Err.Raise 445
        
30      mudtProps.IsDeleted = False
        ' restore object state
40      With mcolStack
50        SetState .Item(.Count)
60        .Remove .Count
70      End With
80        Exit Sub
errHandler:
90        If ErrMustStop Then Debug.Assert False: Resume
100       ErrorIn "a_Currency.CancelEdit"
End Sub

Public Sub ApplyEdit()
10        On Error GoTo errHandler
20      If mcolStack.Count = 0 Then Err.Raise 445

30      If Not mobjParent Is Nothing Then
40        mobjParent.AddCurrency Me
50        Set mobjParent = Nothing
60      End If
70      mcolStack.Remove mcolStack.Count
80        Exit Sub
errHandler:
90        If ErrMustStop Then Debug.Assert False: Resume
100       ErrorIn "a_Currency.ApplyEdit"
End Sub
Friend Sub Initialize(Parent As ch_Currencies)
10        On Error GoTo errHandler
20      If mudtProps.IsNew Then Set mobjParent = Parent
30        Exit Sub
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.Initialize(Parent)", Parent
End Sub

Private Sub Class_Initialize()
10        On Error GoTo errHandler
20      Set mcolStack = New Collection
30      Set mobjValid = New z_BrokenRules
40      mudtProps.IsNew = True
50        LoadClassorsCollection
        ' if we know any rules are broken on startup
        ' then add code here to initialize the list
        ' of broken rules
        '
        ' mobjValid.RuleBroken "RuleName", True
60        Exit Sub
errHandler:
70        If ErrMustStop Then Debug.Assert False: Resume
80        ErrorIn "a_Currency.Class_Initialize"
End Sub

Private Sub Class_Terminate()
10        On Error GoTo errHandler
20      If mobjParent Is Nothing Then
30        If mcolStack.Count > 0 Then _
            Err.Raise vbObjectError + 1001, , "State stack is not empty"
40      End If
          
50      Set mcolStack = Nothing
60        Exit Sub
errHandler:
70        If ErrMustStop Then Debug.Assert False: Resume
80        ErrorIn "a_Currency.Class_Terminate"
End Sub

Public Property Get IsValid() As Boolean
10        On Error GoTo errHandler
20      IsValid = (mobjValid.Count = 0)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.IsValid"
End Property

Private Sub mobjValid_BrokenRule(ors As String)
10        On Error GoTo errHandler
20        RaiseEvent Valid(Translateors(ors))
30        Exit Sub
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.mobjValid_BrokenRule(ors)", ors
End Sub
Private Sub mobjValid_RuleUnbroken(ors As String)
10        On Error GoTo errHandler
20        RaiseEvent Valid(Translateors(ors))
30        Exit Sub
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.mobjValid_RuleUnbroken(ors)", ors
End Sub

Private Sub mobjValid_NoBrokenRules()
10        On Error GoTo errHandler
20        RaiseEvent Valid("")
30        Exit Sub
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
10        On Error GoTo errHandler
20        RaiseEvent Valid(Translateors(pMsg))
30        Exit Sub
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.mobjValid_Status(pMsg)", pMsg
End Sub

Public Property Get IsDeleted() As Boolean
10        On Error GoTo errHandler
20      IsDeleted = mudtProps.IsDeleted
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.IsDeleted"
End Property

Public Property Get IsNew() As Boolean
10        On Error GoTo errHandler
20      IsNew = mudtProps.IsNew
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.IsNew"
End Property

Public Property Get IsDirty() As Boolean
10        On Error GoTo errHandler
20      IsDirty = mudtProps.IsDirty
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.IsDirty"
End Property

Private Function GetState() As String
10        On Error GoTo errHandler
        Dim udtData As CurrencyData
        
20      LSet udtData = mudtProps
30      GetState = udtData.buffer
40        Exit Function
errHandler:
50        If ErrMustStop Then Debug.Assert False: Resume
60        ErrorIn "a_Currency.GetState"
End Function

Private Sub SetState(buffer As String)
10        On Error GoTo errHandler
        Dim udtData As CurrencyData
        
20      udtData.buffer = buffer
30      LSet mudtProps = udtData
40        Exit Sub
errHandler:
50        If ErrMustStop Then Debug.Assert False: Resume
60        ErrorIn "a_Currency.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
10        On Error GoTo errHandler
20      GetSuperState = GetState
30        Exit Function
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
10        On Error GoTo errHandler
20      SetState buffer
30        Exit Sub
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.SetSuperState(buffer)", buffer
End Sub

Public Sub Load(CURRID As Long)
10        On Error GoTo errHandler
        
20      If mcolStack.Count > 0 Then Err.Raise 445
30      If Not mudtProps.IsNew Then Err.Raise 445
        
40      mudtProps.IsNew = False
50        Exit Sub
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "a_Currency.Load(CURRID)", CURRID
End Sub

Public Sub Delete()
10        On Error GoTo errHandler
20      If mcolStack.Count = 0 Then Err.Raise 445
        
30      mudtProps.IsDeleted = True
40      mudtProps.IsDirty = True
50        Exit Sub
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "a_Currency.Delete"
End Sub

Public Property Get ID() As Long
10        On Error GoTo errHandler
20        ID = mudtProps.ID
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.ID"
End Property
Public Property Let ID(val As Long)
10        On Error GoTo errHandler
20        If mcolStack.Count = 0 Then Err.Raise 383
30        mudtProps.ID = val
40        mudtProps.IsDirty = True
50        Exit Property
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "a_Currency.ID(val)", val
End Property
Public Property Get SYSNAME() As String
10        SYSNAME = FNS(mudtProps.SYS)
End Property
Public Property Let SYSNAME(val As String)
10        mudtProps.SYS = FNS(val)
End Property
Public Property Get Description() As String
10        On Error GoTo errHandler
20        Description = FNS(mudtProps.Description)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.Description"
End Property
Public Property Let Description(val As String)
10        On Error GoTo errHandler
20        If mcolStack.Count = 0 Then Err.Raise 383
30        mudtProps.Description = val
40        mudtProps.IsDirty = True
50        Exit Property
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "a_Currency.Description(val)", val
End Property
Public Property Get Symbol() As String
10        On Error GoTo errHandler
20        Symbol = FNS(mudtProps.Symbol)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.Symbol"
End Property
Public Property Let Symbol(val As String)
10        On Error GoTo errHandler
20        If mcolStack.Count = 0 Then Err.Raise 383
30        mudtProps.Symbol = val
40        mudtProps.IsDirty = True
50        Exit Property
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "a_Currency.Symbol(val)", val
End Property
Public Property Get FormatString() As String
10        On Error GoTo errHandler
20        FormatString = FNS(mudtProps.Format)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.FormatString"
End Property
Public Property Let FormatString(val As String)
10        On Error GoTo errHandler
20        If mcolStack.Count = 0 Then Err.Raise 383
30        mudtProps.Format = val
40        mudtProps.IsDirty = True
50        Exit Property
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "a_Currency.FormatString(val)", val
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
10        On Error GoTo errHandler
20        Factor = FNDBL(mudtProps.Factor)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.Factor"
End Property
Public Property Let Factor(val As Double)
10        On Error GoTo errHandler
20        If mcolStack.Count = 0 Then Err.Raise 383
30        mudtProps.Factor = val
40        mudtProps.IsDirty = True
50        Exit Property
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "a_Currency.Factor(val)", val
End Property
Public Property Get FactorINV() As Double
10        On Error GoTo errHandler
20        FactorINV = CStr(Round(1# / mudtProps.Factor, 6))
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.FactorINV"
End Property
Public Property Get FactorF() As String
10        On Error GoTo errHandler
20        If mudtProps.Factor > 1 Then
30            FactorF = Format(Round(mudtProps.Factor, 2), "###0.##")
40        Else
50            FactorF = Format(Round(mudtProps.Factor, 4), "0.####")
60        End If
70        Exit Property
errHandler:
80        If ErrMustStop Then Debug.Assert False: Resume
90        ErrorIn "a_Currency.FactorFormatted"
End Property
Public Property Get FactorINVF() As String
10        On Error GoTo errHandler
20        If FactorINV > 1 Then
30            FactorINVF = Format(Round(FactorINV, 2), "###0.##")
40        Else
50            FactorINVF = Format(Round(FactorINV, 4), "##0.####")
60        End If
70        Exit Property
errHandler:
80        If ErrMustStop Then Debug.Assert False: Resume
90        ErrorIn "a_Currency.FactorFormatted"
End Property
    'txtFactor2 = Round(CStr(1# / CDbl(txtFactor)), 6)

Public Function SetFactor(val As String) As Boolean
10        On Error GoTo errHandler
      Dim dblTmp As Double
20        SetFactor = True
30        If Not ConvertToDBL(val, dblTmp) Then
40            SetFactor = False
50            Exit Function
60        End If
70        mudtProps.Factor = dblTmp
80        mudtProps.IsDirty = True

90        Exit Function
errHandler:
100       If ErrMustStop Then Debug.Assert False: Resume
110       ErrorIn "a_Currency.SetFactor(val)", val
End Function
Public Property Get Divisor() As Integer
10        On Error GoTo errHandler
20        Divisor = FNN(mudtProps.Divisor)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.Divisor"
End Property
Public Property Let Divisor(val As Integer)
10        On Error GoTo errHandler
20        If mcolStack.Count = 0 Then Err.Raise 383
30        mudtProps.Divisor = val
40        mudtProps.IsDirty = True
50        Exit Property
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "a_Currency.Divisor(val)", val
End Property
Public Property Get DivisorF() As String
10        On Error GoTo errHandler
20        DivisorF = Format(mudtProps.Divisor, "##0")
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_Currency.DivisorF"
End Property
Public Function SetDivisor(val As String) As Boolean
10        On Error GoTo errHandler
      Dim iTmp As Integer
20        SetDivisor = True
30        If Not ConvertToInt(val, iTmp) Then
40            SetDivisor = False
50            Exit Function
60        End If
70        mudtProps.Divisor = iTmp
80        mudtProps.IsDirty = True

90        Exit Function
errHandler:
100       If ErrMustStop Then Debug.Assert False: Resume
110       ErrorIn "a_Currency.SetDivisor(val)", val
End Function
Private Function Translateors(pRawors As String) As String
10        On Error GoTo errHandler
      Dim strOut As String
      Dim strRule, strAllRules As String
      Dim NoMoreRules As Boolean
      Dim iMarker, iStart As Integer
20        iMarker = 1
30        strAllRules = ""
40        If Len(pRawors) > 0 Then
50            iMarker = InStr(iMarker + 1, pRawors, ",")
60            If iMarker > 0 Then
70                strAllRules = colClassors(Left(pRawors, iMarker - 1))
80            Else
90                strAllRules = colClassors(pRawors)
100           End If
110           Do Until iMarker = 0
120               iStart = iMarker + 1
130               iMarker = InStr(iStart, pRawors, ",")
140               If iMarker > 0 Then
150                   strRule = colClassors(Mid(pRawors, iStart, iMarker - iStart))
160               Else
170                   strRule = colClassors(Mid(pRawors, iStart))
180               End If
                      
190               strAllRules = strAllRules & vbCrLf & strRule
200           Loop
210       End If
220       Translateors = strAllRules
230       Exit Function
errHandler:
240       If ErrMustStop Then Debug.Assert False: Resume
250       ErrorIn "a_Currency.Translateors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
10        On Error GoTo errHandler
20        Set colClassors = New Collection
30        colClassors.Add "Missing customer", "TP"
40        colClassors.Add "Note too short", "NOTE"
50        Exit Sub
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "a_Currency.LoadClassorsCollection"
End Sub

