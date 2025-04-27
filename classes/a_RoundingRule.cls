VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_RoundingRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mudtProps As RRProps
Private mcolStack As Collection
Private mobjParent As ch_RoundingRule
Private colClassors As Collection
Event Valid(msg As String)
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private mKey As String

Public Sub BeginEdit()
  ' save object state
  mcolStack.Add GetState
End Sub

Public Sub CancelEdit()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = False
  ' restore object state
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
End Sub

Public Sub ApplyEdit()
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddRR Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
End Sub
Friend Sub Initialize(Parent As ch_RoundingRule)
  If mudtProps.IsNew Then Set mobjParent = Parent
End Sub

Private Sub Class_Initialize()
  Set mcolStack = New Collection
  Set mobjValid = New z_BrokenRules
  mudtProps.IsNew = True
  Set colClassors = New Collection
    LoadClassorsCollection
  ' if we know any rules are broken on startup
  ' then add code here to initialize the list
  ' of broken rules
  '
  ' mobjValid.RuleBroken "RuleName", True
End Sub

Private Sub Class_Terminate()
  If mobjParent Is Nothing Then
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
  End If
    
  Set mcolStack = Nothing
  Set colClassors = Nothing
End Sub

Public Property Get IsValid() As Boolean
  IsValid = (mobjValid.Count = 0)
End Property

Private Sub mobjValid_BrokenRule(oRS As String)
    RaiseEvent Valid(TranslateErrors(oRS))
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    RaiseEvent Valid(TranslateErrors(oRS))
End Sub

Private Sub mobjValid_NoBrokenRules()
    RaiseEvent Valid("")
End Sub
Private Sub mobjValid_Status(pMsg As String)
    RaiseEvent Valid(TranslateErrors(pMsg))
End Sub

Public Property Get IsDeleted() As Boolean
  IsDeleted = mudtProps.IsDeleted
End Property

Public Property Get IsNew() As Boolean
  IsNew = mudtProps.IsNew
End Property
Public Property Get IsEditing() As Boolean
  IsEditing = (mcolStack.Count > 0)
End Property

Public Property Get IsDirty() As Boolean
  IsDirty = mudtProps.IsDirty
End Property

Private Function GetState() As String
  Dim udtData As RRData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
End Function

Private Sub SetState(buffer As String)
  Dim udtData As RRData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
End Sub
Friend Function GetSuperState() As String
  GetSuperState = GetState
End Function

Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property


Public Sub Delete()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
End Sub

Public Property Get ID() As Long
    ID = mudtProps.ID
End Property
Public Property Let ID(val As Long)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ID = val
    mudtProps.IsDirty = True
End Property

Public Property Get LowerBound() As Long
    LowerBound = mudtProps.LowerBound
End Property
Public Property Get LowerBoundF() As String
    LowerBoundF = Format(mudtProps.LowerBound / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Function SetLowerBound(val As String) As Boolean
Dim lngLB As Long
    SetLowerBound = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngLB = 0
    ElseIf (Not ConvertToLng(val, lngLB)) Then
        SetLowerBound = False
        Exit Function
    End If
    mudtProps.LowerBound = lngLB
    SetLowerBound = ValidateObject("LB")
    mudtProps.IsDirty = True
End Function
Public Property Get UpperBound() As Long
    UpperBound = mudtProps.UpperBound
End Property
Public Property Get UpperBoundF() As String
    UpperBoundF = Format(mudtProps.UpperBound / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Function SetUpperBound(val As String) As Boolean
Dim lngUB As Long
    SetUpperBound = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngUB = 0
    ElseIf (Not ConvertToLng(val, lngUB)) Then
        SetUpperBound = False
        Exit Function
    End If
    mudtProps.UpperBound = lngUB
    SetUpperBound = ValidateObject("UB")
    mudtProps.IsDirty = True
End Function
Public Property Get RoundTo() As Long
    RoundTo = mudtProps.RoundTo
End Property
Public Property Get RoundToF() As String
    RoundToF = Format(mudtProps.RoundTo / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Function SetRoundTo(val As String) As Boolean
Dim lngRT As Long
    SetRoundTo = True
    val = Trim(val)
    If Trim$(val) = "" Then
        lngRT = 0
    ElseIf (Not ConvertToLng(val, lngRT)) Then
        SetRoundTo = False
        Exit Function
    End If
    mudtProps.RoundTo = lngRT
    SetRoundTo = ValidateObject("RT")
    mudtProps.IsDirty = True
End Function

Private Function ValidateObject(pFld As String) As Boolean
Dim bValid As Boolean
    bValid = True
    If mudtProps.UpperBound < mudtProps.LowerBound Then
        mobjValid.BreakRule "UB", True
        If pFld = "UB" Then bValid = False
    Else
        mobjValid.BreakRule "UB", False
    End If
    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
End Function
Private Function TranslateErrors(pRawors As String) As String
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
End Function
Private Sub LoadClassorsCollection()
    Set colClassors = New Collection
    colClassors.Add "Invalid lower bound", "LB"
    colClassors.Add "Invalid upper bound", "UB"
    colClassors.Add "Invalid Round to figure", "RT"
End Sub

Private Sub SetDirty(pVal As Boolean)
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
End Sub





