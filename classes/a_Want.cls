VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Want"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event Valid(Status As Boolean)
Private mudtProps As WantProps
Private mcolStack As Collection
Private mobjParent As ch_Want
Private colClassors As Collection
Dim mKey As String
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1


Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property
Private Function GetState() As String
  Dim mudtData As WantData
  
  LSet mudtData = mudtProps
  GetState = mudtData.buffer
End Function

Private Sub SetState(buffer As String)
  Dim mudtData As WantData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Function GetSuperState() As String
  GetSuperState = GetState
End Function
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub

Public Sub ApplyEdit()
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddWant Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
End Sub

Public Sub BeginEdit()
  mcolStack.Add GetState
End Sub
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Want.IsEditing"
End Property

Public Sub CancelEdit()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
End Sub

Private Sub Class_Initialize()
  Set mcolStack = New Collection
  Set mobjValid = New z_BrokenRules
    mudtProps.RequestDate = Date
    LoadClassorsCollection
    mobjValid.RuleBroken "DTE", True
    mobjValid.RuleBroken "TP", True
    mudtProps.IsNew = True
End Sub
Friend Sub Initialize(Parent As ch_Want)
  If mudtProps.IsNew Then Set mobjParent = Parent
End Sub

Friend Property Get IsValid() As Boolean
  IsValid = (mobjValid.Count = 0)
End Property

Private Sub Class_Terminate()
  If mobjParent Is Nothing Then
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
  End If
  Set mcolStack = Nothing
End Sub

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

Public Sub Delete()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
End Sub

Public Property Get IsDeleted() As Boolean
  IsDeleted = mudtProps.IsDeleted
End Property

Public Property Get IsNew() As Boolean
  IsNew = mudtProps.IsNew
End Property

Public Property Get IsDirty() As Boolean
  IsDirty = mudtProps.IsDirty
End Property


Public Property Get ID() As Long
    ID = mudtProps.ID
End Property
Public Property Get CustomerName() As String
    CustomerName = FNS(mudtProps.CustomerName)
End Property
Public Property Let CustomerName(val As String)
    mudtProps.CustomerName = val
End Property
Public Property Let PID(val As Long)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.PID = val
    mudtProps.IsDirty = True
End Property
Public Property Get PID() As Long
  PID = mudtProps.PID
End Property
Public Property Get WantID() As Long
    WantID = FNN(mudtProps.ID)
End Property
Public Property Let WantID(val As Long)
  If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ID = val
    mudtProps.IsDirty = True
End Property

Public Property Get TPID() As Long
    TPID = mudtProps.TPID
End Property
Public Function SetTPID(val As Long) As Boolean
Dim bOK As Boolean
    If val > 0 Then
        mudtProps.TPID = val
        bOK = True
    Else
        bOK = False
    End If
    If bOK Then
        bOK = ValidateObject("TP")
    End If
    SetTPID = bOK
End Function
Public Property Get ReqDate() As Date
    ReqDate = mudtProps.RequestDate
End Property
Public Property Let ReqDate(val As Date)
    mudtProps.RequestDate = val
End Property
Public Property Get ReqDateF() As String
    ReqDateF = Format(mudtProps.RequestDate, "dd/mm/yyyy")
End Property
Public Function SetRequestDate(val As String) As Boolean
Dim bOK As Boolean
    bOK = SetField_DATE(mudtProps.RequestDate, val, "DTE")
    If bOK Then
        bOK = ValidateObject("DTE")
    End If
    SetRequestDate = bOK
End Function

Public Property Get Note() As String
    Note = FNS(mudtProps.Notes)
End Property
Public Property Let Note(val As String)
    mudtProps.Notes = val
End Property

Private Function ValidateObject(pFld As String) As Boolean
On Error GoTo H
Dim bValid As Boolean
    bValid = True
    If mudtProps.RequestDate > Date Or mudtProps.RequestDate < DateAdd("yyyy", -1, Date) Then
        If pFld = "DTE" Then bValid = False
        mobjValid.BreakRule "DTE", True
    Else
        mobjValid.BreakRule "DTE", False
    End If
    If mudtProps.TPID > 0 Then
        mobjValid.BreakRule "TP", False
    Else
        If pFld = "TP" Then bValid = False
        mobjValid.BreakRule "TP", True
    End If
    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
EXIT_Handler:
    Exit Function
H:
    MsgBox "CheckOK: " & Error
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
    colClassors.Add "Missing note", "NOTE"
    colClassors.Add "Missing requestor", "TP"
    colClassors.Add "Invalid date", "DTE"
End Sub
Private Function SetField_DATE(fld As Date, val As String, pValidationName As String)
Dim dteTemp As Date
Dim bTemp As Boolean

    If mcolStack.Count = 0 Then Err.Raise 383
    SetField_DATE = True
    If Trim$(val) = "" Then
        dteTemp = 0
    ElseIf Not ConvertToDate(val, dteTemp) Then
        SetField_DATE = False
        Exit Function
    End If
    fld = dteTemp
End Function
Private Sub SetDirty(pVal As Boolean)
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
End Sub





