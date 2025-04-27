VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Staff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As StaffProps
Private mcolStack As Collection
Private mobjParent As ch_Staff
Private colClassors As Collection
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Event SupervisorstatusChange()
Event RepstatusChange()
Event Valid(pMsg As String)
Private mKey As String
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property

Public Sub BeginEdit()
    On Error GoTo errHandler
  ' save object state
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.BeginEdit"
End Sub
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.IsEditing"
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
    ErrorIn "a_Staff.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddStaff Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.ApplyEdit"
End Sub

Friend Sub Initialize(Parent As ch_Staff)
    On Error GoTo errHandler
  If mudtProps.IsNew Then Set mobjParent = Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.Initialize(Parent)", Parent
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
    ErrorIn "a_Staff.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    If mobjParent Is Nothing Then
      If mcolStack.Count > 0 Then _
        Err.Raise vbObjectError + 1001, , "State stack is not empty"
    End If
    
    Set mcolStack = Nothing
    Set mobjValid = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.Class_Terminate"
End Sub

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.IsValid"
End Property

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.mobjValid_Status(pMsg)", pMsg
End Sub



Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.Delete"
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.IsDeleted"
End Property
Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.IsDirty"
End Property

Private Function GetState() As String
    On Error GoTo errHandler
Dim udtData As StaffData
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As StaffData
        udtData.buffer = buffer
    LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.GetSuperState"
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.SetSuperState(buffer)", buffer
End Sub
Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = mudtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.ID"
End Property
Public Property Let ID(val As Long)
    On Error GoTo errHandler
    mudtProps.ID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.ID(val)", val
End Property
Public Property Get Password() As String
    On Error GoTo errHandler
        Password = FNS(mudtProps.Password)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.Password"
End Property
Public Function SetPassword(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetPassword = True
    If Len(val) > Len(mudtProps.Password) Then
        Err.Raise 384
    End If
    mudtProps.Password = val
    SetPassword = True
    mudtProps.IsDirty = True
    Exit Function
End Function
'Public Property Get Level() As Long
'    On Error GoTo ErrHandler
'    Level = mudtProps.Level
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Staff.Level"
'End Property
'Public Property Let Level(val As Long)
'    On Error GoTo ErrHandler
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.Level = val
'    mudtProps.IsDirty = True
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Staff.Level(val)", val
'End Property
Public Property Get IsActive() As Boolean
    On Error GoTo errHandler
    IsActive = GetRole(enSECURITY_ACTIVE)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.IsActive"
End Property
'Public Property Let Active(val As Boolean)
'    On Error GoTo ErrHandler
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.Active = val
'    mudtProps.IsDirty = True
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Staff.Active(val)", val
'End Property
'
Public Property Get IsRep() As Boolean
    On Error GoTo errHandler
    IsRep = GetRole(enSECURITY_ISREP)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.IsRep"
End Property
'Public Property Let IsRep(val As Boolean)
'    On Error GoTo ErrHandler
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.IsRep = val
'    mudtProps.IsDirty = True
'    RaiseEvent RepstatusChange
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Staff.IsRep(val)", val
'End Property
'
Public Property Get IsOperator() As Boolean
    On Error GoTo errHandler
    IsOperator = GetRole(enSECURITY_ISOPERATOR)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.IsOperator"
End Property
'Public Property Let IsOperator(val As Boolean)
'    On Error GoTo ErrHandler
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.IsOperator = val
'    mudtProps.IsDirty = True
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Staff.IsOperator(val)", val
'End Property
'
Public Property Get IsSupervisor() As Boolean
    On Error GoTo errHandler
    IsSupervisor = GetRole(enSECURITY_ISSUPERVISOR)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.IsSupervisor"
End Property
'Public Property Let IsSupervisor(val As Boolean)
'    On Error GoTo ErrHandler
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.IsSupervisor = val
'    mudtProps.IsDirty = True
'    RaiseEvent SupervisorstatusChange
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Staff.IsSupervisor(val)", val
'End Property




Public Property Get Shortname() As String
    On Error GoTo errHandler
    Shortname = FNS(mudtProps.Shortname)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.Shortname"
End Property
Public Function SetShortname(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetShortname = True
    If Len(val) > Len(mudtProps.Shortname) Then
        Err.Raise 384
    End If
    mudtProps.Shortname = val
    SetShortname = True
    mudtProps.IsDirty = True

    Exit Function
End Function

Public Property Get StaffName() As String
    On Error GoTo errHandler
    StaffName = FNS(mudtProps.StaffName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.StaffName"
End Property
Public Function SetStaffName(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetStaffName = True
    If Len(val) > Len(mudtProps.StaffName) Then
        Err.Raise 384
    End If
    mudtProps.StaffName = val
    SetStaffName = True
    mudtProps.IsDirty = True

    Exit Function
End Function

Public Property Get Signature() As String
    On Error GoTo errHandler
    Signature = FNS(mudtProps.Signature)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.Signature"
End Property
Public Function SetSignature(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetSignature = True
    If Len(val) > Len(mudtProps.Signature) Then
        Err.Raise 384
    End If
    mudtProps.Signature = val
    SetSignature = True
    mudtProps.IsDirty = True

    Exit Function
End Function

Public Property Get EMail() As String
    On Error GoTo errHandler
    EMail = FNS(mudtProps.EMail)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.Email"
End Property
Public Function SetEmail(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetEmail = True
    If Len(val) > Len(mudtProps.EMail) Then
        Err.Raise 384
    End If
    mudtProps.EMail = val
    SetEmail = True
    mudtProps.IsDirty = True

    Exit Function
End Function



Public Property Get StaffTel() As String
    On Error GoTo errHandler
    StaffTel = FNS(mudtProps.StaffTel)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.StaffTel"
End Property
Public Function SetRole(iPos As enumSecurityNode, val As Boolean)
     Mid(mudtProps.Role, iPos, 1) = IIf(val, "Y", "N")
End Function
Public Function GetRole(iPos As enumSecurityNode) As Boolean
    GetRole = IIf(Mid(mudtProps.Role, iPos, 1) = "Y", True, False)
End Function
Public Function SetStaffTel(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetStaffTel = True
    If Len(val) > Len(mudtProps.StaffTel) Then
        Err.Raise 384
    End If
    mudtProps.StaffTel = val
    SetStaffTel = True
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get StaffCell() As String
    On Error GoTo errHandler
    StaffCell = FNS(mudtProps.StaffCell)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.StaffCell"
End Property
Public Function SetStaffCell(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    val = Trim(val)
    SetStaffCell = True
    If Len(val) > Len(mudtProps.StaffCell) Then
        Err.Raise 384
    End If
    mudtProps.StaffCell = val
    SetStaffCell = True
    mudtProps.IsDirty = True
    Exit Function
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
    ErrorIn "a_Staff.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Missing name", "NAME"
    colClassors.Add "Invalid password", "PASSWORD"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.LoadClassorsCollection"
End Sub


Public Property Get SQLSTatus() As String
    On Error GoTo errHandler
    SQLSTatus = FNS(mudtProps.SQLSTatus)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.SQLSTatus"
End Property
Public Property Get SQLMsg() As String
    On Error GoTo errHandler
    SQLMsg = FNS(mudtProps.SQLMsg)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.SQLMsg"
End Property

