VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Store"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mudtProps As StoreProps
Private mcolStack As Collection
Private mobjParent As ch_Store
Private colClassors As Collection
Event Valid(pMsg As String)
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Dim mKey As String
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property
Public Property Get IsActive() As Boolean
    On Error GoTo errHandler
    IsActive = mudtProps.IsActive
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.IsActive"
End Property
Public Function SetActive(val As Boolean)
      mudtProps.IsActive = val
      SetDirty True
End Function
Public Property Get IsExternal() As Boolean
    On Error GoTo errHandler
    IsExternal = mudtProps.IsExternal
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff.IsExternal"
End Property
Public Function SetExternal(val As Boolean)
      mudtProps.IsExternal = val
      SetDirty True
End Function
Public Sub BeginEdit()
    On Error GoTo errHandler
  ' save object state
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.BeginEdit"
End Sub
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_store.IsEditing"
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
    ErrorIn "a_Store.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddStore Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.ApplyEdit"
End Sub
Friend Sub Initialize(Parent As ch_Store)
    On Error GoTo errHandler
  If mudtProps.IsNew Then Set mobjParent = Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.Initialize(Parent)", Parent
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set mcolStack = New Collection
  Set mobjValid = New z_BrokenRules
  mudtProps.IsNew = True
  mudtProps.IsExternal = True
  mudtProps.IsActive = True
  mudtProps.SystemName = "PAPYRUS"
    LoadClassorsCollection
' if we know any rules are broken on startup
  ' then add code here to initialize the list
  ' of broken rules
  '
   mobjValid.RuleBroken "DESCRIPTION", True
   mobjValid.RuleBroken "ADDRESS", True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
  If mobjParent Is Nothing Then
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
  End If
    
  Set mcolStack = Nothing
  Set colClassors = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.Class_Terminate"
End Sub

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.IsValid"
End Property

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.mobjValid_Status(pMsg)", pMsg
End Sub

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.IsDeleted"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.IsDirty"
End Property

Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As StoreData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As StoreData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  SetState buffer
  ValidateObject ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.SetSuperState(buffer)", buffer
End Sub


Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.Delete"
End Sub

Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = mudtProps.STID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.ID"
End Property
Public Property Let ID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.STID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.ID(val)", val
End Property



Public Property Get code() As String
    On Error GoTo errHandler
    code = FNS(mudtProps.code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.Code"
End Property
Public Function SetCode(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.code) Then
        Err.Raise 384
    End If
    mudtProps.code = val
    SetCode = ValidateObject("Code")
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get SystemName() As String
    On Error GoTo errHandler
    SystemName = FNS(mudtProps.SystemName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.SystemName"
End Property
Public Function SetSystemName(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.SystemName) Then
        Err.Raise 384
    End If
    mudtProps.SystemName = val
    SetSystemName = ValidateObject("SystemName")
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get DescriptionandDelAddress() As String
    On Error GoTo errHandler
    DescriptionandDelAddress = FNS(mudtProps.Description) & vbCrLf & FNS(mudtProps.DelAddress)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.DescriptionandDelAddress"
End Property


Public Property Get Description() As String
    On Error GoTo errHandler
    Description = FNS(mudtProps.Description)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.Description"
End Property
Public Function SetDescription(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Description) Then
        Err.Raise 384
    End If
    mudtProps.Description = val
    SetDescription = ValidateObject("DESCRIPTION")
    mudtProps.IsDirty = True
    Exit Function
End Function

Public Property Get VPNAddress() As String
    On Error GoTo errHandler
    VPNAddress = FNS(mudtProps.VPNAddress)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.VPNAddress"
End Property
Public Function SetVPNAddress(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.VPNAddress) Then
        Err.Raise 384
    End If
    mudtProps.VPNAddress = val
    SetVPNAddress = ValidateObject("VPNAddress")
    mudtProps.IsDirty = True
    Exit Function
End Function


Public Property Get BillAddress() As String
    On Error GoTo errHandler
    BillAddress = FNS(mudtProps.BillAddress)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.BillAddress"
End Property
Public Function SetBillAddress(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.BillAddress) Then
        Err.Raise 384
    End If
    mudtProps.BillAddress = val
    SetBillAddress = ValidateObject("ADDRESS")
    mudtProps.IsDirty = True
    Exit Function
End Function
Public Property Get DelAddress() As String
    On Error GoTo errHandler
    DelAddress = FNS(mudtProps.DelAddress)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.DelAddress"
End Property
Public Function SetDelAddress(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.DelAddress) Then
        Err.Raise 384
    End If
    mudtProps.DelAddress = val
    SetDelAddress = ValidateObject("ADDRESS")
    mudtProps.IsDirty = True
    Exit Function
End Function

Private Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean
    bValid = True
     If Not (Len(Trim(mudtProps.Description)) > 2 Or Len(Trim(mudtProps.Description)) = 0) Then
        mobjValid.BreakRule "DESCRIPTION", True
        If pFld = "DESCRIPTION" Then bValid = False
    Else
        mobjValid.BreakRule "DESCRIPTION", False
    End If
     If Not (Len(Trim(mudtProps.SystemName)) > 2 Or Len(Trim(mudtProps.SystemName)) = 0) Then
        mobjValid.BreakRule "SYSTEMNAME", True
        If pFld = "SYSTEMNAME" Then bValid = False
    Else
        mobjValid.BreakRule "SYSTEMNAME", False
    End If
    If Len(Trim(mudtProps.code)) > 3 Or Not (IsAlphaCaps(Trim(mudtProps.code))) Then
        mobjValid.BreakRule "CODE", True
        If pFld = "CODE" Then bValid = False
    Else
        mobjValid.BreakRule "CODE", False
    End If
    If Len(FNS(mudtProps.BillAddress)) < 5 Then
        mobjValid.BreakRule "ADDRESS", True
        If pFld = "ADDRESS" Then bValid = False
    Else
        mobjValid.BreakRule "ADDRESS", False
    End If
    If Len(FNS(mudtProps.DelAddress)) < 5 Then
        mobjValid.BreakRule "ADDRESS", True
        If pFld = "ADDRESS" Then bValid = False
    Else
        mobjValid.BreakRule "ADDRESS", False
    End If
    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.ValidateObject(pFld)", pFld
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
    ErrorIn "a_Store.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Store description too short", "DESCRIPTION"
    colClassors.Add "System name too short", "SYSTEMNAME"
    colClassors.Add "Address missing", "ADDRESS"
    colClassors.Add "Invalid code", "CODE"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.LoadClassorsCollection"
End Sub
Public Sub GetStatus()
    On Error GoTo errHandler
    ValidateObject ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.GetSTatus"
End Sub

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal Then   ' can only set to true, not to false, because other properties may have been changed
        mudtProps.IsDirty = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Store.SetDirty(pVal)", pVal
End Sub

Public Function StoreIndexClashes() As Boolean
Dim oSM As New a_store_P
    If Me.IsNew Then
        StoreIndexClashes = oSM.StoreIndexClashes(Me.code, Me.ID)
    Else
        StoreIndexClashes = oSM.StoreIndexClashes(Me.code, Me.ID)
    End If

End Function

