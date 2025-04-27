VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_ProductSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mudtProps As ProductSectionProps
Private mcolStack As Collection
Private mobjParent As ch_ProductSection
Private colClassors As Collection
Event Valid(pMsg As String)
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Dim mKey As String
Public Property Get Key() As String
    Key = Trim(mKey)
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
    ErrorIn "a_ProductSection.IsEditing"
End Property

Public Sub BeginEdit()
    On Error GoTo errHandler
     mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.BeginEdit"
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
    ErrorIn "a_ProductSection.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445

  If Not mobjParent Is Nothing Then
    mobjParent.AddMember Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.ApplyEdit"
End Sub
Friend Sub Initialize(Parent As ch_ProductSection)
    On Error GoTo errHandler
  If mudtProps.IsNew Then Set mobjParent = Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.Initialize(Parent)", Parent
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set mcolStack = New Collection
  Set mobjValid = New z_BrokenRules
  mudtProps.IsNew = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.Class_Initialize"
End Sub
Public Property Get Description() As String
    On Error GoTo errHandler
    Description = FNS(mudtProps.Description)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.Description"
End Property
Public Property Let Description(val As String)
    mudtProps.Description = FNS(val)
End Property
Public Property Get PID() As String
    On Error GoTo errHandler
    PID = mudtProps.PID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.PID"
End Property
Public Property Get SECID() As Long
    On Error GoTo errHandler
    SECID = mudtProps.SECID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.SECID"
End Property
Public Property Let PID(val As String)
    On Error GoTo errHandler
    mudtProps.PID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.PID(val)", val
End Property
Public Property Let SECID(val As Long)
    On Error GoTo errHandler
    mudtProps.SECID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.SECID(val)", val
End Property

Public Property Get Priority() As Long
    Priority = FNN(mudtProps.Priority)
End Property
Public Property Get PriorityF() As String
    PriorityF = IIf(FNN(mudtProps.Priority) = 99, "***", "")
End Property
Public Property Let Priority(val As Long)
    
    mudtProps.Priority = FNN(val)
    SetDirty True
    
End Property

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
    ErrorIn "a_ProductSection.Class_Terminate"
End Sub

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
  IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.IsValid"
End Property

Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.IsDeleted"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.IsDirty"
End Property

Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As ProductSectionData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.GetState"
End Function

Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As ProductSectionData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.SetState(buffer)", buffer
End Sub
Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  SetState buffer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.SetSuperState(buffer)", buffer
End Sub


Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.Delete"
End Sub

Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    mudtProps.IsDirty = pVal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection.SetDirty(pVal)", pVal
End Sub




