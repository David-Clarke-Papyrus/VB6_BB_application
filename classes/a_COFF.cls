VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_COFF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mudtProps As COFFProps
Private mcolStack As Collection
Private mobjParent As ch_COFF
Event Valid(IsValid As Boolean)
Private strParentType As String
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private mOwnerCollection As ch_COFF
Dim mKey As String
Event QtyChanged()


Sub Parent(Value As Object)
    On Error GoTo errHandler
    Set mobjParent = Value
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.Parent(Value)", Value
End Sub

Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.BeginEdit"
End Sub
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = (mcolStack.Count > 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.IsEditing"
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
    ErrorIn "a_COFF.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  If Not mobjParent Is Nothing Then
    mobjParent.AddCOFF Me
    Set mobjParent = Nothing
  End If
  mcolStack.Remove mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.ApplyEdit"
End Sub
Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    mudtProps.IsNew = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.Class_Initialize"
End Sub
Friend Sub Initialize(Parent As ch_COFF)
    On Error GoTo errHandler
  If mudtProps.IsNew Then Set mobjParent = Parent
  Set mOwnerCollection = Parent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.Initialize(Parent)", Parent
End Sub


Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.IsValid"
End Property

Private Sub mobjValid_BrokenRule(pors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.mobjValid_BrokenRule(pors)", pors
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid(True)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.mobjValid_NoBrokenRules"
End Sub
Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  mudtProps.IsDirty = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.Delete"
End Sub
Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.IsDeleted"
End Property
Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.IsNew"
End Property
Public Property Let IsNew(val As Boolean)
    On Error GoTo errHandler
    mudtProps.IsNew = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.IsNew(val)", val
End Property
Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.IsDirty"
End Property
Friend Function GetState() As String
    On Error GoTo errHandler
Dim udtData As COFFData
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.GetState"
End Function
Friend Function SetState(buffer As String)
    On Error GoTo errHandler
Dim udtData As COFFData
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.SetState(buffer)", buffer
End Function
Public Property Get ID() As Long
    On Error GoTo errHandler
    ID = mudtProps.COFFID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.ID"
End Property

Public Property Let COCode(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COCode = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COCode(val)", val
End Property
Public Property Get COCode() As String
    On Error GoTo errHandler
    COCode = mudtProps.COCode
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COCode"
End Property
Public Property Let COLID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COLID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COLID(val)", val
End Property
Public Property Get COLID() As Long
    On Error GoTo errHandler
    COLID = mudtProps.COLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COLID"
End Property


Public Property Let COFFCOLID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COFFCOLID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFCOLID(val)", val
End Property
Public Property Get COFFCOLID() As Long
    On Error GoTo errHandler
    COFFCOLID = mudtProps.COFFCOLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFCOLID"
End Property
Public Property Let COLQtyDispatched(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COLQtyDispatched = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COLQtyDispatched(val)", val
End Property
Public Property Get COLQtyDispatched() As Long
    On Error GoTo errHandler
    COLQtyDispatched = mudtProps.COLQtyDispatched
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COLQtyDispatched"
End Property

Public Property Let CODate(val As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CODate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.CODate(val)", val
End Property
Public Property Get CODate() As Date
    On Error GoTo errHandler
    CODate = mudtProps.CODate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.CODate"
End Property
Public Property Get CODateF() As String
    On Error GoTo errHandler
    CODateF = Format(mudtProps.CODate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.CODateF"
End Property


Public Property Let COLQty(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COLQty = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COLQty(val)", val
End Property
Public Property Get COLQty() As Long
    On Error GoTo errHandler
    COLQty = mudtProps.COLQty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COLQty"
End Property

Public Property Let COFFILID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COFFILID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFILID(val)", val
End Property
Public Property Get COFFILID() As Long
    On Error GoTo errHandler
    COFFILID = mudtProps.COFFILID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFILID"
End Property
Public Property Let COFFCOLLID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COFFCOLLID = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFCOLLID(val)", val
End Property
Public Property Get COFFCOLLID() As Long
    On Error GoTo errHandler
    COFFCOLLID = mudtProps.COFFCOLLID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFCOLLID"
End Property

Public Function SetCOFFQTY(val As Long) As Boolean
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COFFQTY = val
    mudtProps.IsDirty = True
'    RaiseEvent QtyChanged
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.SetCOFFQTY(val)", val
End Function

Public Property Get COFFQTY() As Long
    On Error GoTo errHandler
    COFFQTY = mudtProps.COFFQTY
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFQTY"
End Property
Public Property Let COFFILQTY(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COFFILQTY = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFILQTY(val)", val
End Property
Public Property Get COFFILQTY() As Long
    On Error GoTo errHandler
    COFFILQTY = mudtProps.COFFILQTY
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFILQTY"
End Property
Public Property Let COFFCOLQTY(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COFFCOLQTY = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFCOLQTY(val)", val
End Property
Public Property Get COFFCOLQTY() As Long
    On Error GoTo errHandler
    COFFCOLQTY = mudtProps.COFFCOLQTY
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFCOLQTY"
End Property
Public Property Let COFFCOCode(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COFFCOCode = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFCOCode(val)", val
End Property
Public Property Get COFFCOCode() As String
    On Error GoTo errHandler
    COFFCOCode = mudtProps.COFFCOCode
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFCOCode"
End Property
Public Property Let COFFINVCode(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COFFINVCode = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFINVCode(val)", val
End Property
Public Property Get COFFINVCode() As String
    On Error GoTo errHandler
    COFFINVCode = mudtProps.COFFINVCode
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFINVCode"
End Property
Public Property Let COFFCode(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COFFCode = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFCode(val)", val
End Property
Public Property Get COFFCode() As String
    On Error GoTo errHandler
    COFFCode = mudtProps.COFFCode
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.COFFCode"
End Property


Friend Function GetSuperState() As String
    On Error GoTo errHandler
  GetSuperState = GetState
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String, pOwnerCollection As ch_COFF)
    On Error GoTo errHandler
  SetState buffer
  Set mOwnerCollection = pOwnerCollection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.SetSuperState(buffer,pOwnerCollection)", Array(buffer, pOwnerCollection)
End Sub

Public Property Get Key() As String
    On Error GoTo errHandler
    Key = mKey
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF.Key"
End Property
Public Property Let Key(val As String)
    mKey = val
End Property


