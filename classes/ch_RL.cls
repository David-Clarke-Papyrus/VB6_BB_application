VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_RL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mParent As a_R
Private mcolItems As Collection
Private colDeleted As Collection
Private mlngEditing As Long
Dim MaxKeyUsed As Long
Dim ooRL As a_RL
Event rowsChange()

Public Sub component(pParent As a_R)
    On Error GoTo errHandler
    Set mParent = pParent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.Component(pParent)", pParent
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set mcolItems = New Collection
  Set colDeleted = New Collection
  MaxKeyUsed = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.Class_Initialize"
End Sub

Public Function Item(ByVal Index As Variant) As a_RL
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
  Set Item = mcolItems(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.Item(Index)", Index
End Function

Public Function Count() As Long
    On Error GoTo errHandler
  Count = mcolItems.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.Count"
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
  Set NewEnum = mcolItems.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.NewEnum"
End Function

Friend Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim ooRL As a_RL
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + colDeleted.Count
    For Each ooRL In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), ooRL.GetSuperState
    Next
    For Each ooRL In colDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), ooRL.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim ooRL As a_RL
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set colDeleted = Nothing
  Set colDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set ooRL = New a_RL
      ooRL.SetParentR mParent
      ooRL.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      ooRL.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        ooRL.BeginEdit
      Next
      If Not ooRL.IsDeleted Then
        mcolItems.Add ooRL, ooRL.Key
      Else
        colDeleted.Add ooRL
      End If
      Set ooRL = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex - 1
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.SetSuperState(buffer)", buffer
End Sub

Public Function IsDirty() As Boolean
    On Error GoTo errHandler
  Dim ooRL As a_RL
  
  If colDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each ooRL In mcolItems
      If ooRL.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.IsDirty"
End Function

Public Sub BeginEdit()
    On Error GoTo errHandler
  Dim oAL As a_RL

  For Each oAL In mcolItems
    oAL.BeginEdit
  Next
  For Each oAL In colDeleted
    oAL.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.BeginEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim oAL As a_RL

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oAL In mcolItems
    oAL.ApplyEdit
  Next
  For Each oAL In colDeleted
    oAL.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.ApplyEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  Dim lngIndex As Long
  Dim oAL As a_RL

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oAL In mcolItems
    oAL.CancelEdit
  Next
  For lngIndex = colDeleted.Count To 1 Step -1
    Set oAL = colDeleted(lngIndex)
    With oAL
      If Not .IsNew Then
        .CancelEdit
        If Not .IsDeleted Then
          mcolItems.Add oAL, oAL.Key
          colDeleted.Remove lngIndex
        End If
      Else
        colDeleted.Remove lngIndex
      End If
    End With
    If oAL.IsEditing Then oAL.CancelEdit
    Set oAL = Nothing
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.CancelEdit"
End Sub

Public Function Add() As a_RL
    On Error GoTo errHandler
Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set ooRL = New a_RL
  With ooRL
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    ooRL.Key = MaxKeyUsed & "k"
    ooRL.SetParentR mParent
  Set Add = ooRL
  Set ooRL = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.Add"
End Function

Friend Sub AddRLine(Child As a_RL)
    On Error GoTo errHandler
    Child.SetParentR mParent
    mcolItems.Add Child, Child.Key
    RaiseEvent rowsChange
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.AddRLine(Child)", Child
End Sub

Public Sub Remove(ByVal Index As Variant)
    On Error GoTo errHandler
  If mlngEditing = 0 Then Err.Raise 445

  With mcolItems(Index)
    .BeginEdit
    .Delete
    .ApplyEdit
  End With
  colDeleted.Add mcolItems(Index)
  mcolItems.Remove Index
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.Remove(Index)", Index
End Sub
Function FindLineByID(PID As Long) As a_RL
    On Error GoTo errHandler
Dim ooRL As a_RL
    
    For Each ooRL In mcolItems
        If ooRL.ID = PID Then
            Set FindLineByID = ooRL
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.FindLineByID(pID)", PID
End Function

Public Sub SortLines(pField As enSortField, pASCEND As Boolean)
    On Error GoTo errHandler
    CollectionSort mcolItems, pField, pASCEND
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_RL.SortLines(pField,pASCEND)", Array(pField, pASCEND)
End Sub



