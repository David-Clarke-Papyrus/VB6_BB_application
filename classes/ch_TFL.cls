VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_TFL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mParent As a_TF
Private mcolItems As Collection
Private colDeleted As Collection
Private mlngEditing As Long
Dim MaxKeyUsed As Long
Dim oTFL As a_TFL


Event rowsChange()

Public Sub component(pParent As a_TF)
    On Error GoTo errHandler
    Set mParent = pParent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_TFL.Component(pParent)", pParent, EA_NORERAISE
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolItems = New Collection
    Set colDeleted = New Collection
    MaxKeyUsed = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_TFL.Class_Initialize", , EA_NORERAISE
End Sub
Public Sub DecrementMaxKeyUsed()
    If MaxKeyUsed > 0 Then MaxKeyUsed = MaxKeyUsed - 1
End Sub

Public Function Count() As Long
    On Error GoTo errHandler
    Count = mcolItems.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_TFL.Count", , EA_NORERAISE
End Function

Public Function Item(ByVal Index As Variant) As a_TFL
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
    Set Item = mcolItems(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_TFL.Item(Index)", Index, EA_NORERAISE
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
    Set NewEnum = mcolItems.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_TFL.NewEnum", , EA_NORERAISE
End Function

Public Function Add() As a_TFL
    On Error GoTo errHandler
Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
    Set oTFL = New a_TFL
    With oTFL
        .Initialize Me
        For lngEdit = 1 To mlngEditing
        .BeginEdit
        Next
    End With
    MaxKeyUsed = MaxKeyUsed + 1
    oTFL.Key = MaxKeyUsed & "k"
    Set Add = oTFL
    Set oTFL = Nothing
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_TFL.Add", , EA_NORERAISE
End Function

Friend Sub AddTFLine(Child As a_TFL)
    On Error GoTo errHandler
    Child.SetParentTF mParent
    mcolItems.Add Child, Child.Key
    RaiseEvent rowsChange
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_TFL.AddTFLine(Child)", Child
End Sub

Friend Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oTFL As a_TFL
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + colDeleted.Count
    For Each oTFL In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oTFL.GetSuperState
    Next
    For Each oTFL In colDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oTFL.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_TFL.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oTFL As a_TFL
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set colDeleted = Nothing
  Set colDeleted = New Collection
  
    With objPB
        arBuffer = buffer
        .Contents = arBuffer
        For lngIndex = 1 To .ReadProperty("Count")
            Set oTFL = New a_TFL
            oTFL.SetParentTF mParent
            oTFL.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
            oTFL.Key = lngIndex & "k"
            For lngEdit = 1 To mlngEditing
              oTFL.BeginEdit
            Next
            If Not oTFL.IsDeleted Then
              mcolItems.Add oTFL, oTFL.Key
            Else
              colDeleted.Add oTFL
            End If
            oTFL.RecalculateLine  'recalculate dependant values
            Set oTFL = Nothing
        Next
        MaxKeyUsed = .ReadProperty("Count")
    End With
    Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_TFL.SetSuperState(buffer)", buffer
End Sub

Public Function IsDirty() As Boolean
    On Error GoTo errHandler
  Dim oTFL As a_TFL
  
  If colDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oTFL In mcolItems
      If oTFL.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_TFL.IsDirty", , EA_NORERAISE
End Function

Public Sub BeginEdit()
    On Error GoTo errHandler
  Dim oAL As a_TFL

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
    ErrorIn "ch_TFL.BeginEdit", , EA_NORERAISE
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim oAL As a_TFL

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
    ErrorIn "ch_TFL.ApplyEdit", , EA_NORERAISE
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  Dim lngIndex As Long
  Dim oAL As a_TFL

    If mlngEditing = 0 Then Err.Raise 445
    
    For Each oAL In mcolItems
      If oAL.IsEditing Then oAL.CancelEdit
    Next
    For lngIndex = colDeleted.Count To 1 Step -1
      Set oAL = colDeleted(lngIndex)
      With oAL
 '         .CancelEdit
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
    
    MaxKeyUsed = mcolItems.Count
    mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_TFL.CancelEdit", , EA_NORERAISE
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
    RaiseEvent rowsChange
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_TFL.Remove(Index)", Index, EA_NORERAISE
End Sub
Function FindLineByID(PID As Long) As a_TFL
    On Error GoTo errHandler
Dim oTFL As a_TFL
    
    For Each oTFL In mcolItems
        If oTFL.ID = PID Then
            Set FindLineByID = oTFL
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_TFL.FindLineByID(PID)", PID, EA_NORERAISE
End Function


Public Sub SortLines(pField As enSortField, pASCEND As Boolean)
    On Error GoTo errHandler
    CollectionSort mcolItems, pField, pASCEND
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_TFL.SortLines(pField,pASCEND)", Array(pField, pASCEND), EA_NORERAISE
End Sub

Public Sub swap(iUP As String, iDown As String)
    On Error GoTo errHandler
Dim sTemp As String
Dim iTemp As Long
    iTemp = mcolItems(iDown).Sequence
    mcolItems(iDown).Sequence = mcolItems(iUP).Sequence
    mcolItems(iUP).Sequence = iTemp
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_TFL.swap(iUP,iDown)", Array(iUP, iDown)
End Sub
