VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_POL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event rowsChange()
Event CurrRowStatus(pMsg As String)
Private mcolItems As Collection
Private mcolDeleted As Collection
Private mParent As a_PO
Private mlngEditing As Long
Dim MaxKeyUsed As Long

Dim WithEvents oPOL As a_POL
Attribute oPOL.VB_VarHelpID = -1

Public Sub component(pParent As a_PO)
    On Error GoTo errHandler
    Set mParent = pParent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.Component(pParent)", pParent
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolItems = New Collection
    Set mcolDeleted = New Collection
    MaxKeyUsed = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.Class_Initialize"
End Sub

Public Function Count() As Long
    On Error GoTo errHandler
    Count = mcolItems.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.Count"
End Function

Public Function Item(ByVal Index As Variant) As a_POL
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
    Set Item = mcolItems.Item(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.Item(Index)", Index
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    On Error GoTo errHandler
    Set NewEnum = mcolItems.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.NewEnum"
End Function

Private Sub oPOL_Valid(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent CurrRowStatus(pMsg)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.oPOL_Valid(pMsg)", pMsg
End Sub


Public Function Add() As a_POL
    On Error GoTo errHandler
Dim lngEdit As Long
 
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oPOL = New a_POL
  With oPOL
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    oPOL.Key = MaxKeyUsed & "k"
    Set Add = oPOL
    oPOL.SetParentPO mParent
    Set oPOL = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.Add"
End Function

Friend Sub AddSOLine(Child As a_POL)
    On Error GoTo errHandler
    Child.SetParentPO mParent
    Child.Sequence = mcolItems.Count + 1
    mcolItems.Add Child, Child.Key
    RaiseEvent rowsChange
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.AddSOLine(Child)", Child
End Sub

Friend Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oSO As a_POL
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oSO In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oSO.GetSuperState
    Next
    For Each oSO In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oSO.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oPOL As a_POL
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
    With objPB
        arBuffer = buffer
        .Contents = arBuffer
        For lngIndex = 1 To .ReadProperty("Count")
            Set oPOL = New a_POL
            oPOL.SetParentPO mParent
            oPOL.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
            oPOL.Key = lngIndex & "k"
            For lngEdit = 1 To mlngEditing
              oPOL.BeginEdit
            Next
            If Not oPOL.IsDeleted Then
              mcolItems.Add oPOL, oPOL.Key
            Else
              mcolDeleted.Add oPOL
            End If
            Set oPOL = Nothing
        Next
        MaxKeyUsed = .ReadProperty("Count")
    End With
    Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.SetSuperState(buffer)", buffer
End Sub

Public Sub Remove(ByVal Index As Variant)
    On Error GoTo errHandler
  If mlngEditing = 0 Then Err.Raise 445

    With mcolItems(Index)
      .BeginEdit
      .Delete
      .ApplyEdit
    End With
    mcolDeleted.Add mcolItems(Index)
    mcolItems.Remove Index
    RaiseEvent rowsChange
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.Remove(Index)", Index
End Sub

Public Function IsDirty() As Boolean
    On Error GoTo errHandler
  Dim oSO As a_POL
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oSO In mcolItems
      If oSO.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.IsDirty"
End Function
Public Function IsEditing() As Boolean
    On Error GoTo errHandler
    IsEditing = mlngEditing > 0
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.IsEditing"
End Function
Public Sub BeginEdit()
    On Error GoTo errHandler
  Dim oSO As a_POL

  For Each oSO In mcolItems
    oSO.BeginEdit
  Next
  For Each oSO In mcolDeleted
    oSO.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.BeginEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim oSO As a_POL

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oSO In mcolItems
    oSO.ApplyEdit
  Next
  For Each oSO In mcolDeleted
    oSO.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.ApplyEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  Dim lngIndex As Long
  Dim oPOL As a_POL

    If mlngEditing = 0 Then Err.Raise 445
    
    For Each oPOL In mcolItems
      If oPOL.IsEditing Then oPOL.CancelEdit
    Next
    For lngIndex = mcolDeleted.Count To 1 Step -1
      Set oPOL = mcolDeleted(lngIndex)
      With oPOL
          If Not .IsNew Then
            .CancelEdit
              If Not .IsDeleted Then
                  mcolItems.Add oPOL, oPOL.Key
                  mcolDeleted.Remove lngIndex
              End If
          Else
              mcolDeleted.Remove lngIndex
          End If
      End With
      If oPOL.IsEditing Then oPOL.CancelEdit
      Set oPOL = Nothing
    Next
    
    mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.CancelEdit"
End Sub

Public Sub Delete()
    On Error GoTo errHandler
  Dim lngIndex As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  For lngIndex = mcolItems.Count To 1 Step -1
    Remove lngIndex
  Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.Delete"
End Sub
Function FindLineByID(PID As Long) As a_POL
    On Error GoTo errHandler
Dim oPOL As a_POL
    
    For Each oPOL In mcolItems
        If oPOL.POLID = PID Then
            Set FindLineByID = oPOL
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_POL.FindLineByID(PID)", PID
End Function
Public Sub SortPOLines(pField As enSortField, pASCEND As Boolean)
    On Error GoTo errHandler
    CollectionSort mcolItems, pField, pASCEND
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_IPOL.SortIPOLines(pField,pASCEND)", Array(pField, pASCEND)
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
    ErrorIn "ch_POL.swap(iUP,iDown)", Array(iUP, iDown)
End Sub
Function CancelInactiveLines()
Dim oPOL As a_POL
    
    For Each oPOL In mcolItems
        If oPOL.Fulfilled = "OS" Then
            If oPOL.QtyReceivedSoFar = 0 Then
                oPOL.CancelLine
            Else
                oPOL.MarkLineasFulfilled
            End If
        End If
    Next
End Function

