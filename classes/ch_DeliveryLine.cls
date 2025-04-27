VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_DeliveryLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event rowsChange()
Event CurrRowStatus(pMsg As String)
Private mParent As a_Delivery
Private mcolItems As Collection
Private mcolDeleted As Collection
Private mlngEditing As Long
Dim MaxKeyUsed As Long
Dim WithEvents oDELL As a_DeliveryLine
Attribute oDELL.VB_VarHelpID = -1

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
  MaxKeyUsed = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.Class_Initialize"
End Sub
Public Sub DecrementMaxKeyUsed()
    If MaxKeyUsed > 0 Then MaxKeyUsed = MaxKeyUsed - 1
End Sub

Public Sub component(pParent As a_Delivery)
    On Error GoTo errHandler
    Set mParent = pParent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.Component(pParent)", pParent
End Sub



Public Function Count() As Long
    On Error GoTo errHandler
  Count = mcolItems.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.Count"
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
  Set NewEnum = mcolItems.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.NewEnum"
End Function

Private Sub oDELL_Valid(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent CurrRowStatus(pMsg)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.oDELL_Valid(pMsg)", pMsg
End Sub

Public Function Item(ByVal Index As Variant) As a_DeliveryLine
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
  Set Item = mcolItems.Item(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.Item(Index)", Index, , , "Array of all indexes,arg", Array(GetAllIndexes(), Index, mcolItems.Count)
End Function
Private Function GetAllIndexes() As String
    On Error GoTo errHandler
Dim i As Integer
Dim strIndexes As String

    For i = 1 To mcolItems.Count
        strIndexes = strIndexes & " " & mcolItems(i).Key
    Next i
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.GetAllIndexes"
End Function

Friend Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oDL As a_DeliveryLine
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oDL In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oDL.GetSuperState
    Next
    For Each oDL In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oDL.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oDL As a_DeliveryLine
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oDL = New a_DeliveryLine
      oDL.SetParentDEL mParent
      oDL.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oDL.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oDL.BeginEdit
      Next
      If Not oDL.IsDeleted Then
        mcolItems.Add oDL, oDL.Key
      Else
        mcolDeleted.Add oDL
      End If
      Set oDL = Nothing
    Next
    MaxKeyUsed = .ReadProperty("Count")
  End With
'  MaxKeyUsed = lngIndex - 1
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.SetSuperState(buffer)", buffer
End Sub

Public Function Add() As a_DeliveryLine
    On Error GoTo errHandler
Dim lngEdit As Long
 
    If mlngEditing = 0 Then Err.Raise 445
    
    Set oDELL = New a_DeliveryLine
    With oDELL
        .Initialize Me
        For lngEdit = 1 To mlngEditing
          .BeginEdit
        Next
    End With
    MaxKeyUsed = MaxKeyUsed + 1
    oDELL.Key = MaxKeyUsed & "k"
    oDELL.SetParentDEL mParent
   Set Add = oDELL
   Set oDELL = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.Add"
End Function

Friend Sub AddDELLine(Child As a_DeliveryLine)
    On Error GoTo errHandler
    Child.SetParentDEL mParent
    mcolItems.Add Child, Child.Key
    RaiseEvent rowsChange
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.AddDELLine(Child)", Child
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
  '  MaxKeyUsed = mcolItems.Count
    RaiseEvent rowsChange
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.Remove(Index)", Index
End Sub
Public Function IsDirty() As Boolean
    On Error GoTo errHandler
  Dim oDELL As a_DeliveryLine
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oDELL In mcolItems
      If oDELL.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.IsDirty"
End Function
Public Function IsEditing() As Boolean
    On Error GoTo errHandler
    IsEditing = mlngEditing > 0
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.IsEditing"
End Function
Public Sub BeginEdit()
    On Error GoTo errHandler
  Dim oDELL As a_DeliveryLine

  For Each oDELL In mcolItems
    oDELL.BeginEdit
  Next
  For Each oDELL In mcolDeleted
    oDELL.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.BeginEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim oDELL As a_DeliveryLine

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oDELL In mcolItems
    oDELL.ApplyEdit
  Next
  For Each oDELL In mcolDeleted
    oDELL.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.ApplyEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  Dim lngIndex As Long
  Dim oDELL As a_DeliveryLine

    If mlngEditing = 0 Then Err.Raise 445
    
    For Each oDELL In mcolItems
      If oDELL.IsEditing Then oDELL.CancelEdit
    Next
    For lngIndex = mcolDeleted.Count To 1 Step -1
      Set oDELL = mcolDeleted(lngIndex)
      With oDELL
 '         .CancelEdit
          If Not .IsNew Then
            .CancelEdit
              If Not .IsDeleted Then
                mcolItems.Add oDELL, oDELL.Key
                mcolDeleted.Remove lngIndex
              End If
          Else
              mcolDeleted.Remove lngIndex
          End If
      End With
      If oDELL.IsEditing Then oDELL.CancelEdit
      Set oDELL = Nothing
    Next
    MaxKeyUsed = mcolItems.Count
    
    mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.CancelEdit"
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
    ErrorIn "ch_DeliveryLine.Delete"
End Sub
Function FindLineByID(PID As Long) As a_DeliveryLine
    On Error GoTo errHandler
Dim oDELL As a_DeliveryLine
    
    For Each oDELL In mcolItems
        If oDELL.DELLID = PID Then
            Set FindLineByID = oDELL
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.FindLineByID(PID)", PID
End Function
Function FindLineByEAN(pEAN As String) As a_DeliveryLine
    On Error GoTo errHandler
Dim oDELL As a_DeliveryLine
    
    Set FindLineByEAN = Nothing
    For Each oDELL In mcolItems
        If oDELL.EAN = pEAN Then
            Set FindLineByEAN = oDELL
            Exit For
        End If
    Next
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ch_DeliveryLine.FindLineByID(PID)", PID
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.FindLineByEAN(pEAN)", pEAN
End Function

Public Sub SortLines(pField As enSortField, pASCEND As Boolean)
    On Error GoTo errHandler
    CollectionSort mcolItems, pField, pASCEND
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_DeliveryLine.SortLines(pField,pASCEND)", Array(pField, pASCEND)
End Sub

