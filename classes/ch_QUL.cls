VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_QUL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event rowsChange()
Event CurrRowStatus(pMsg As String)
Private mcolItems As Collection
Private mcolDeleted As Collection
Private mParent As a_QU
Private mlngEditing As Long
Dim MaxKeyUsed As Long
Dim WithEvents oQUL As a_QUL
Attribute oQUL.VB_VarHelpID = -1


Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
  MaxKeyUsed = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.Class_Initialize"
End Sub
Public Sub DecrementMaxKeyUsed()
    If MaxKeyUsed > 0 Then MaxKeyUsed = MaxKeyUsed - 1
End Sub
Public Sub component(pParent As a_QU)
    On Error GoTo errHandler
    Set mParent = pParent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.Component(pParent)", pParent
End Sub
Public Function Count() As Long
    On Error GoTo errHandler
  Count = mcolItems.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.Count"
End Function
Function FindLineByCOID(pCOLID As Long) As a_QUL
    On Error GoTo errHandler
Dim oQUL As a_QUL
    
    For Each oQUL In mcolItems
        If oQUL.COLID = pCOLID Then
            Set FindLineByCOID = oQUL
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.FindLineByCOID(pCOLID)", pCOLID
End Function

Function FindLineByID(pQULID As Long) As a_QUL
    On Error GoTo errHandler
Dim oQUL As a_QUL
    
    For Each oQUL In mcolItems
        If oQUL.QuoteLineID = pQULID Then
            Set FindLineByID = oQUL
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.FindLineByID(pQULID)", pQULID
End Function
'Function FindLineBySeqNum(Seq As Long) As a_QUL
'    On Error GoTo ErrHandler
'Dim oQUL As a_QUL
'
'    For Each oQUL In mcolItems
'        If oQUL.Sequence = Seq Then
'            Set FindLineBySeqNum = oQUL
'            Exit For
'        End If
'    Next
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ch_QUL.FindLineBySeqNum(Seq)", Seq
'End Function
Function FindLineBySeqNum(Seq As Long) As a_QUL
    On Error GoTo errHandler
Dim ol As a_QUL
    
    For Each ol In mcolItems
        If ol.Sequence = Seq Then
            Set FindLineBySeqNum = ol
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.FindLineBySeqNum(Seq)", Seq
End Function


Private Sub oQUL_Valid(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent CurrRowStatus(pMsg)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.oQUL_Valid(pMsg)", pMsg
End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
  Set NewEnum = mcolItems.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.NewEnum"
End Function

Public Function Item(ByVal Index As Variant) As a_QUL
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
    Set Item = mcolItems.Item(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.Item(Index)", Index
End Function

Public Function Add() As a_QUL
    On Error GoTo errHandler
Dim lngEdit As Long
 
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oQUL = New a_QUL
  With oQUL
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    oQUL.Key = MaxKeyUsed & "k"
  Set Add = oQUL
  Set oQUL = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.Add"
End Function

Friend Sub AddQuoteLine(Child As a_QUL)
    On Error GoTo errHandler
    Child.SetParentQuote mParent
    Child.Sequence = mcolItems.Count + 1
    mcolItems.Add Child, Child.Key
    RaiseEvent rowsChange
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.AddQuoteLine(Child)", Child
End Sub

Public Sub Remove(ByVal Index As Variant)
    On Error GoTo errHandler
Dim iTmp As Long

  If mlngEditing = 0 Then Err.Raise 445

    With mcolItems(Index)
        iTmp = .Sequence
        .BeginEdit
        .Delete
        .ApplyEdit
    End With
    mcolDeleted.Add mcolItems(Index)
    mcolItems.Remove Index
    CloseUpSequence iTmp
    RaiseEvent rowsChange
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.Remove(Index)", Index
End Sub
Private Sub CloseUpSequence(pTmp As Long)
Dim i As Integer

    For i = 1 To mcolItems.Count
        If mcolItems(i).Sequence > pTmp Then
            mcolItems(i).Sequence = mcolItems(i).Sequence - 1
        End If
    Next i
End Sub
Public Function IsDirty() As Boolean
    On Error GoTo errHandler
  Dim oQUL As a_QUL
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oQUL In mcolItems
      If oQUL.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.IsDirty"
End Function
Public Function IsEditing() As Boolean
    On Error GoTo errHandler
    IsEditing = mlngEditing > 0
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.IsEditing"
End Function
Public Sub BeginEdit()
    On Error GoTo errHandler
  Dim oQUL As a_QUL

  For Each oQUL In mcolItems
    oQUL.BeginEdit
  Next
  For Each oQUL In mcolDeleted
    oQUL.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.BeginEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim oQUL As a_QUL

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oQUL In mcolItems
    oQUL.ApplyEdit
  Next
  For Each oQUL In mcolDeleted
    oQUL.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.ApplyEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  Dim lngIndex As Long
  Dim oQUL As a_QUL

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oQUL In mcolItems
    If oQUL.IsEditing Then oQUL.CancelEdit
  Next
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oQUL = mcolDeleted(lngIndex)
    With oQUL
      If Not .IsNew Then
        .CancelEdit
        If Not .IsDeleted Then
          mcolItems.Add oQUL, oQUL.Key
          mcolDeleted.Remove lngIndex
        End If
      Else
        mcolDeleted.Remove lngIndex
      End If
    End With
    If oQUL.IsEditing Then oQUL.CancelEdit
    Set oQUL = Nothing
  Next
    MaxKeyUsed = mcolItems.Count

  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.CancelEdit"
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
    ErrorIn "ch_QUL.Delete"
End Sub

Friend Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oQUL As a_QUL
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oQUL In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oQUL.GetSuperState
    Next
    For Each oQUL In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oQUL.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oQUL As a_QUL
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oQUL = New a_QUL
      oQUL.SetParentQuote mParent
      oQUL.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oQUL.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oQUL.BeginEdit
      Next
      If Not oQUL.IsDeleted Then
        mcolItems.Add oQUL, oQUL.Key
      Else
        mcolDeleted.Add oQUL
      End If
      oQUL.CalculateLine  'recalculate dependant values
      Set oQUL = Nothing
    Next
    MaxKeyUsed = .ReadProperty("Count")
  End With

  Set objPB = Nothing
  
  
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.SetSuperState(buffer)", buffer
End Sub
Public Sub SortLines(pField As enSortField, pASCEND As Boolean)
    On Error GoTo errHandler
    CollectionSort mcolItems, pField, pASCEND
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_QUL.SortLines(pField,pASCEND)", Array(pField, pASCEND)
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
    ErrorIn "ch_QUL.swap(iUP,iDown)", Array(iUP, iDown)
End Sub



