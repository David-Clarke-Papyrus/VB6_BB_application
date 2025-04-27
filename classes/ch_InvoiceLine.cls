VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_InvoiceLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event rowsChange()
Event CurrRowStatus(pMsg As String)
Private mcolItems As Collection
Private mcolDeleted As Collection
Private mParent As a_Invoice
Private mlngEditing As Long
Dim MaxKeyUsed As Long
Dim WithEvents oIL As a_InvoiceLine
Attribute oIL.VB_VarHelpID = -1
'Dim ar() As Integer

'Private bCanChangeDiscount As Boolean

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
'  ReDim ar(200, 2)
  MaxKeyUsed = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.Class_Initialize"
End Sub
Public Sub DecrementMaxKeyUsed()
    If MaxKeyUsed > 0 Then MaxKeyUsed = MaxKeyUsed - 1
End Sub

Public Sub component(pParent As a_Invoice)
    On Error GoTo errHandler
    Set mParent = pParent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.Component(pParent)", pParent
End Sub
Public Function Count() As Long
    On Error GoTo errHandler
  Count = mcolItems.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.Count"
End Function
Function FindLineByCOID(pCOLID As Long) As a_InvoiceLine
    On Error GoTo errHandler
Dim oIL As a_InvoiceLine
    
    For Each oIL In mcolItems
        If oIL.COLID = pCOLID Then
            Set FindLineByCOID = oIL
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.FindLineByCOID(pCOLID)", pCOLID
End Function

Function FindLineByID(pILID As Long) As a_InvoiceLine
    On Error GoTo errHandler
Dim oIL As a_InvoiceLine
    
    For Each oIL In mcolItems
        If oIL.InvoiceLineID = pILID Then
            Set FindLineByID = oIL
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.FindLineByID(pILID)", pILID
End Function
Function FindLineBySeqNum(Seq As Long) As a_InvoiceLine
    On Error GoTo errHandler
Dim oIL As a_InvoiceLine
    
    For Each oIL In mcolItems
        If oIL.Sequence = Seq Then
            Set FindLineBySeqNum = oIL
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.FindLineBySeqNum(Seq)", Seq
End Function


Private Sub oIL_Valid(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent CurrRowStatus(pMsg)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.oIL_Valid(pMsg)", pMsg
End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
  Set NewEnum = mcolItems.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.NewEnum"
End Function

Public Function Item(ByVal Index As Variant) As a_InvoiceLine
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
    Set Item = mcolItems.Item(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.Item(Index)", Index
End Function

Public Function Add() As a_InvoiceLine
    On Error GoTo errHandler
Dim lngEdit As Long
 
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oIL = New a_InvoiceLine
  With oIL
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    oIL.Key = MaxKeyUsed & "k"
  Set Add = oIL
  Set oIL = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.Add"
End Function

Public Sub AddDocLine(Child As a_InvoiceLine)
    On Error GoTo errHandler
    Child.SetParentInvoice mParent
    Child.Sequence = mcolItems.Count + 1
    mcolItems.Add Child, Child.Key
    RaiseEvent rowsChange
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.AddDocLine(Child)", Child
End Sub

Public Sub Remove(ByVal Index As Variant)
    On Error GoTo errHandler
Dim iTmp As Long

'LogSaveToFile "Invoice line removed:" & CStr(Index)
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
    ErrorIn "ch_InvoiceLine.Remove(Index)", Index
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
  Dim oIL As a_InvoiceLine
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oIL In mcolItems
      If oIL.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.IsDirty"
End Function
Public Function IsEditing() As Boolean
    On Error GoTo errHandler
    IsEditing = mlngEditing > 0
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.IsEditing"
End Function
Public Sub BeginEdit()
    On Error GoTo errHandler
  Dim oIL As a_InvoiceLine

  For Each oIL In mcolItems
    oIL.BeginEdit
  Next
  For Each oIL In mcolDeleted
    oIL.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.BeginEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim oIL As a_InvoiceLine

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oIL In mcolItems
    oIL.ApplyEdit
  Next
  For Each oIL In mcolDeleted
    oIL.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.ApplyEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  Dim lngIndex As Long
  Dim oIL As a_InvoiceLine

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oIL In mcolItems
    If oIL.IsEditing Then oIL.CancelEdit
  Next
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oIL = mcolDeleted(lngIndex)
    With oIL
      If Not .IsNew Then
        .CancelEdit
        If Not .IsDeleted Then
          mcolItems.Add oIL, oIL.Key
          mcolDeleted.Remove lngIndex
        End If
      Else
        mcolDeleted.Remove lngIndex
      End If
    End With
    If oIL.IsEditing Then oIL.CancelEdit
    Set oIL = Nothing
  Next
    MaxKeyUsed = mcolItems.Count
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.CancelEdit"
End Sub

'Public Sub Delete()
'    On Error GoTo errHandler
'  Dim lngIndex As Long
'
'  If mlngEditing = 0 Then Err.Raise 445
'
'  For lngIndex = mcolItems.Count To 1 Step -1
'    Remove lngIndex
'  Next
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "ch_InvoiceLine.Delete"
'End Sub

Friend Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oIL As a_InvoiceLine
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oIL In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oIL.GetSuperState
    Next
    For Each oIL In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oIL.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oIL As a_InvoiceLine
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oIL = New a_InvoiceLine
      oIL.SetParentInvoice mParent
      oIL.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oIL.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oIL.BeginEdit
      Next
      If Not oIL.IsDeleted Then
        mcolItems.Add oIL, oIL.Key
      Else
        mcolDeleted.Add oIL
      End If
      oIL.CalculateLine  'recalculate dependant values
      Set oIL = Nothing
    Next
    MaxKeyUsed = .ReadProperty("Count")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.SetSuperState(buffer)", buffer
End Sub
Public Sub SortInvoiceLines(pField As enSortField, pASCEND As Boolean)
    On Error GoTo errHandler
    CollectionSort mcolItems, pField, pASCEND
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_InvoiceLine.SortInvoiceLines(pField,pASCEND)", Array(pField, pASCEND)
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
    ErrorIn "ch_InvoiceLine.swap(iUP,iDown)", Array(iUP, iDown)
End Sub

