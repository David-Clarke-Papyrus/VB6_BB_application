VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_COL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event rowsChange()
Event CurrRowStatus(pMsg As String)
Private mcolItems As Collection
Private mcolDeleted As Collection
Private mParent As a_CO
Private mlngEditing As Long
Dim MaxKeyUsed As Long

Dim WithEvents oCOL As a_COL
Attribute oCOL.VB_VarHelpID = -1

Private Sub Class_Initialize()
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
  MaxKeyUsed = 0
End Sub
Public Sub component(pParent As a_CO)
    Set mParent = pParent
End Sub
Public Function Count() As Long
  Count = mcolItems.Count
End Function
Public Sub DecrementMaxKeyUsed()
    If MaxKeyUsed > 0 Then MaxKeyUsed = MaxKeyUsed - 1
End Sub


Private Sub oCOL_Valid(pMsg As String)
    RaiseEvent CurrRowStatus(pMsg)
End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As a_COL
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function

Public Function Add() As a_COL
Dim lngEdit As Long
 
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oCOL = New a_COL
  With oCOL
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    oCOL.Key = MaxKeyUsed & "k"
    oCOL.SetParentCO mParent
  Set Add = oCOL
  Set oCOL = Nothing
End Function

Friend Sub AddCOLine(Child As a_COL)
    Child.SetParentCO mParent
    mcolItems.Add Child, Child.Key
    RaiseEvent rowsChange
End Sub

Public Sub Remove(ByVal Index As Variant)
  If mlngEditing = 0 Then Err.Raise 445

    With mcolItems(Index)
      .BeginEdit
      .Delete
      .ApplyEdit
    End With
    mcolDeleted.Add mcolItems(Index)
    mcolItems.Remove Index
    RaiseEvent rowsChange
End Sub

Public Function IsDirty() As Boolean
  Dim oCOL As a_COL
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oCOL In mcolItems
      If oCOL.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function
Public Function IsEditing() As Boolean
    IsEditing = mlngEditing > 0
End Function
Public Sub BeginEdit()
  Dim oCOL As a_COL

  For Each oCOL In mcolItems
    oCOL.BeginEdit
  Next
  For Each oCOL In mcolDeleted
    oCOL.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub ApplyEdit()
  Dim oCOL As a_COL

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oCOL In mcolItems
    oCOL.ApplyEdit
  Next
  For Each oCOL In mcolDeleted
    oCOL.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
End Sub

Public Sub CancelEdit()
  Dim lngIndex As Long
  Dim oCOL As a_COL

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oCOL In mcolItems
    If oCOL.IsEditing Then oCOL.CancelEdit
  Next
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oCOL = mcolDeleted(lngIndex)
    With oCOL
        .CancelEdit
        If Not .IsNew Then
            If Not .IsDeleted Then
                mcolItems.Add oCOL, oCOL.Key
                mcolDeleted.Remove lngIndex
            End If
        Else
            mcolDeleted.Remove lngIndex
        End If
    End With
    If oCOL.IsEditing Then oCOL.CancelEdit
    Set oCOL = Nothing
  Next
  MaxKeyUsed = mcolItems.Count
  
  mlngEditing = mlngEditing - 1
End Sub

Public Sub Delete()
  Dim lngIndex As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  For lngIndex = mcolItems.Count To 1 Step -1
    Remove lngIndex
  Next
End Sub

Friend Function GetSuperState() As String
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oCOL As a_COL
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oCOL In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oCOL.GetSuperState
    Next
    For Each oCOL In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oCOL.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
End Function

Friend Sub SetSuperState(buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oCOL As a_COL
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oCOL = New a_COL
      oCOL.SetParentCO mParent
      oCOL.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oCOL.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oCOL.BeginEdit
      Next
      If Not oCOL.IsDeleted Then
        mcolItems.Add oCOL, oCOL.Key
      Else
        mcolDeleted.Add oCOL
      End If
      Set oCOL = Nothing
    Next
    MaxKeyUsed = .ReadProperty("Count")
  End With
  Set objPB = Nothing
End Sub
'Public Sub CanChangeDiscount(val As Boolean)
'     bCanChangeDiscount = val
'End Sub
Function FindLineByID(PID As Long) As a_COL
Dim oCOL As a_COL
    
    For Each oCOL In mcolItems
        If oCOL.COLineID = PID Then
            Set FindLineByID = oCOL
            Exit For
        End If
    Next
End Function

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

