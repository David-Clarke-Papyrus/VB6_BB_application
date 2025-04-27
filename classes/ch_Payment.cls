VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_Payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mcolItems As Collection
Private mcolDeleted As Collection
Private mParent As a_Exchange
Private mlngEditing As Long
Dim MaxKeyUsed As Long
Dim WithEvents oPayment As a_Payment
Attribute oPayment.VB_VarHelpID = -1
Dim ar() As Integer

'Private bCanChangeDiscount As Boolean

Private Sub Class_Initialize()
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
  ReDim ar(200, 2)
  MaxKeyUsed = 0
End Sub
Public Sub component(pParent As a_Exchange)
    Set mParent = pParent
End Sub
Public Function Count() As Long
  Count = mcolItems.Count
End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As a_Payment
Attribute Item.VB_UserMemId = 0
    Set Item = mcolItems.Item(Index)
End Function
Public Function TotalVoucherValue() As Long
Dim oPay As a_Payment
Dim lngTotal As Long

    For Each oPay In mcolItems
        If oPay.PaymentType = "V" Then
            lngTotal = lngTotal + oPay.Amt
        End If
    Next
    TotalVoucherValue = lngTotal
End Function
Public Function Add() As a_Payment
Dim lngEdit As Long
 
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oPayment = New a_Payment
  With oPayment
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    oPayment.Key = MaxKeyUsed & "k"
  Set Add = oPayment
  Set oPayment = Nothing
End Function

Friend Sub AddPaymentLine(Child As a_Payment)
    Child.SetParentExchange mParent
'    Child.Sequence = mcolItems.Count + 1
    mcolItems.Add Child, Child.Key
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
End Sub

Public Function IsDirty() As Boolean
  Dim oPayment As a_Payment
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oPayment In mcolItems
      If oPayment.IsDirty Then
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
    On Error GoTo errHandler
  Dim oPayment As a_Payment

  For Each oPayment In mcolItems
    oPayment.BeginEdit
  Next
  For Each oPayment In mcolDeleted
    oPayment.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Payment.BeginEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim oPayment As a_Payment

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oPayment In mcolItems
    oPayment.ApplyEdit
  Next
  For Each oPayment In mcolDeleted
    oPayment.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Payment.ApplyEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  Dim lngIndex As Long
  Dim oPayment As a_Payment

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oPayment In mcolItems
    oPayment.CancelEdit
  Next
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oPayment = mcolDeleted(lngIndex)
    With oPayment
      If Not .IsNew Then
        .CancelEdit
        If Not .IsDeleted Then
          mcolItems.Add oPayment
          mcolDeleted.Remove lngIndex
        End If
      Else
        mcolDeleted.Remove lngIndex
      End If
    End With
    If oPayment.IsEditing Then oPayment.CancelEdit
    Set oPayment = Nothing
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_sale.CancelEdit"
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
  Dim oPayment As a_Payment
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oPayment In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oPayment.GetSuperState
    Next
    For Each oPayment In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oPayment.GetSuperState
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
  Dim oPayment As a_Payment
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oPayment = New a_Payment
      oPayment.SetParentExchange mParent
      oPayment.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oPayment.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oPayment.BeginEdit
      Next
      If Not oPayment.IsDeleted Then
        mcolItems.Add oPayment, oPayment.Key
      Else
        mcolDeleted.Add oPayment
      End If
      oPayment.CalculateLine  'recalculate dependant values
      Set oPayment = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex
  Set objPB = Nothing
  
  
End Sub
Public Sub SortPaymentLines(pField As enSortField, pASCEND As Boolean)
    CollectionSort mcolItems, pField, pASCEND
End Sub

Private Function CollectionSort(ByRef oCollection As Collection, pProperty As enSortField, Optional bSortAscending As Boolean = True) As Long
    Dim lSort1 As Long, lSort2 As Long
    Dim vTempItem1 As Variant, vTempItem2 As Variant, bSwap As Boolean
    
    On Error GoTo ErrFailed
    For lSort1 = 1 To oCollection.Count - 1
        For lSort2 = lSort1 + 1 To oCollection.Count
            If bSortAscending Then
                If oCollection(lSort1).Properties(pProperty) > oCollection(lSort2).Properties(pProperty) Then
                    bSwap = True
                Else
                    bSwap = False
                End If
            Else
                If oCollection(lSort1).Properties(pProperty) < oCollection(lSort2).Properties(pProperty) Then
                    bSwap = True
                Else
                    bSwap = False
                End If
            End If
            If bSwap Then
                'Store the items
                If VarType(oCollection(lSort1)) = vbObject Then
                    Set vTempItem1 = oCollection(lSort1)
                Else
                    vTempItem1 = oCollection(lSort1)
                End If
                
                If VarType(oCollection(lSort2)) = vbObject Then
                    Set vTempItem2 = oCollection(lSort2)
                Else
                    vTempItem2 = oCollection(lSort2)
                End If
                
                'Swap the items over
                oCollection.Add vTempItem1, , lSort2
                oCollection.Add vTempItem2, , lSort1
                'Delete the original items
                oCollection.Remove lSort1 + 1
                oCollection.Remove lSort2 + 1
            End If
        Next
    Next
    Exit Function

ErrFailed:
    Debug.Print "Error with CollectionSort: " & Err.Description
    CollectionSort = Err.Number
    On Error GoTo 0
End Function

Public Property Get NeedsCustomerInfo() As Boolean
Dim oPay As a_Payment
Dim i As Integer

    NeedsCustomerInfo = False
    For i = 1 To mcolItems.Count
        If mcolItems(i).PaymentType = "DDP" Then
            NeedsCustomerInfo = True
        End If
    Next
End Property

