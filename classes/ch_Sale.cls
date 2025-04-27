VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_sale"
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
Dim WithEvents oSALE As a_Sale
Attribute oSALE.VB_VarHelpID = -1
Dim ar() As Integer
Event ContainsLines(pYesNo As Boolean)

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
Function FindLineByID(saleID As Long) As a_Sale
Dim oSALE As a_Sale

    For Each oSALE In mcolItems
        If oSALE.saleID = saleID Then
            Set FindLineByID = oSALE
            Exit For
        End If
    Next
End Function
Function FindLineBySeqNum(Seq As Long) As a_Sale
Dim oSALE As a_Sale
    
    For Each oSALE In mcolItems
        If oSALE.Sequence = Seq Then
            Set FindLineBySeqNum = oSALE
            Exit For
        End If
    Next
End Function


'Private Sub oSale_Valid(pMsg As String)
'    RaiseEvent CurrRowStatus(pMsg)
'End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As a_Sale
Attribute Item.VB_UserMemId = 0
    Set Item = mcolItems.Item(Index)
End Function

Public Function Add() As a_Sale
Dim lngEdit As Long
 
    If mlngEditing = 0 Then Err.Raise 445
    
    Set oSALE = New a_Sale
    With oSALE
      .Initialize Me
      For lngEdit = 1 To mlngEditing
        .BeginEdit
      Next
    End With
    MaxKeyUsed = MaxKeyUsed + 1
    oSALE.Key = MaxKeyUsed & "k"
    Set Add = oSALE
    Set oSALE = Nothing
End Function

Friend Sub AddSaleLine(Child As a_Sale)
    Child.SetParentExchange mParent
    Child.Sequence = mcolItems.Count + 1
    mcolItems.Add Child, Child.Key
    RaiseEvent ContainsLines(mcolItems.Count > 0)
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
    RaiseEvent ContainsLines(mcolItems.Count > 0)
End Sub

Public Function IsDirty() As Boolean
  Dim oSALE As a_Sale
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oSALE In mcolItems
      If oSALE.IsDirty Then
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
  Dim oSALE As a_Sale

  For Each oSALE In mcolItems
    oSALE.BeginEdit
  Next
  For Each oSALE In mcolDeleted
    oSALE.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_sale.BeginEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim oSALE As a_Sale

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oSALE In mcolItems
    oSALE.ApplyEdit
  Next
  For Each oSALE In mcolDeleted
    oSALE.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_sale.ApplyEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  Dim lngIndex As Long
  Dim oSALE As a_Sale

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oSALE In mcolItems
    If oSALE.IsEditing Then oSALE.CancelEdit
  Next
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oSALE = mcolDeleted(lngIndex)
    With oSALE
      If Not .IsNew Then
        .CancelEdit
        If Not .IsDeleted Then
          mcolItems.Add oSALE, oSALE.Key
          mcolDeleted.Remove lngIndex
        End If
      Else
        mcolDeleted.Remove lngIndex
      End If
    End With
    If oSALE.IsEditing Then oSALE.CancelEdit
    Set oSALE = Nothing
  Next
  
    MaxKeyUsed = mcolItems.Count
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
  Dim oSALE As a_Sale
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oSALE In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oSALE.GetSuperState
    Next
    For Each oSALE In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oSALE.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
End Function

Friend Sub SetSuperState(Buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oSALE As a_Sale
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = Buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oSALE = New a_Sale
      oSALE.SetParentExchange mParent
      oSALE.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oSALE.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oSALE.BeginEdit
      Next
      If Not oSALE.IsDeleted Then
        mcolItems.Add oSALE, oSALE.Key
      Else
        mcolDeleted.Add oSALE
      End If
      oSALE.CalculateLine  'recalculate dependant values
      Set oSALE = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex
  Set objPB = Nothing
  
  
End Sub
Public Sub SortSaleLines(pField As enSortField, pASCEND As Boolean)
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

Public Sub swap(iUP As String, iDown As String)
Dim sTemp As String
Dim iTemp As Long
    iTemp = mcolItems(iDown).Sequence
    mcolItems(iDown).Sequence = mcolItems(iUP).Sequence
    mcolItems(iUP).Sequence = iTemp
End Sub

