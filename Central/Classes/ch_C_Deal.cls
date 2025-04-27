VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_Deal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mcolItems As Collection
Private mcolDeleted As Collection
Dim MaxKeyUsed As Long

Private mlngEditing As Long

Private Sub Class_Initialize()
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
  MaxKeyUsed = 0
End Sub

Public Function Count() As Long
  Count = mcolItems.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As a_Deal
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function
Public Property Get FindDealByID(PID As Long) As a_Deal
Dim oDeal As a_Deal
    For Each oDeal In mcolItems
        If oDeal.ID = PID Then
            Set FindDealByID = oDeal
            Exit For
        End If
    Next
End Property
Public Function Add() As a_Deal
  Dim oDeal As a_Deal
  Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oDeal = New a_Deal
  With oDeal
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    oDeal.Key = MaxKeyUsed & "k"
  Set Add = oDeal
  Set oDeal = Nothing
End Function

Friend Sub AddDeal(pDeal As a_Deal)

  mcolItems.Add pDeal, pDeal.Key

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
  Dim oDeal As a_Deal
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oDeal In mcolItems
      If oDeal.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
Dim oDeal As a_Deal
  
  For Each oDeal In mcolItems
    oDeal.BeginEdit
  Next
  For Each oDeal In mcolDeleted
    oDeal.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub CancelEdit()
Dim oDeal As a_Deal
Dim lngIndex As Long
  
    If mlngEditing = 0 Then Err.Raise 445
    For lngIndex = mcolItems.Count To 1 Step -1
        Set oDeal = mcolItems(lngIndex)
        If oDeal.IsNew Then
            oDeal.CancelEdit
            mcolItems.Remove lngIndex
        Else
            oDeal.CancelEdit
        End If
    Next
    For lngIndex = mcolDeleted.Count To 1 Step -1
      Set oDeal = mcolDeleted(lngIndex)
      With oDeal
          .CancelEdit
        If Not .IsNew Then
          If Not .IsDeleted Then
            mcolItems.Add oDeal, oDeal.Key
            mcolDeleted.Remove lngIndex
          End If
        Else
          mcolDeleted.Remove lngIndex
        End If
      End With
        If oDeal.IsEditing Then oDeal.CancelEdit
      Set oDeal = Nothing
    Next
  
  mlngEditing = mlngEditing - 1
  
End Sub

Public Sub ApplyEdit()
  Dim oDeal As a_Deal

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oDeal In mcolItems
    oDeal.ApplyEdit
  Next
  For Each oDeal In mcolDeleted
    oDeal.ApplyEdit
  Next
  
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
  Dim oDeal As a_Deal
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oDeal In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oDeal.GetSuperState
    Next
    For Each oDeal In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oDeal.GetSuperState
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
  Dim oDeal As a_Deal
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oDeal = New a_Deal
      oDeal.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oDeal.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oDeal.BeginEdit
      Next
      If Not oDeal.IsDeleted Then
        mcolItems.Add oDeal, oDeal.Key
      Else
        mcolDeleted.Add oDeal
      End If
      Set oDeal = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex - 1
  Set objPB = Nothing
End Sub

