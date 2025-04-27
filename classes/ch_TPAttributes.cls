VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_TPAttributes"
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
End Sub

Public Function Count() As Long
  Count = mcolItems.Count
End Function

Public Function NewEnum() As IUnknown
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As a_Store
    If mcolItems.Count > 0 Then
        Set Item = mcolItems.Item(Index)
    End If
End Function
Public Property Get FindStoreByID(pID As Long) As a_Store
Dim oSt As a_Store
    For Each oSt In mcolItems
        If oSt.ID = pID Then
            Set FindStoreByID = oSt
            Exit For
        End If
    Next
End Property

Public Function Add() As a_Store
  Dim objStore As a_Store
  Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set objStore = New a_Store
  With objStore
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    objStore.Key = MaxKeyUsed & "k"
  Set Add = objStore
  Set objStore = Nothing
End Function

Friend Sub AddStore(pStore As a_Store)

  mcolItems.Add pStore, pStore.Key

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
  Dim objStore As a_Store
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each objStore In mcolItems
      If objStore.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
Dim oStore As a_Store
  
  For Each oStore In mcolItems
    oStore.BeginEdit
  Next
  For Each oStore In mcolDeleted
    oStore.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub CancelEdit()
Dim oStore As a_Store
Dim lngIndex As Long
  
    If mlngEditing = 0 Then Err.Raise 445
    For lngIndex = mcolItems.Count To 1 Step -1
        Set oStore = mcolItems(lngIndex)
        If oStore.IsNew Then
            oStore.CancelEdit
            mcolItems.Remove lngIndex
        Else
            oStore.CancelEdit
        End If
    Next
'    For Each objComp In mcolItems
'      objComp.CancelEdit
'      objComp.
'    Next
    For lngIndex = mcolDeleted.Count To 1 Step -1
      Set oStore = mcolDeleted(lngIndex)
      With oStore
        If Not .IsNew Then
          .CancelEdit
          If Not .IsDeleted Then
            mcolItems.Add oStore
            mcolDeleted.Remove lngIndex
          End If
        Else
          mcolDeleted.Remove lngIndex
        End If
      End With
      Set oStore = Nothing
    Next
  
  mlngEditing = mlngEditing - 1
  
End Sub

Public Sub ApplyEdit()
  Dim oStore As a_Store
  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oStore In mcolItems
    oStore.ApplyEdit
  Next
  For Each oStore In mcolDeleted
    oStore.ApplyEdit
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
  Dim oStore As a_Store
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oStore In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oStore.GetSuperState
    Next
    For Each oStore In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oStore.GetSuperState
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
  Dim oStore As a_Store
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oStore = New a_Store
      oStore.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oStore.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oStore.BeginEdit
      Next
      If Not oStore.IsDeleted Then
        mcolItems.Add oStore, oStore.Key
      Else
        mcolDeleted.Add oStore
      End If
      Set oStore = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex
  Set objPB = Nothing
End Sub




