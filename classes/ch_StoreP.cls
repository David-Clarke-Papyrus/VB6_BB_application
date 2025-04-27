VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_StoreP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mcolItems As Collection
Private mcolDeleted As Collection

Private mlngEditing As Long

'Public Sub Load(pDefaultCompany As a_Company, pDefaultCOMPID As Long)
'  Fetch pDefaultCompany, pDefaultCOMPID
'End Sub

Private Sub Class_Initialize()
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
End Sub

Public Function Count() As Long
  Count = mcolItems.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As a_StoreP
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function
Public Function Add() As a_StoreP
  Dim objStoreP As a_StoreP
  Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set objStoreP = New a_StoreP
  With objStoreP
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
  Set Add = objStoreP
  Set objStoreP = Nothing
End Function

Friend Sub AddStoreP(pStoreP As a_StoreP)

  mcolItems.Add pStoreP, pStoreP.STID & "k"

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
  Dim objComp As a_Company
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each objComp In mcolItems
      If objComp.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
Dim oStoreP As a_StoreP
  
  For Each oStoreP In mcolItems
    oStoreP.BeginEdit
  Next
  For Each oStoreP In mcolDeleted
    oStoreP.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub CancelEdit()
Dim oStoreP As a_StoreP
Dim lngIndex As Long
  
    If mlngEditing = 0 Then Err.Raise 445
    For lngIndex = mcolItems.Count To 1 Step -1
        Set oStoreP = mcolItems(lngIndex)
        If oStoreP.IsNew Then
            oStoreP.CancelEdit
            mcolItems.Remove lngIndex
        Else
            oStoreP.CancelEdit
        End If
    Next
'    For Each objComp In mcolItems
'      objComp.CancelEdit
'      objComp.
'    Next
    For lngIndex = mcolDeleted.Count To 1 Step -1
      Set oStoreP = mcolDeleted(lngIndex)
      With oStoreP
        If Not .IsNew Then
          .CancelEdit
          If Not .IsDeleted Then
            mcolItems.Add oStoreP
            mcolDeleted.Remove lngIndex
          End If
        Else
          mcolDeleted.Remove lngIndex
        End If
      End With
      If oStoreP.IsEditing Then oStoreP.CancelEdit
      Set oStoreP = Nothing
    Next
  
  mlngEditing = mlngEditing - 1
  
End Sub

Public Sub ApplyEdit()
  Dim oStoreP As a_StoreP
  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oStoreP In mcolItems
    oStoreP.ApplyEdit
  Next
  For Each oStoreP In mcolDeleted
    oStoreP.ApplyEdit
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
  Dim oStoreP As a_StoreP
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oStoreP In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oStoreP.GetSuperState
    Next
    For Each oStoreP In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oStoreP.GetSuperState
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
  Dim oStoreP As a_StoreP
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oStoreP = New a_StoreP
      oStoreP.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      For lngEdit = 1 To mlngEditing
        oStoreP.BeginEdit
      Next
      If Not oStoreP.IsDeleted Then
        mcolItems.Add oStoreP, oStoreP.STID & "k"
      Else
        mcolDeleted.Add oStoreP
      End If
      Set oStoreP = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub


