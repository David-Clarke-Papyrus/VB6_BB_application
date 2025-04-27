VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_Want"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public Function Item(ByVal Index As Variant) As a_Want
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function

Public Function Add() As a_Want
Dim oWant As a_Want
Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oWant = New a_Want
  With oWant
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    oWant.Key = MaxKeyUsed & "k"
  Set Add = oWant
  Set oWant = Nothing
End Function

Friend Sub AddWant(oObj As a_Want)
    
 ' mcolItems.Add oObj
    
    oObj.Initialize Me
    mcolItems.Add oObj, oObj.ID & "k"

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
  Dim oWant As a_Want
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oWant In mcolItems
      If oWant.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
  Dim oWant As a_Want

  For Each oWant In mcolItems
    oWant.BeginEdit
  Next
  For Each oWant In mcolDeleted
    oWant.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub ApplyEdit()
  Dim oWant As a_Want

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oWant In mcolItems
    oWant.ApplyEdit
  Next
  For Each oWant In mcolDeleted
    oWant.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
End Sub

Public Sub CancelEdit()
  Dim lngIndex As Long
  Dim oWant As a_Want

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oWant In mcolItems
    oWant.CancelEdit
  Next
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oWant = mcolDeleted(lngIndex)
    With oWant
      If Not .IsNew Then
        .CancelEdit
        If Not .IsDeleted Then
          mcolItems.Add oWant, oWant.Key
          mcolDeleted.Remove lngIndex
        End If
      Else
        mcolDeleted.Remove lngIndex
      End If
    End With
    If oWant.IsEditing Then oWant.CancelEdit
    Set oWant = Nothing
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
  Dim oWant As a_Want
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oWant In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oWant.GetSuperState
    Next
    For Each oWant In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oWant.GetSuperState
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
  Dim oWant As a_Want
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oWant = New a_Want
      oWant.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
       oWant.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oWant.BeginEdit
      Next
      If Not oWant.IsDeleted Then
        mcolItems.Add oWant, oWant.Key
      Else
        mcolDeleted.Add oWant
      End If
      Set oWant = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex - 1
  Set objPB = Nothing
End Sub



