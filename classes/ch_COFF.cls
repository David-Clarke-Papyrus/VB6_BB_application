VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_COFF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mcolItems As Collection
Private mcolDeleted As Collection
Private mlngEditing As Long
Dim MaxKeyUsed As Long
Event CoffQtyChanged()

Private Sub mcolItems_QtyChanged()
'MsgBox "here"
End Sub

Public Function IsDirty() As Boolean
Dim oCoff As a_COFF
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oCoff In mcolItems
      If oCoff.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function
Public Sub BeginEdit()
Dim oCoff As a_COFF

  For Each oCoff In mcolItems
    oCoff.BeginEdit
  Next
  For Each oCoff In mcolDeleted
    oCoff.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub
Public Sub CancelEdit()
Dim lngIndex As Long
Dim oCoff As a_COFF

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oCoff In mcolItems
    oCoff.CancelEdit
  Next
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oCoff = mcolDeleted(lngIndex)
    With oCoff
        .CancelEdit
      If Not .IsNew Then
        If Not .IsDeleted Then
          mcolItems.Add oCoff, oCoff.Key
          mcolDeleted.Remove lngIndex
        End If
      Else
        mcolDeleted.Remove lngIndex
      End If
    End With
    If oCoff.IsEditing Then oCoff.CancelEdit
    Set oCoff = Nothing
  Next
  
  mlngEditing = mlngEditing - 1
End Sub

Public Sub ApplyEdit()
Dim oCoff As a_COFF
Dim iresult As Integer
  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oCoff In mcolItems
    oCoff.ApplyEdit
  Next
  For Each oCoff In mcolDeleted
    oCoff.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
End Sub

Public Function Add() As a_COFF
Dim oCoff As a_COFF
Dim lngEdit As Long
    If mlngEditing = 0 Then Err.Raise 445
    
    Set oCoff = New a_COFF
    With oCoff
        .Initialize Me
        For lngEdit = 1 To mlngEditing
          .BeginEdit
        Next
    End With
    MaxKeyUsed = MaxKeyUsed + 1
    oCoff.Key = MaxKeyUsed & "k"
    Set Add = oCoff
    Set oCoff = Nothing
    
End Function

Private Sub Class_Initialize()
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
  MaxKeyUsed = 0
End Sub
Public Function Item(ByVal Index As Variant) As a_COFF
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function
Public Function Count() As Long
  Count = mcolItems.Count
End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolItems.[_NewEnum]
End Function

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
Public Sub Delete()
  Dim lngIndex As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  For lngIndex = mcolItems.Count To 1 Step -1
    Remove lngIndex
  Next
End Sub
Friend Sub AddCOFF(Child As a_COFF)
    Child.Initialize Me
    mcolItems.Add Child, Child.Key
End Sub

Friend Function GetSuperState() As String
  Dim objPB As PropertyBag
  Dim lngCount As Long
Dim oCoff As a_COFF
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oCoff In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oCoff.GetSuperState
    Next
    For Each oCoff In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oCoff.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
  
End Function
Friend Sub SetSuperState(ByVal buffer As String)
Dim objPB As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngEdit As Long
Dim oCoff As a_COFF
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oCoff = New a_COFF
      oCoff.SetSuperState .ReadProperty("Item" & CStr(lngIndex)), Me
       oCoff.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oCoff.BeginEdit
      Next
      If Not oCoff.IsDeleted Then
        mcolItems.Add oCoff, oCoff.Key
      Else
        mcolDeleted.Add oCoff
      End If
      Set oCoff = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex - 1
  Set objPB = Nothing
End Sub


