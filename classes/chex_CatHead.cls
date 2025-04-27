VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "chex_Cathead"
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
End Sub

Public Function Count() As Long
  Count = mcolItems.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As a_cathead
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function
Public Function Add() As a_cathead
  Dim obj As a_cathead
  Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set obj = New a_cathead
  With obj
    .Manager = Me
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    obj.Key = MaxKeyUsed & "k"
  Set Add = obj
  Set obj = Nothing
End Function
Public Property Get IsEditing() As Boolean
    IsEditing = (mlngEditing > 0)
End Property

Friend Sub AddCathead(Child As a_cathead)
        mcolItems.Add Child, Child.Key
End Sub
Friend Sub ValidateChild(Child As a_cathead, por As String)
Dim obj As a_cathead
    por = ""
    For Each obj In mcolItems
        If obj.Description = Child.Description Or obj.SortTag = Child.SortTag Then
            por = "There are duplicate descriptions or sorttags"
            Exit For
        End If
    Next
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
  Dim obj As a_cathead
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each obj In mcolItems
      If obj.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
Dim obj As a_cathead
  
  For Each obj In mcolItems
    obj.BeginEdit
  Next
  For Each obj In mcolDeleted
    obj.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
  
End Sub

Public Sub CancelEdit()
Dim obj As a_cathead
Dim lngIndex As Long
    If mlngEditing = 0 Then Err.Raise 445
    For lngIndex = mcolItems.Count To 1 Step -1
        Set obj = mcolItems(lngIndex)
        If obj.IsNew Then
            obj.CancelEdit
            mcolItems.Remove lngIndex
        Else
            obj.CancelEdit
        End If
    Next
    For lngIndex = mcolDeleted.Count To 1 Step -1
      Set obj = mcolDeleted(lngIndex)
      With obj
        If Not .IsNew Then
          .CancelEdit
          If Not .IsDeleted Then
            mcolItems.Add obj
            mcolDeleted.Remove lngIndex
          End If
        Else
          mcolDeleted.Remove lngIndex
        End If
      End With
      Set obj = Nothing
    Next
  
  mlngEditing = mlngEditing - 1
  
End Sub

Public Sub ApplyEdit(por As String)
Dim obj As a_cathead
Dim stror As String
  If mlngEditing = 0 Then Err.Raise 445
  
  For Each obj In mcolItems
    obj.ApplyEdit
  Next
  For Each obj In mcolDeleted
    obj.ApplyEdit
  Next
  Save por
  mlngEditing = mlngEditing - 1
End Sub
Public Sub Delete()
  Dim lngIndex As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  For lngIndex = mcolItems.Count To 1 Step -1
    Remove lngIndex
  Next
End Sub
Public Function Load()
Dim objPersist As a_cathead_P
    Set objPersist = New a_cathead_P
    SetSuperState objPersist.Fetch()
End Function
Private Function Save(por As String)
Dim strBuffer As String
Dim objPersist As a_cathead_P
    Set objPersist = New a_cathead_P
    strBuffer = objPersist.Save(GetSuperState)
    If strBuffer = "" Then
        por = "Duplicate values in description or sorttag"
 '   Else
 '       SetSuperState strBuffer
    End If

End Function
Friend Function GetSuperState() As String
  Dim oPB As PropertyBag
  Dim lngCount As Long
  Dim obj As a_cathead
  Set oPB = New PropertyBag
  With oPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each obj In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), obj.GetSuperState
    Next
    For Each obj In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), obj.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set oPB = Nothing
End Function

Friend Sub SetSuperState(buffer As String)
  Dim oPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim obj As a_cathead
  
  Set oPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With oPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set obj = New a_cathead
      obj.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
       obj.Key = lngIndex & "k"
       obj.Manager = Me
      For lngEdit = 1 To mlngEditing
        obj.BeginEdit
      Next
      If Not obj.IsDeleted Then
        mcolItems.Add obj, obj.Key
      Else
        mcolDeleted.Add obj
      End If
      Set obj = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex - 1
  Set oPB = Nothing
End Sub






