VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_PInstCATAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolItems As Collection
Private mcolDeleted As Collection

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

Public Function Item(ByVal Index As Variant) As a_PICATAL
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function

Public Function Add() As a_PICATAL
Dim oPICatal As a_PICATAL
Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oPICatal = New a_PICATAL
  With oPICatal
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
  Set Add = oPICatal
  Set oPICatal = Nothing
End Function

Friend Sub AddPICATAL(oObj As a_PICATAL)
    
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
  Dim oPICatal As a_PICATAL
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oPICatal In mcolItems
      If oPICatal.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
  Dim oPICatal As a_PICATAL

  For Each oPICatal In mcolItems
    oPICatal.BeginEdit
  Next
  For Each oPICatal In mcolDeleted
    oPICatal.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub ApplyEdit()
  Dim oPICatal As a_PICATAL

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oPICatal In mcolItems
    oPICatal.ApplyEdit
  Next
  For Each oPICatal In mcolDeleted
    oPICatal.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
End Sub

Public Sub CancelEdit()
  Dim lngIndex As Long
  Dim oPICatal As a_PICATAL

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oPICatal In mcolItems
    oPICatal.CancelEdit
  Next
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oPICatal = mcolDeleted(lngIndex)
    With oPICatal
      If Not .IsNew Then
        .CancelEdit
        If Not .IsDeleted Then
          mcolItems.Add oPICatal
          mcolDeleted.Remove lngIndex
        End If
      Else
        mcolDeleted.Remove lngIndex
      End If
    End With
    Set oPICatal = Nothing
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
  Dim oPICatal As a_PICATAL
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oPICatal In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oPICatal.GetSuperState
    Next
    For Each oPICatal In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oPICatal.GetSuperState
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
  Dim oPICatal As a_PICATAL
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = Buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oPICatal = New a_PICATAL
      oPICatal.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      For lngEdit = 1 To mlngEditing
        oPICatal.BeginEdit
      Next
      If Not oPICatal.IsDeleted Then
        mcolItems.Add oPICatal, oPICatal.ID & "k"
      Else
        mcolDeleted.Add oPICatal
      End If
      Set oPICatal = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub


