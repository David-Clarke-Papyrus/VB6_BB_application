VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_CATALP"
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

Public Function Item(ByVal Index As Variant) As a_CATALP
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function

Public Function Add() As a_CATALP
Dim oCATALP As a_CATALP
Dim lngEdit As Long
  
    If mlngEditing = 0 Then Err.Raise 445
    
    Set oCATALP = New a_CATALP
    With oCATALP
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
    End With
    MaxKeyUsed = MaxKeyUsed + 1
    oCATALP.Key = MaxKeyUsed & "k"
    Set Add = oCATALP
    Set oCATALP = Nothing
End Function

Friend Sub AddCATALP(Child As a_CATALP)
On Error GoTo H
 ' mcolItems.Add oObj
    
    Child.Initialize Me
    mcolItems.Add Child, Child.Key
Exit Sub
H:
    If Err = 457 Then
        Resume Next
    Else
        MsgBox Error
    End If
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
  Dim oCATALP As a_CATALP
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oCATALP In mcolItems
      If oCATALP.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
  Dim oCATALP As a_CATALP

  For Each oCATALP In mcolItems
    oCATALP.BeginEdit
  Next
  For Each oCATALP In mcolDeleted
    oCATALP.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub ApplyEdit()
  Dim oCATALP As a_CATALP

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oCATALP In mcolItems
    oCATALP.ApplyEdit
  Next
  For Each oCATALP In mcolDeleted
    oCATALP.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
End Sub

Public Sub CancelEdit()
  Dim lngIndex As Long
  Dim oCATALP As a_CATALP

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oCATALP In mcolItems
    oCATALP.CancelEdit
  Next
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oCATALP = mcolDeleted(lngIndex)
    With oCATALP
        .CancelEdit
      If Not .IsNew Then
        If Not .IsDeleted Then
          mcolItems.Add oCATALP, oCATALP.Key
          mcolDeleted.Remove lngIndex
        End If
      Else
        mcolDeleted.Remove lngIndex
      End If
    End With
    If oCATALP.IsEditing Then oCATALP.CancelEdit
    Set oCATALP = Nothing
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
  Dim oCATALP As a_CATALP
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oCATALP In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oCATALP.GetSuperState
    Next
    For Each oCATALP In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oCATALP.GetSuperState
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
  Dim oCATALP As a_CATALP
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oCATALP = New a_CATALP
        oCATALP.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        oCATALP.Key = lngIndex & "k"
        For lngEdit = 1 To mlngEditing
            oCATALP.BeginEdit
        Next
        If Not oCATALP.IsDeleted Then
            mcolItems.Add oCATALP, oCATALP.Key
        Else
            mcolDeleted.Add oCATALP
        End If
        Set oCATALP = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex - 1
  Set objPB = Nothing
End Sub


