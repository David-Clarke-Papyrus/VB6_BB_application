VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_CATALPI"
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

Public Function Item(ByVal Index As Variant) As a_CATALPI
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function

Public Function Add() As a_CATALPI
Dim oCATALPI As a_CATALPI
Dim lngEdit As Long
  
    If mlngEditing = 0 Then Err.Raise 445
    
    Set oCATALPI = New a_CATALPI
    With oCATALPI
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
    End With
    MaxKeyUsed = MaxKeyUsed + 1
    oCATALPI.Key = MaxKeyUsed & "k"
    Set Add = oCATALPI
    Set oCATALPI = Nothing
End Function

Friend Sub AddCATALPI(Child As a_CATALPI)
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
  Dim oCATALPI As a_CATALPI
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oCATALPI In mcolItems
      If oCATALPI.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
  Dim oCATALPI As a_CATALPI

  For Each oCATALPI In mcolItems
    oCATALPI.BeginEdit
  Next
  For Each oCATALPI In mcolDeleted
    oCATALPI.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub ApplyEdit()
  Dim oCATALPI As a_CATALPI

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oCATALPI In mcolItems
    oCATALPI.ApplyEdit
  Next
  For Each oCATALPI In mcolDeleted
    oCATALPI.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
End Sub

Public Sub CancelEdit()
  Dim lngIndex As Long
  Dim oCATALPI As a_CATALPI

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oCATALPI In mcolItems
    oCATALPI.CancelEdit
  Next
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oCATALPI = mcolDeleted(lngIndex)
    With oCATALPI
        .CancelEdit
      If Not .IsNew Then
        If Not .IsDeleted Then
          mcolItems.Add oCATALPI, oCATALPI.Key
          mcolDeleted.Remove lngIndex
        End If
      Else
        mcolDeleted.Remove lngIndex
      End If
    End With
    If oCATALPI.IsEditing Then oCATALPI.CancelEdit
    Set oCATALPI = Nothing
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
  Dim oCATALPI As a_CATALPI
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oCATALPI In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oCATALPI.GetSuperState
    Next
    For Each oCATALPI In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oCATALPI.GetSuperState
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
  Dim oCATALPI As a_CATALPI
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
        Set oCATALPI = New a_CATALPI
        oCATALPI.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
        oCATALPI.Key = lngIndex & "k"
        For lngEdit = 1 To mlngEditing
            oCATALPI.BeginEdit
        Next
        If Not oCATALPI.IsDeleted Then
            mcolItems.Add oCATALPI, oCATALPI.Key
        Else
            mcolDeleted.Add oCATALPI
        End If
        Set oCATALPI = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex - 1
  Set objPB = Nothing
End Sub

