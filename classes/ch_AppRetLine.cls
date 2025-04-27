VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_APPRL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mParent As a_APPR
Private mcolItems As Collection
Private colDeleted As Collection
Private mlngEditing As Long
Dim MaxKeyUsed As Long
Dim WithEvents oAPPRL As a_APPRL
Attribute oAPPRL.VB_VarHelpID = -1
Event rowsChange()

Public Sub component(pParent As a_APPR)
    Set mParent = pParent
End Sub

Private Sub Class_Initialize()
  Set mcolItems = New Collection
  Set colDeleted = New Collection
  MaxKeyUsed = 0
End Sub

Public Function Item(ByVal Index As Variant) As a_APPRL
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems(Index)
End Function

Public Function Count() As Long
  Count = mcolItems.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Friend Function GetSuperState() As String
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oAPPRL As a_APPRL
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + colDeleted.Count
    For Each oAPPRL In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oAPPRL.GetSuperState
    Next
    For Each oAPPRL In colDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oAPPRL.GetSuperState
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
  Dim oAPPRL As a_APPRL
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set colDeleted = Nothing
  Set colDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oAPPRL = New a_APPRL
      oAPPRL.SetParentAPPR mParent
      oAPPRL.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oAPPRL.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oAPPRL.BeginEdit
      Next
      If Not oAPPRL.IsDeleted Then
        mcolItems.Add oAPPRL, oAPPRL.Key
      Else
        colDeleted.Add oAPPRL
      End If
      Set oAPPRL = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex - 1
  Set objPB = Nothing
End Sub

Public Function IsDirty() As Boolean
  Dim oAPPRL As a_APPRL
  
  If colDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oAPPRL In mcolItems
      If oAPPRL.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
  Dim oAL As a_APPRL

  For Each oAL In mcolItems
    oAL.BeginEdit
  Next
  For Each oAL In colDeleted
    oAL.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub ApplyEdit()
  Dim oAL As a_APPRL

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oAL In mcolItems
    oAL.ApplyEdit
  Next
  For Each oAL In colDeleted
    oAL.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
End Sub

Public Sub CancelEdit()
  Dim lngIndex As Long
  Dim oAL As a_APPRL

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oAL In mcolItems
    oAL.CancelEdit
  Next
  For lngIndex = colDeleted.Count To 1 Step -1
    Set oAL = colDeleted(lngIndex)
    With oAL
      .CancelEdit
      If Not .IsNew Then
        If Not .IsDeleted Then
          mcolItems.Add oAL
          colDeleted.Remove lngIndex
        End If
      Else
        colDeleted.Remove lngIndex
      End If
    End With
    If oAL.IsEditing Then oAL.CancelEdit
    Set oAL = Nothing
  Next
  
  mlngEditing = mlngEditing - 1
End Sub

Public Function Add() As a_APPRL
Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oAPPRL = New a_APPRL
  With oAPPRL
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    oAPPRL.Key = MaxKeyUsed & "k"
    oAPPRL.SetParentAPPR mParent
  Set Add = oAPPRL
  Set oAPPRL = Nothing
End Function

Friend Sub AddApproReturnLine(Child As a_APPRL)
    Child.SetParentAPPR mParent
    mcolItems.Add Child, Child.Key
    RaiseEvent rowsChange
End Sub

Public Sub Remove(ByVal Index As Variant)
  If mlngEditing = 0 Then Err.Raise 445

  With mcolItems(Index)
    .BeginEdit
    .Delete
    .ApplyEdit
  End With
  colDeleted.Add mcolItems(Index)
  mcolItems.Remove Index
End Sub
Function FindLineByID(PID As Long) As a_APPRL
Dim oAPPRL As a_APPRL
    
    For Each oAPPRL In mcolItems
        If oAPPRL.APPRLID = PID Then
            Set FindLineByID = oAPPRL
            Exit For
        End If
    Next
End Function


