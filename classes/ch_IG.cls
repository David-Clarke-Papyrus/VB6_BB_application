VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_IG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mcolKeys As Collection
Private mcolItems As Collection
Private mcolDeleted As Collection
Dim bIsLoyaltyClubMember As Boolean
Dim bIsLitLunch As Boolean
Dim bIsLaunch As Boolean
Dim bIsPromotion As Boolean
Dim bIsSale As Boolean
Dim MaxKeyUsed As Long
Private mlngEditing As Long

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolKeys = New Collection
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
  MaxKeyUsed = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_IG.Class_Initialize"
End Sub

Public Function Count() As Long
    On Error GoTo errHandler
  Count = mcolItems.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_IG.Count"
End Function
Public Function Key(ByVal text As String) As String
On Error GoTo ERR_Handler
    Key = (mcolKeys.Item(text))
EXIT_Handler:
    Exit Function
ERR_Handler:
    Key = 0
    GoTo EXIT_Handler
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
  Set NewEnum = mcolItems.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_IG.NewEnum"
End Function

Public Function Item(ByVal Index As Variant) As a_IG
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
  Set Item = mcolItems.Item(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_IG.Item(Index)", Index
End Function
Public Sub BeginEdit()
    On Error GoTo errHandler
  Dim oMember As a_IG

  For Each oMember In mcolItems
    oMember.BeginEdit
  Next
  For Each oMember In mcolDeleted
    oMember.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_IG.BeginEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim oMember As a_IG

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oMember In mcolItems
    oMember.ApplyEdit
  Next
  For Each oMember In mcolDeleted
    oMember.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_IG.ApplyEdit"
    Exit Sub
    Resume
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  Dim lngIndex As Long
  Dim oMember As a_IG

  If mlngEditing = 0 Then Err.Raise 445
  
    For lngIndex = mcolItems.Count To 1 Step -1
        Set oMember = mcolItems(lngIndex)
        If oMember.IsNew Then
            oMember.CancelEdit
            mcolItems.Remove lngIndex
        Else
            oMember.CancelEdit
        End If
    Next
  
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oMember = mcolDeleted(lngIndex)
    With oMember
      If Not .IsNew Then
        .CancelEdit
        If Not .IsDeleted Then
          mcolItems.Add oMember, oMember.Key
          mcolDeleted.Remove lngIndex
        End If
      Else
        mcolDeleted.Remove lngIndex
      End If
    End With
    If oMember.IsEditing Then oMember.CancelEdit
    Set oMember = Nothing
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_IG.CancelEdit"
End Sub

Public Sub Delete()
    On Error GoTo errHandler
  Dim lngIndex As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  For lngIndex = mcolItems.Count To 1 Step -1
    Remove lngIndex
  Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_IG.Delete"
End Sub

Friend Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oMember As a_IG
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oMember In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oMember.GetSuperState
    Next
    For Each oMember In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oMember.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_IG.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oMember As a_IG
  
  Set objPB = New PropertyBag
  Set mcolKeys = Nothing
  Set mcolKeys = New Collection
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    bIsLoyaltyClubMember = False
    For lngIndex = 1 To .ReadProperty("Count")
      Set oMember = New a_IG
      oMember.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oMember.Key = lngIndex & "k"
      mcolKeys.Add oMember.Key, oMember.Description
      For lngEdit = 1 To mlngEditing
        oMember.BeginEdit
      Next
      If Not oMember.IsDeleted Then
        mcolItems.Add oMember, oMember.Key
      Else
        mcolDeleted.Add oMember
      End If
      If oMember.IsLoyaltyClubMember Then bIsLoyaltyClubMember = True
      If oMember.IsLaunch Then bIsLaunch = True
      If oMember.IsSale Then bIsSale = True
      If oMember.ISPromotion Then bIsPromotion = True
      If oMember.IsLitLunch Then bIsLitLunch = True
      Set oMember = Nothing
    Next
        MaxKeyUsed = .ReadProperty("Count")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_IG.SetSuperState(buffer)", buffer
End Sub
Public Sub Remove(ByVal Index As Variant)
    On Error GoTo errHandler
  If mlngEditing = 0 Then Err.Raise 445
    If Index = 0 Then Exit Sub
  With mcolItems(Index)
    .BeginEdit
    .Delete
    .ApplyEdit
  End With
  mcolDeleted.Add mcolItems(Index)
  mcolKeys.Remove mcolItems(Index).Description
  mcolItems.Remove Index
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_IG.Remove(Index)", Index
End Sub

Public Function IsDirty() As Boolean
    On Error GoTo errHandler
  Dim oMember As a_IG
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oMember In mcolItems
      If oMember.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_IG.IsDirty"
End Function
Public Function Add() As a_IG
    On Error GoTo errHandler
  Dim oMember As a_IG
  Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oMember = New a_IG
  With oMember
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    oMember.Key = MaxKeyUsed & "k"
  Set Add = oMember
  Set oMember = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_IG.Add"
End Function

Friend Sub AddMember(oMember As a_IG)
    On Error GoTo errHandler

    mcolItems.Add oMember, oMember.Key
    mcolKeys.Add oMember.Key, oMember.Description
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_IG.AddMember(oMember)", oMember
End Sub

Public Function IsALoyaltyMember() As Boolean
    IsALoyaltyMember = bIsLoyaltyClubMember
End Function
Public Function IsaLaunchMember() As Boolean
    IsaLaunchMember = bIsLaunch
End Function
Public Function IsALitLunchMember() As Boolean
    IsALitLunchMember = bIsLitLunch
End Function
Public Function IsASaleMember() As Boolean
    IsASaleMember = bIsSale
End Function
Public Function IsAPromotionMember() As Boolean
    IsAPromotionMember = bIsPromotion
End Function

