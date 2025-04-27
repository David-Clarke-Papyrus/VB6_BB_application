VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_ProductSection"
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
Event PriorityCategory(ID As Long)
Public Function Mark(pPSECKey As String)
 Dim oPS As a_ProductSection
    
    For Each oPS In mcolItems
        If oPS.Key = pPSECKey Then
            oPS.Priority = 99
            RaiseEvent PriorityCategory(oPS.SECID)
        Else
            oPS.Priority = 0
        End If
    Next
    
End Function
Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolKeys = New Collection
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
  MaxKeyUsed = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_ProductSection.Class_Initialize"
End Sub

Public Function Count() As Long
    On Error GoTo errHandler
  Count = mcolItems.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_ProductSection.Count"
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
    ErrorIn "ch_ProductSection.NewEnum"
End Function

Public Function Item(ByVal Index As Variant) As a_ProductSection
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
  Set Item = mcolItems.Item(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_ProductSection.Item(Index)", Index
End Function
Public Sub BeginEdit()
    On Error GoTo errHandler
  Dim oMember As a_ProductSection

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
    ErrorIn "ch_ProductSection.BeginEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim oMember As a_ProductSection

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
    ErrorIn "ch_ProductSection.ApplyEdit"
    Exit Sub
    Resume
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  Dim lngIndex As Long
  Dim oMember As a_ProductSection

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
    ErrorIn "ch_ProductSection.CancelEdit"
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
    ErrorIn "ch_ProductSection.Delete"
End Sub

Friend Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oMember As a_ProductSection
  
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
    ErrorIn "ch_ProductSection.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oMember As a_ProductSection
  
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
      Set oMember = New a_ProductSection
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
      Set oMember = Nothing
    Next
  MaxKeyUsed = .ReadProperty("Count")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_ProductSection.SetSuperState(buffer)", buffer
End Sub
Public Function Remove(ByVal Index As Variant) As Boolean
    On Error GoTo errHandler
  If mlngEditing = 0 Then Err.Raise 445
    Remove = False
    If Index = 0 Then Exit Function
    If mcolItems(Index).Priority = 99 Then
        Remove = False
        Exit Function
    End If
  With mcolItems(Index)
    .BeginEdit
    .Delete
    Remove = True
    .ApplyEdit
  End With
    mcolDeleted.Add mcolItems(Index)
    mcolKeys.Remove mcolItems(Index).Description
    mcolItems.Remove Index
    Remove = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_ProductSection.Remove(Index)", Index
End Function

Public Function IsDirty() As Boolean
    On Error GoTo errHandler
  Dim oMember As a_ProductSection
  
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
    ErrorIn "ch_ProductSection.IsDirty"
End Function
Public Function Add() As a_ProductSection
    On Error GoTo errHandler
  Dim oMember As a_ProductSection
  Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oMember = New a_ProductSection
  With oMember
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    oMember.Key = MaxKeyUsed & "k"
    'Priority cannot be greater that number of members
    oMember.Priority = Me.Count + 1
  Set Add = oMember
  Set oMember = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_ProductSection.Add"
End Function

Friend Sub AddMember(oMember As a_ProductSection)
    On Error GoTo errHandler

    On Error Resume Next
        mcolKeys.Add oMember.Key, oMember.Description
        If Err = 0 Then
            mcolItems.Add oMember, oMember.Key
        End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_ProductSection.AddMember(oMember)", oMember
End Sub

Public Property Get SectionsAsList() As String
Dim oPS As a_ProductSection
Dim str As String

    For Each oPS In Me
        If Not oPS.IsDeleted Then
            str = str & IIf(Len(str) > 0, ",", "") & oPS.Description
        End If
    Next
    SectionsAsList = str
End Property
