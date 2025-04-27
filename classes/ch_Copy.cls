VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_Copy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolItems As Collection
Private mcolDeleted As Collection

Private mlngEditing As Long
Dim MaxKeyUsed As Long

Private Sub Class_Initialize()
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
  MaxKeyUsed = 0
End Sub

Public Function Count() As Long
  Count = mcolItems.Count
End Function
Public Function CountForSale() As Long
Dim i As Integer
Dim j As Integer
    i = 0
    j = 1
    Do While j <= mcolItems.Count
        If mcolItems(j).SoldDate < CDate("1980-01-01") Then
            i = i + 1
        End If
        j = j + 1
    Loop
    CountForSale = i
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As a_Copy
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function
Public Function PreviousItem(PID As Long) As a_Copy
Dim oCopy As a_Copy
Dim i As Integer

    Set PreviousItem = Nothing
    If PID = 0 Then
        i = mcolItems.Count
        If i > 0 Then
            Set PreviousItem = mcolItems(i)
        Else
            Set PreviousItem = Nothing
        End If
        Exit Function
    End If
    For i = 1 To mcolItems.Count
        If mcolItems(i).ID = PID And i > 1 Then
            Set PreviousItem = mcolItems(i - 1)
            Exit For
        End If
    Next
    
End Function
Public Property Get QtyCopiesOnHand() As Long
Dim lngCnt As Long
Dim oCopy As a_Copy
    lngCnt = 0
    For Each oCopy In mcolItems
        If oCopy.SoldDate > CDate("1990-01-01") Then
            lngCnt = lngCnt + 1
        End If
    Next
    QtyCopiesOnHand = lngCnt
End Property

Public Function FindBySerial(pSerial As Integer) As a_Copy
Dim oCopy As a_Copy
    Set FindBySerial = Nothing
    For Each oCopy In mcolItems
        If oCopy.Serial = pSerial Then
            Set FindBySerial = oCopy
            Exit For
        End If
    Next
End Function
Public Function Add() As a_Copy
Dim oCopy As a_Copy
Dim lngEdit As Long
  
    If mlngEditing = 0 Then Err.Raise 445
    
    Set oCopy = New a_Copy
    With oCopy
        .Initialize Me
        For lngEdit = 1 To mlngEditing
          .BeginEdit
        Next
    End With
    MaxKeyUsed = MaxKeyUsed + 1
    oCopy.Key = MaxKeyUsed & "k"
    Set Add = oCopy
    Set oCopy = Nothing
    End Function

Friend Sub AddCopy(Child As a_Copy)
    Child.Initialize Me
    mcolItems.Add Child, Child.Key
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
  Dim oCopy As a_Copy
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oCopy In mcolItems
      If oCopy.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
  Dim oCopy As a_Copy

  For Each oCopy In mcolItems
    oCopy.BeginEdit
  Next
  For Each oCopy In mcolDeleted
    oCopy.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub ApplyEdit()
  Dim oCopy As a_Copy

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oCopy In mcolItems
    oCopy.ApplyEdit
  Next
  For Each oCopy In mcolDeleted
    oCopy.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
End Sub

Public Sub CancelEdit()
  Dim lngIndex As Long
  Dim oCopy As a_Copy

  If mlngEditing = 0 Then Err.Raise 445
  
'  For Each oCopy In mcolItems
'    oCopy.CancelEdit
'  Next
'
'
    For lngIndex = mcolItems.Count To 1 Step -1
        Set oCopy = mcolItems(lngIndex)
        If oCopy.IsNew Then
            oCopy.CancelEdit
            mcolItems.Remove lngIndex
        Else
            oCopy.CancelEdit
        End If
    Next
  
  
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oCopy = mcolDeleted(lngIndex)
    With oCopy
        .CancelEdit
      If Not .IsNew Then
        If Not .IsDeleted Then
          mcolItems.Add oCopy, oCopy.Key
          mcolDeleted.Remove lngIndex
        End If
      Else
        mcolDeleted.Remove lngIndex
      End If
    End With
    If oCopy.IsEditing Then oCopy.CancelEdit
    Set oCopy = Nothing
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
  Dim oCopy As a_Copy
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oCopy In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oCopy.GetSuperState
    Next
    For Each oCopy In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oCopy.GetSuperState
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
  Dim oCopy As a_Copy
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oCopy = New a_Copy
      oCopy.SetSuperState .ReadProperty("Item" & CStr(lngIndex)), Me
       oCopy.Key = lngIndex & "k"
     For lngEdit = 1 To mlngEditing
        oCopy.BeginEdit
      Next
      If Not oCopy.IsDeleted Then
        mcolItems.Add oCopy, oCopy.Key   '%%%
      Else
        mcolDeleted.Add oCopy
      End If
      Set oCopy = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex - 1
  Set objPB = Nothing
End Sub

