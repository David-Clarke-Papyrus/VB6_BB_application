VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_Addresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event FirstAddressAdded()
Private mcolItems As Collection
Private mcolDeleted As Collection
Dim MaxKeyUsed As Long
Private mlngEditing As Long
Private mCustomer As a_Customer
Public Sub SetCustomerParent(pParent As a_Customer)
    Set mCustomer = pParent
End Sub
Public Property Get Customer() As a_Customer
    Set Customer = mCustomer
End Property
Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
  MaxKeyUsed = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.Class_Initialize"
End Sub

Public Function Count() As Long
    On Error GoTo errHandler
  Count = mcolItems.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.Count"
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
  Set NewEnum = mcolItems.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.NewEnum"
End Function

Public Function Item(ByVal Index As Variant) As a_Address
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
  Set Item = mcolItems.Item(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.Item(Index)", Index
End Function

Public Function Add() As a_Address
    On Error GoTo errHandler
Dim oAdd As a_Address
Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oAdd = New a_Address
  With oAdd
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    oAdd.Key = MaxKeyUsed & "k"
  Set Add = oAdd
  Set oAdd = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.Add"
End Function

Friend Sub AddAddress(Child As a_Address)
    On Error GoTo errHandler

    mcolItems.Add Child, Child.Key
    If mcolItems.Count = 1 Then
        RaiseEvent FirstAddressAdded
        Child.BillTo = True
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.AddAddress(Child)", Child
End Sub

Public Sub Remove(ByVal Index As Variant)
    On Error GoTo errHandler
  If mlngEditing = 0 Then Err.Raise 445

  With mcolItems(Index)
    .BeginEdit
    .Delete
    .ApplyEdit
  End With
  mcolDeleted.Add mcolItems(Index)
  mcolItems.Remove Index
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.Remove(Index)", Index
End Sub

Public Function IsDirty() As Boolean
    On Error GoTo errHandler
  Dim oAdd As a_Address
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oAdd In mcolItems
      If oAdd.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.IsDirty"
End Function

Public Sub BeginEdit()
    On Error GoTo errHandler
  Dim oAdd As a_Address

  For Each oAdd In mcolItems
    oAdd.BeginEdit
  Next
  For Each oAdd In mcolDeleted
    oAdd.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.BeginEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim oAdd As a_Address

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oAdd In mcolItems
    oAdd.ApplyEdit
  Next
  For Each oAdd In mcolDeleted
    oAdd.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.ApplyEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  Dim lngIndex As Long
  Dim oAdd As a_Address

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oAdd In mcolItems
    oAdd.CancelEdit
  Next
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oAdd = mcolDeleted(lngIndex)
    With oAdd
      If Not .IsNew Then
        .CancelEdit
        If Not .IsDeleted Then
          mcolItems.Add oAdd
          mcolDeleted.Remove lngIndex
        End If
      Else
        mcolDeleted.Remove lngIndex
      End If
    End With
    If oAdd.IsEditing Then oAdd.CancelEdit
    Set oAdd = Nothing
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.CancelEdit"
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
    ErrorIn "ch_Addresses.Delete"
End Sub

Friend Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oAdd As a_Address
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oAdd In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oAdd.GetSuperState
    Next
    For Each oAdd In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oAdd.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oAdd As a_Address
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oAdd = New a_Address
      oAdd.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oAdd.Key = lngIndex & "k"
      oAdd.SetParent Me
      For lngEdit = 1 To mlngEditing
        oAdd.BeginEdit
      Next
      If Not oAdd.IsDeleted Then
        mcolItems.Add oAdd, oAdd.Key
      Else
        mcolDeleted.Add oAdd
      End If
      Set oAdd = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex - 1
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.SetSuperState(buffer)", buffer
End Sub
Public Sub SetAppro(pIdx As Integer)
    On Error GoTo errHandler
Dim oAdd As a_Address
    For Each oAdd In mcolItems
        oAdd.Appro = False
    Next
    mcolItems(pIdx).Appro = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.SetAppro(pIdx)", pIdx
End Sub
Public Sub SetBillTo(pIdx As Integer)
    On Error GoTo errHandler
Dim oAdd As a_Address
    For Each oAdd In mcolItems
        oAdd.BillTo = False
    Next
    mcolItems(pIdx).BillTo = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.SetBillTo(pIdx)", pIdx
End Sub
Public Sub SetDelTo(pIdx As Integer)
    On Error GoTo errHandler
Dim oAdd As a_Address
    For Each oAdd In mcolItems
        oAdd.DelTo = False
    Next
    mcolItems(pIdx).DelTo = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.SetDelTo(pIdx)", pIdx
End Sub
Public Sub SetOrderTo(pIdx As Integer)
    On Error GoTo errHandler
Dim oAdd As a_Address
    For Each oAdd In mcolItems
        oAdd.OrderTo = False
    Next
    mcolItems(pIdx).OrderTo = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.SetOrderTo(pIdx)", pIdx
End Sub
Public Function FindByDescription(pDescription As String) As a_Address
    On Error GoTo errHandler
Dim oAdd As a_Address
    For Each oAdd In mcolItems
        If oAdd.Description = pDescription Then
            Set FindByDescription = oAdd
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.FindByDescription(pDescription)", pDescription
End Function
Public Function FindByID(PID As Long) As a_Address
    On Error GoTo errHandler
Dim oAdd As a_Address
    For Each oAdd In mcolItems
        If oAdd.ID = PID Then
            Set FindByID = oAdd
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Addresses.FindByID(pID)", PID
End Function
Public Property Get DefaultAddress() As a_Address
Dim oAdd As a_Address
    For Each oAdd In mcolItems
        If oAdd.BillTo = True Then
            Set DefaultAddress = oAdd
            Exit For
        End If
    Next
    Exit Function

End Property
Public Property Get BillTOAddress() As a_Address

End Property
Public Property Get DelToAddress() As a_Address

End Property
Public Property Get OrderToAddress() As a_Address

End Property

