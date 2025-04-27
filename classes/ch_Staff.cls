VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_Staff"
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

Public Function Item(ByVal Index As Variant) As a_Staff
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function
Public Property Get FindStaffByID(PID As Long) As a_Staff
Dim oSt As a_Staff
    For Each oSt In mcolItems
        If oSt.ID = PID Then
            Set FindStaffByID = oSt
            Exit For
        End If
    Next
End Property
Public Property Get FindStaffByName(pName As String) As a_Staff
Dim oSt As a_Staff
    For Each oSt In mcolItems
        If oSt.StaffName = pName Then
            Set FindStaffByName = oSt
            Exit For
        End If
    Next
End Property

Public Function Add() As a_Staff
  Dim oSt As a_Staff
  Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oSt = New a_Staff
  With oSt
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
    MaxKeyUsed = MaxKeyUsed + 1
    oSt.Key = MaxKeyUsed & "k"
  End With
  Set Add = oSt
  Set oSt = Nothing
End Function

Friend Sub AddStaff(pST As a_Staff)

  mcolItems.Add pST, pST.Key

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
  Dim oSt As a_Staff
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oSt In mcolItems
      If oSt.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
Dim oSt As a_Staff
  
  For Each oSt In mcolItems
    oSt.BeginEdit
  Next
  For Each oSt In mcolDeleted
    oSt.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub CancelEdit()
Dim oSt As a_Staff
Dim lngIndex As Long
  
    If mlngEditing = 0 Then Err.Raise 445
    For lngIndex = mcolItems.Count To 1 Step -1
        Set oSt = mcolItems(lngIndex)
        If oSt.IsNew Then
            oSt.CancelEdit
            mcolItems.Remove lngIndex
        Else
            oSt.CancelEdit
        End If
    Next
    For lngIndex = mcolDeleted.Count To 1 Step -1
      Set oSt = mcolDeleted(lngIndex)
      With oSt
        If Not .IsNew Then
          .CancelEdit
          If Not .IsDeleted Then
            mcolItems.Add oSt, oSt.Key
            mcolDeleted.Remove lngIndex
          End If
        Else
          mcolDeleted.Remove lngIndex
        End If
      End With
    If oSt.IsEditing Then oSt.CancelEdit
      Set oSt = Nothing
    Next
  
  mlngEditing = mlngEditing - 1
  
End Sub

Public Sub ApplyEdit()
  Dim oSt As a_Staff

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oSt In mcolItems
    oSt.ApplyEdit
  Next
  For Each oSt In mcolDeleted
    oSt.ApplyEdit
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
Public Function GetLevel(pPWD As String, Optional pName, Optional PID) As Long
Dim oStaff As a_Staff
    GetLevel = 0
    For Each oStaff In mcolItems
        If UCase(oStaff.Shortname & oStaff.Password) = UCase(pPWD) And pPWD > "" Then
            GetLevel = 99
            If Not IsMissing(pName) Then
                pName = oStaff.StaffName
            End If
            If Not IsMissing(PID) Then
                PID = oStaff.ID
            End If
        End If
    Next
End Function
Public Function IsSupervisor(pPWD As String, Optional pName, Optional PID) As Long
Dim oStaff As a_Staff
    IsSupervisor = 0
    For Each oStaff In mcolItems
        If UCase(oStaff.Shortname & oStaff.Password) = UCase(pPWD) And pPWD > "" And oStaff.IsActive Then
            IsSupervisor = oStaff.IsSupervisor
            If Not IsMissing(pName) Then
                pName = oStaff.StaffName
            End If
            If Not IsMissing(PID) Then
                PID = oStaff.ID
            End If
            Exit For
        End If
    Next
End Function
Public Function IsSecurityOK(iPos As enumSecurityNode, pSignature As String, Optional pName, Optional PID, Optional Fullsignature) As Boolean
Dim oStaff As a_Staff
Dim bOK As Boolean

    IsSecurityOK = False
    For Each oStaff In mcolItems
        If UCase(UCase(oStaff.Shortname & oStaff.Password)) = UCase(pSignature) And pSignature > "" Then
            bOK = oStaff.GetRole(iPos)
            If Not IsMissing(pName) Then
                pName = oStaff.StaffName
                Fullsignature = oStaff.Signature
            End If
            If Not IsMissing(PID) Then
                PID = oStaff.ID
            End If
            If bOK Then
                IsSecurityOK = True
                Exit For
            End If
        End If
    Next
End Function

Public Function IsLastSupervisor(pstaff As a_Staff) As Boolean
Dim oStaff As a_Staff
Dim SupervisorCount As Long
Dim bTmp As Boolean
        SupervisorCount = 0
        bTmp = False
        For Each oStaff In mcolItems
            If oStaff.IsSupervisor = True Then
                If pstaff Is oStaff Then
                    bTmp = True
                End If
                SupervisorCount = SupervisorCount + 1
            End If
        Next
        IsLastSupervisor = (SupervisorCount = 1) And bTmp
End Function

Friend Function GetSuperState() As String
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oSt As a_Staff
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oSt In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oSt.GetSuperState
    Next
    For Each oSt In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oSt.GetSuperState
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
  Dim oSt As a_Staff
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oSt = New a_Staff
      oSt.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oSt.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oSt.BeginEdit
      Next
      If Not oSt.IsDeleted Then
        mcolItems.Add oSt, oSt.Key
      Else
        mcolDeleted.Add oSt
      End If
      Set oSt = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex - 1
  Set objPB = Nothing
End Sub


