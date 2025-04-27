VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_Companies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mcolItems As Collection
Private mcolDeleted As Collection
Dim MaxKeyUsed As Long
Private mlngEditing As Long

'Public Sub Load(pDefaultCompany As a_Company, pDefaultCOMPID As Long)
'  Fetch pDefaultCompany, pDefaultCOMPID
'End Sub

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

Public Function Item(ByVal Index As Variant) As a_Company
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function
Public Property Get FindCompanyByID(PID As Long) As a_Company
    On Error GoTo errHandler
Dim oComp As a_Company
    For Each oComp In mcolItems
        If oComp.ID = PID Then
            Set FindCompanyByID = oComp
            Exit For
        End If
    Next
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_Companies.FindCompanyByID(PID)", PID
End Property
Public Function Add() As a_Company
  Dim objComp As a_Company
  Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set objComp = New a_Company
  With objComp
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
    MaxKeyUsed = MaxKeyUsed + 1
    objComp.Key = MaxKeyUsed & "k"
  End With
  Set Add = objComp
  Set objComp = Nothing
End Function

Friend Sub AddCompany(pComp As a_Company)

  mcolItems.Add pComp, CStr(pComp.ID)

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
  Dim objComp As a_Company
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each objComp In mcolItems
      If objComp.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
Dim objComp As a_Company
  
  For Each objComp In mcolItems
    objComp.BeginEdit
  Next
  For Each objComp In mcolDeleted
    objComp.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub CancelEdit()
Dim objComp As a_Company
Dim lngIndex As Long
  
    If mlngEditing = 0 Then Err.Raise 445
    For lngIndex = mcolItems.Count To 1 Step -1
        Set objComp = mcolItems(lngIndex)
        If objComp.IsNew Then
            objComp.CancelEdit
            mcolItems.Remove lngIndex
        Else
            objComp.CancelEdit
        End If
    Next
'    For Each objComp In mcolItems
'      objComp.CancelEdit
'      objComp.
'    Next
    For lngIndex = mcolDeleted.Count To 1 Step -1
      Set objComp = mcolDeleted(lngIndex)
      With objComp
          .CancelEdit
        If Not .IsNew Then
          If Not .IsDeleted Then
            mcolItems.Add objComp, objComp.Key
            mcolDeleted.Remove lngIndex
          End If
        Else
          mcolDeleted.Remove lngIndex
        End If
      End With
        If objComp.IsEditing Then objComp.CancelEdit
      Set objComp = Nothing
    Next
  
  mlngEditing = mlngEditing - 1
  
End Sub

Public Sub ApplyEdit()
  Dim objComp As a_Company

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each objComp In mcolItems
    objComp.ApplyEdit
  Next
  For Each objComp In mcolDeleted
    objComp.ApplyEdit
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
  Dim objComp As a_Company
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each objComp In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), objComp.GetSuperState
    Next
    For Each objComp In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), objComp.GetSuperState
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
  Dim objComp As a_Company
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set objComp = New a_Company
      objComp.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      objComp.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        objComp.BeginEdit
      Next
      If Not objComp.IsDeleted Then
        mcolItems.Add objComp, CStr(objComp.ID)
      Else
        mcolDeleted.Add objComp
      End If
      Set objComp = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex - 1
  Set objPB = Nothing
End Sub

