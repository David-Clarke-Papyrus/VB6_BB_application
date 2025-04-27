VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_RoundingRule"
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

Public Function Item(ByVal Index As Variant) As a_RoundingRule
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems.Item(Index)
End Function
Public Function Add() As a_RoundingRule
  Dim objRR As a_RoundingRule
  Dim lngEdit As Long
  
    If mlngEditing = 0 Then Err.Raise 445
    
    Set objRR = New a_RoundingRule
    With objRR
        .Initialize Me
        For lngEdit = 1 To mlngEditing
          .BeginEdit
        Next
    End With
    MaxKeyUsed = MaxKeyUsed + 1
    objRR.Key = MaxKeyUsed & "k"
    Set Add = objRR
    Set objRR = Nothing
End Function

Friend Sub AddRR(pRR As a_RoundingRule)

  mcolItems.Add pRR, pRR.Key

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
  Dim objRR As a_RoundingRule
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each objRR In mcolItems
      If objRR.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
Dim objRR As a_RoundingRule
  
  For Each objRR In mcolItems
    objRR.BeginEdit
  Next
  For Each objRR In mcolDeleted
    objRR.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub CancelEdit()
Dim objRR As a_RoundingRule
Dim lngIndex As Long
  
    If mlngEditing = 0 Then Err.Raise 445
    For lngIndex = mcolItems.Count To 1 Step -1
        Set objRR = mcolItems(lngIndex)
        If objRR.IsNew Then
            objRR.CancelEdit
            mcolItems.Remove lngIndex
        Else
            objRR.CancelEdit
        End If
    Next
'    For Each objComp In mcolItems
'      objComp.CancelEdit
'      objComp.
'    Next
    For lngIndex = mcolDeleted.Count To 1 Step -1
      Set objRR = mcolDeleted(lngIndex)
      With objRR
        If Not .IsNew Then
          .CancelEdit
          If Not .IsDeleted Then
            mcolItems.Add objRR, objRR.Key
            mcolDeleted.Remove lngIndex
          End If
        Else
          mcolDeleted.Remove lngIndex
        End If
      End With
    If objRR.IsEditing Then objRR.CancelEdit
      Set objRR = Nothing
    Next
  
  mlngEditing = mlngEditing - 1
  
End Sub

Public Sub ApplyEdit()
  Dim objRR As a_RoundingRule
  If mlngEditing = 0 Then Err.Raise 445
  
  For Each objRR In mcolItems
    objRR.ApplyEdit
  Next
  For Each objRR In mcolDeleted
    objRR.ApplyEdit
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
  Dim objRR As a_RoundingRule
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each objRR In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), objRR.GetSuperState
    Next
    For Each objRR In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), objRR.GetSuperState
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
  Dim objRR As a_RoundingRule
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set objRR = New a_RoundingRule
      objRR.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      objRR.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        objRR.BeginEdit
      Next
      If Not objRR.IsDeleted Then
        mcolItems.Add objRR, objRR.Key
      Else
        mcolDeleted.Add objRR
      End If
      Set objRR = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex - 1
  Set objPB = Nothing
End Sub

Public Function GetRoundTo(pVal As Long) As Long
Dim oRR As a_RoundingRule
Dim lng As Long
    GetRoundTo = 1
    For Each oRR In Me
        If pVal >= oRR.LowerBound And pVal < oRR.UpperBound Then
            lng = oRR.RoundTo
            Exit For
        End If
    Next
    If lng = 0 Then
        GetRoundTo = 1
    Else
        GetRoundTo = lng
    End If
End Function


