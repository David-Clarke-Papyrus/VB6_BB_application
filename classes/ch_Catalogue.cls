VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_Catalogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolItems As Collection
Private mcolDeleted As Collection

Private mlngEditing As Long

Private Sub Class_Initialize()
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
End Sub

Public Function Count() As Long
  Count = mcolItems.Count
End Function

Public Function NewEnum() As IUnknown
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As a_Currency
  Set Item = mcolItems.Item(Index)
End Function
'''
Public Sub BeginEdit()
  Dim objCurrency As a_Currency

  For Each objCurrency In mcolItems
    objCurrency.BeginEdit
  Next
  For Each objCurrency In mcolDeleted
    objCurrency.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub ApplyEdit()
  Dim objCurrency As a_Currency

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each objCurrency In mcolItems
    objCurrency.ApplyEdit
  Next
  For Each objCurrency In mcolDeleted
    objCurrency.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
End Sub

Public Sub CancelEdit()
  Dim lngIndex As Long
  Dim objCurrency As a_Currency

  If mlngEditing = 0 Then Err.Raise 445
  
    For lngIndex = mcolItems.Count To 1 Step -1
        Set objCurrency = mcolItems(lngIndex)
        If objCurrency.IsNew Then
            objCurrency.CancelEdit
            mcolItems.Remove lngIndex
        Else
            objCurrency.CancelEdit
        End If
    Next
  
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set objCurrency = mcolDeleted(lngIndex)
    With objCurrency
      If Not .IsNew Then
        .CancelEdit
        If Not .IsDeleted Then
          mcolItems.Add objCurrency
          mcolDeleted.Remove lngIndex
        End If
      Else
        mcolDeleted.Remove lngIndex
      End If
    End With
    Set objCurrency = Nothing
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
  Dim objCurrency As a_Currency
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each objCurrency In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), objCurrency.GetSuperState
    Next
    For Each objCurrency In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), objCurrency.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
End Function

Friend Sub SetSuperState(Buffer As String)
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim objCurrency As a_Currency
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = Buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set objCurrency = New a_Currency
      objCurrency.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      For lngEdit = 1 To mlngEditing
        objCurrency.BeginEdit
      Next
      If Not objCurrency.IsDeleted Then
        mcolItems.Add objCurrency, objCurrency.ID & "k"
      Else
        mcolDeleted.Add objCurrency
      End If
      Set objCurrency = Nothing
    Next
  End With
  Set objPB = Nothing
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
  Dim objCurrency As a_Currency
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each objCurrency In mcolItems
      If objCurrency.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function
Public Function Add() As a_Currency
  Dim objCurrency As a_Currency
  Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set objCurrency = New a_Currency
  With objCurrency
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
  Set Add = objCurrency
  Set objCurrency = Nothing
End Function

Friend Sub AddCurrency(Child As a_Currency)

  mcolItems.Add Child, Child.ID & "k"

End Sub

Public Function Find(pDescription) As a_Currency
    Set Find = mcolItems(pDescription)
End Function
Public Function FindByID(PID As Long) As a_Currency
Dim oCurr As a_Currency
    For Each oCurr In mcolItems
        If oCurr.ID = PID Then
            Set FindByID = oCurr
            Exit For
        End If
    Next
End Function

'Public Function GetRateByName(pType As String) As Currency
'    GetRateByName = objCurrencies.Find(pType).ConvertToLocalFactor
'End Function


