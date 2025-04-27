VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_CNL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event rowsChange()
Event CurrRowStatus(pMsg As String)
Private mcolItems As Collection
Private mcolDeleted As Collection
Private mParent As a_CN
Private mlngEditing As Long
Dim MaxKeyUsed As Long

Dim WithEvents oCNL As a_CNL
Attribute oCNL.VB_VarHelpID = -1

'Private bCanChangeDiscount As Boolean

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
  MaxKeyUsed = 0
    Exit Sub
errHandler:
    ErrorIn "ch_CNL.Class_Initialize"
End Sub
Public Sub component(pParent As a_CN)
    On Error GoTo errHandler
    Set mParent = pParent
    Exit Sub
errHandler:
    ErrorIn "ch_CNL.component(pParent)", pParent
End Sub
Public Function Count() As Long
    On Error GoTo errHandler
  Count = mcolItems.Count
    Exit Function
errHandler:
    ErrorIn "ch_CNL.Count"
End Function


Private Sub oCNL_Valid(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent CurrRowStatus(pMsg)
    Exit Sub
errHandler:
    ErrorIn "ch_CNL.oCNL_Valid(pMsg)", pMsg
End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
  Set NewEnum = mcolItems.[_NewEnum]
    Exit Function
errHandler:
    ErrorIn "ch_CNL.NewEnum"
End Function

Public Function Item(ByVal Index As Variant) As a_CNL
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
  Set Item = mcolItems.Item(Index)
    Exit Function
errHandler:
    ErrorIn "ch_CNL.Item(Index)", Index
End Function

Public Function Add() As a_CNL
    On Error GoTo errHandler
Dim lngEdit As Long
 
    If mlngEditing = 0 Then Err.Raise 445
    
    Set oCNL = New a_CNL
    With oCNL
      .Initialize Me
      For lngEdit = 1 To mlngEditing
        .BeginEdit
      Next
    End With
    MaxKeyUsed = MaxKeyUsed + 1
    oCNL.Key = MaxKeyUsed & "k"
    Set Add = oCNL
    oCNL.SetParentCN mParent
    Set oCNL = Nothing
    Exit Function
errHandler:
    ErrorIn "ch_CNL.Add"
End Function

Friend Sub AddCNLine(Child As a_CNL)
    On Error GoTo errHandler
    Child.SetParentCN mParent
    mcolItems.Add Child, Child.Key
    RaiseEvent rowsChange
    Exit Sub
errHandler:
    ErrorIn "ch_CNL.AddCNLine(Child)", Child
End Sub

Public Sub Remove(ByVal Index As Variant)
    On Error GoTo errHandler
  If mlngEditing = 0 Then Err.Raise 445

    With mcolItems(Index)
      .BeginEdit
      .Delete   'marks mudtProps.IsDeleted
      .ApplyEdit
    End With
    mcolDeleted.Add mcolItems(Index)
    mcolItems.Remove Index
    RaiseEvent rowsChange
    Exit Sub
errHandler:
    ErrorIn "ch_CNL.Remove(Index)", Index
End Sub

Public Function IsDirty() As Boolean
    On Error GoTo errHandler
  Dim oCNL As a_CNL
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oCNL In mcolItems
      If oCNL.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
    Exit Function
errHandler:
    ErrorIn "ch_CNL.IsDirty"
End Function
Public Function IsEditing() As Boolean
    On Error GoTo errHandler
    IsEditing = mlngEditing > 0
    Exit Function
errHandler:
    ErrorIn "ch_CNL.IsEditing"
End Function
Public Sub BeginEdit()
    On Error GoTo errHandler
  Dim oCNL As a_CNL

  For Each oCNL In mcolItems
    oCNL.BeginEdit
  Next
  For Each oCNL In mcolDeleted
    oCNL.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
    Exit Sub
errHandler:
    ErrorIn "ch_CNL.BeginEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim oCNL As a_CNL

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oCNL In mcolItems
    oCNL.ApplyEdit
  Next
  For Each oCNL In mcolDeleted
    oCNL.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    ErrorIn "ch_CNL.ApplyEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  Dim lngIndex As Long
  Dim oCNL As a_CNL

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oCNL In mcolItems
    If oCNL.IsEditing Then oCNL.CancelEdit
  Next
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oCNL = mcolDeleted(lngIndex)
    With oCNL
        If Not oCNL.IsNew Then
            .CancelEdit
            If Not .IsDeleted Then
                mcolItems.Add oCNL, .Key
                mcolDeleted.Remove lngIndex
            End If
        Else
            mcolDeleted.Remove lngIndex
        End If
    End With
    If oCNL.IsEditing Then oCNL.CancelEdit
    Set oCNL = Nothing
  Next
  
  mlngEditing = mlngEditing - 1
  Exit Sub
errHandler:
    ErrorIn "ch_CNL.CancelEdit"
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
    ErrorIn "ch_CNL.Delete"
End Sub

Friend Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oCNL As a_CNL
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oCNL In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oCNL.GetSuperState
    Next
    For Each oCNL In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oCNL.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    ErrorIn "ch_CNL.GetSuperState"
End Function
Function FindLineByID(PID As Long) As a_CNL
    On Error GoTo errHandler
Dim oCNL As a_CNL
    
    For Each oCNL In mcolItems
        If oCNL.CNLineID = PID Then
            Set FindLineByID = oCNL
            Exit For
        End If
    Next
    Exit Function
errHandler:
    ErrorIn "ch_CNL.FindLineByID(PID)", PID
End Function
Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oCNL As a_CNL
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oCNL = New a_CNL
      oCNL.SetParentCN mParent
      oCNL.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oCNL.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oCNL.BeginEdit
      Next
      If Not oCNL.IsDeleted Then
        mcolItems.Add oCNL, oCNL.Key
      Else
        mcolDeleted.Add oCNL
      End If
      oCNL.CalculateLine
      Set oCNL = Nothing
    Next
    MaxKeyUsed = .ReadProperty("Count")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    ErrorIn "ch_CNL.SetSuperState(buffer)", buffer
End Sub
Public Sub SortLines(pField As enSortField, pASCEND As Boolean)
    On Error GoTo errHandler
    CollectionSort mcolItems, pField, pASCEND
    Exit Sub
errHandler:
    ErrorIn "ch_CNL.SortLines(pField,pASCEND)", Array(pField, pASCEND)
End Sub

Private Function CollectionSort(ByRef oCollection As Collection, pProperty As enSortField, Optional bSortAscending As Boolean = True) As Long
    On Error GoTo errHandler
    Dim lSort1 As Long, lSort2 As Long
    Dim vTempItem1 As Variant, vTempItem2 As Variant, bSwap As Boolean
    
    For lSort1 = 1 To oCollection.Count - 1
        For lSort2 = lSort1 + 1 To oCollection.Count
            If bSortAscending Then
                If UCase(oCollection(lSort1).Properties(pProperty)) > UCase(oCollection(lSort2).Properties(pProperty)) Then
                    bSwap = True
                Else
                    bSwap = False
                End If
            Else
                If UCase(oCollection(lSort1).Properties(pProperty)) < UCase(oCollection(lSort2).Properties(pProperty)) Then
                    bSwap = True
                Else
                    bSwap = False
                End If
            End If
            If bSwap Then
                'Store the items
                If VarType(oCollection(lSort1)) = vbObject Then
                    Set vTempItem1 = oCollection(lSort1)
                Else
                    vTempItem1 = oCollection(lSort1)
                End If
                
                If VarType(oCollection(lSort2)) = vbObject Then
                    Set vTempItem2 = oCollection(lSort2)
                Else
                    vTempItem2 = oCollection(lSort2)
                End If
                
                
                oCollection.Remove lSort2
                oCollection.Add vTempItem2, vTempItem2.Key, lSort1
                
            End If
        Next
    Next
    Exit Function

    Exit Function
errHandler:
    ErrorIn "ch_CNL.CollectionSort(oCollection,pProperty,bSortAscending)", Array(oCollection, _
         pProperty, bSortAscending)
End Function

Public Sub swap(iUP As String, iDown As String)
    On Error GoTo errHandler
Dim sTemp As String
Dim iTemp As Long
    iTemp = mcolItems(iDown).Sequence
    mcolItems(iDown).Sequence = mcolItems(iUP).Sequence
    mcolItems(iUP).Sequence = iTemp
    Exit Sub
errHandler:
    ErrorIn "ch_CNL.swap(iUP,iDown)", Array(iUP, iDown)
End Sub



