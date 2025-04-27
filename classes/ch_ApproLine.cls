VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_APPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mParent As a_APP
Private mcolItems As Collection
Private colDeleted As Collection
Private mlngEditing As Long
Dim MaxKeyUsed As Long
Dim WithEvents oAPPL As a_APPL
Attribute oAPPL.VB_VarHelpID = -1
Event rowsChange()

Public Sub component(pParent As a_APP)
    Set mParent = pParent
End Sub

Private Sub Class_Initialize()
  Set mcolItems = New Collection
  Set colDeleted = New Collection
  MaxKeyUsed = 0
End Sub

Public Function Item(ByVal Index As Variant) As a_APPL
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems(Index)
End Function

Public Function Count() As Long
  Count = mcolItems.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Friend Function GetSuperState() As String
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oAPPL As a_APPL
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + colDeleted.Count
    For Each oAPPL In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oAPPL.GetSuperState
    Next
    For Each oAPPL In colDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oAPPL.GetSuperState
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
  Dim oAPPL As a_APPL
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set colDeleted = Nothing
  Set colDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oAPPL = New a_APPL
      oAPPL.SetParentAPP mParent
      oAPPL.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oAPPL.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oAPPL.BeginEdit
      Next
      If Not oAPPL.IsDeleted Then
        mcolItems.Add oAPPL, oAPPL.Key
      Else
        colDeleted.Add oAPPL
      End If
      Set oAPPL = Nothing
    Next
  End With
  MaxKeyUsed = lngIndex - 1
  Set objPB = Nothing
End Sub

Public Function IsDirty() As Boolean
  Dim oAPPL As a_APPL
  
  If colDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oAPPL In mcolItems
      If oAPPL.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
  Dim oAL As a_APPL

  For Each oAL In mcolItems
    oAL.BeginEdit
  Next
  For Each oAL In colDeleted
    oAL.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub ApplyEdit()
  Dim oAL As a_APPL

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
  Dim oAL As a_APPL

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
          mcolItems.Add oAL, oAL.Key
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

Public Function Add() As a_APPL
    On Error GoTo errHandler
Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oAPPL = New a_APPL
  With oAPPL
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    oAPPL.Key = MaxKeyUsed & "k"
    oAPPL.SetParentAPP mParent
  Set Add = oAPPL
  Set oAPPL = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_APPL.Add"
End Function

Friend Sub AddApproLine(Child As a_APPL)
    Child.SetParentAPP mParent
    mcolItems.Add Child, Child.Key
    RaiseEvent rowsChange
End Sub

Public Sub Remove(ByVal Index As Variant)
  If mlngEditing = 0 Then Err.Raise 445

  With mcolItems.Item(Index)
    .BeginEdit
    .Delete
    .ApplyEdit
  End With
  colDeleted.Add mcolItems(Index)
  mcolItems.Remove Index
End Sub
Function FindLineByID(PID As Long) As a_APPL
Dim oAPPL As a_APPL
    
    For Each oAPPL In mcolItems
        If oAPPL.APPLID = PID Then
            Set FindLineByID = oAPPL
            Exit For
        End If
    Next
End Function

Function FindLineBySeqNum(Seq As Integer) As a_APPL
    On Error GoTo errHandler
Dim oAPL As a_APPL
    
    For Each oAPL In mcolItems
        If oAPL.Sequence = Seq Then
            Set FindLineBySeqNum = oAPL
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_APPL.FindLineBySeqNum(Seq)", Seq
End Function
Public Sub SortLines(pField As enSortField, pASCEND As Boolean)
    On Error GoTo errHandler
    CollectionSort mcolItems, pField, pASCEND
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_APPL.SortLines(pField,pASCEND)", Array(pField, pASCEND)
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

errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_APPL.CollectionSort(oCollection,pProperty,bSortAscending)", Array(oCollection, _
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
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_APPL.swap(iUP,iDown)", Array(iUP, iDown)
End Sub

