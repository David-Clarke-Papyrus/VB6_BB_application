VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_GDNLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event rowsChange()
Event CurrRowStatus(pMsg As String)
Private mcolItems As Collection
Private mcolDeleted As Collection
Private mParent As a_GDN
Private mlngEditing As Long
Dim MaxKeyUsed As Long
Dim WithEvents oGDNL As a_InvoiceLine
Attribute oGDNL.VB_VarHelpID = -1
Dim ar() As Integer

'Private bCanChangeDiscount As Boolean

Private Sub Class_Initialize()
    On Error GoTo errHandler
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
  ReDim ar(200, 2)
  MaxKeyUsed = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.Class_Initialize"
End Sub
Public Sub DecrementMaxKeyUsed()
    If MaxKeyUsed > 0 Then MaxKeyUsed = MaxKeyUsed - 1
End Sub

Public Sub component(pParent As a_GDN)
    On Error GoTo errHandler
    Set mParent = pParent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.Component(pParent)", pParent
End Sub
Public Function Count() As Long
    On Error GoTo errHandler
  Count = mcolItems.Count
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.Count"
End Function
Function FindLineByCOID(pCOLID As Long) As a_InvoiceLine
    On Error GoTo errHandler
Dim oGDNL As a_InvoiceLine
    
    For Each oGDNL In mcolItems
        If oGDNL.COLID = pCOLID Then
            Set FindLineByCOID = oGDNL
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.FindLineByCOID(pCOLID)", pCOLID
End Function

Function FindLineByID(pGDNLID As Long) As a_InvoiceLine
    On Error GoTo errHandler
Dim oGDNL As a_InvoiceLine
    
    For Each oGDNL In mcolItems
        If oGDNL.InvoiceLineID = pGDNLID Then
            Set FindLineByID = oGDNL
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.FindLineByID(pGDNLID)", pGDNLID
End Function
Function FindLineBySeqNum(Seq As Long) As a_InvoiceLine
    On Error GoTo errHandler
Dim oGDNL As a_InvoiceLine
    
    For Each oGDNL In mcolItems
        If oGDNL.Sequence = Seq Then
            Set FindLineBySeqNum = oGDNL
            Exit For
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.FindLineBySeqNum(Seq)", Seq
End Function


Private Sub oGDNL_Valid(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent CurrRowStatus(pMsg)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.oGDNL_Valid(pMsg)", pMsg
End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    On Error GoTo errHandler
  Set NewEnum = mcolItems.[_NewEnum]
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.NewEnum"
End Function

Public Function Item(ByVal Index As Variant) As a_InvoiceLine
Attribute Item.VB_UserMemId = 0
    On Error GoTo errHandler
    Set Item = mcolItems.Item(Index)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.Item(Index)", Index
End Function

Public Function Add() As a_InvoiceLine
    On Error GoTo errHandler
Dim lngEdit As Long
 
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oGDNL = New a_InvoiceLine
  With oGDNL
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
    MaxKeyUsed = MaxKeyUsed + 1
    oGDNL.Key = MaxKeyUsed & "k"
  Set Add = oGDNL
  Set oGDNL = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.Add"
End Function

Public Sub AddDocLine(Child As a_InvoiceLine)
    On Error GoTo errHandler
    Child.SetParentInvoice mParent
    Child.Sequence = mcolItems.Count + 1
    mcolItems.Add Child, Child.Key
    RaiseEvent rowsChange
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.AddDocLine(Child)", Child
End Sub

Public Sub Remove(ByVal Index As Variant)
    On Error GoTo errHandler
Dim iTmp As Long

'LogSaveToFile "Invoice line removed:" & CStr(Index)
  If mlngEditing = 0 Then Err.Raise 445

    With mcolItems(Index)
        iTmp = .Sequence
        .BeginEdit
        .Delete
        .ApplyEdit
    End With
    mcolDeleted.Add mcolItems(Index)
    mcolItems.Remove Index
    CloseUpSequence iTmp
    RaiseEvent rowsChange
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.Remove(Index)", Index
End Sub
Private Sub CloseUpSequence(pTmp As Long)
Dim i As Integer

    For i = 1 To mcolItems.Count
        If mcolItems(i).Sequence > pTmp Then
            mcolItems(i).Sequence = mcolItems(i).Sequence - 1
        End If
    Next i
End Sub
Public Function IsDirty() As Boolean
    On Error GoTo errHandler
  Dim oGDNL As a_InvoiceLine
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oGDNL In mcolItems
      If oGDNL.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.IsDirty"
End Function
Public Function IsEditing() As Boolean
    On Error GoTo errHandler
    IsEditing = mlngEditing > 0
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.IsEditing"
End Function
Public Sub BeginEdit()
    On Error GoTo errHandler
  Dim oGDNL As a_InvoiceLine

  For Each oGDNL In mcolItems
    oGDNL.BeginEdit
  Next
  For Each oGDNL In mcolDeleted
    oGDNL.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.BeginEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
  Dim oGDNL As a_InvoiceLine

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oGDNL In mcolItems
    oGDNL.ApplyEdit
  Next
  For Each oGDNL In mcolDeleted
    oGDNL.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.ApplyEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  Dim lngIndex As Long
  Dim oGDNL As a_InvoiceLine

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oGDNL In mcolItems
    If oGDNL.IsEditing Then oGDNL.CancelEdit
  Next
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oGDNL = mcolDeleted(lngIndex)
    With oGDNL
      If Not .IsNew Then
        .CancelEdit
        If Not .IsDeleted Then
          mcolItems.Add oGDNL, oGDNL.Key
          mcolDeleted.Remove lngIndex
        End If
      Else
        mcolDeleted.Remove lngIndex
      End If
    End With
    If oGDNL.IsEditing Then oGDNL.CancelEdit
    Set oGDNL = Nothing
  Next
  '''''''
    MaxKeyUsed = mcolItems.Count
  
  mlngEditing = mlngEditing - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.CancelEdit"
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
    ErrorIn "ch_GDNLine.Delete"
End Sub

Friend Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim lngCount As Long
  Dim oGDNL As a_InvoiceLine
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oGDNL In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oGDNL.GetSuperState
    Next
    For Each oGDNL In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oGDNL.GetSuperState
    Next
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.GetSuperState"
End Function

Friend Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngEdit As Long
  Dim oGDNL As a_InvoiceLine
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oGDNL = New a_InvoiceLine
      oGDNL.SetParentInvoice mParent
      oGDNL.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      oGDNL.Key = lngIndex & "k"
      For lngEdit = 1 To mlngEditing
        oGDNL.BeginEdit
      Next
      If Not oGDNL.IsDeleted Then
        mcolItems.Add oGDNL, oGDNL.Key
      Else
        mcolDeleted.Add oGDNL
      End If
      oGDNL.CalculateLine  'recalculate dependant values
      Set oGDNL = Nothing
    Next
    MaxKeyUsed = .ReadProperty("Count")
  End With

  Set objPB = Nothing
  
  
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.SetSuperState(buffer)", buffer
End Sub
Public Sub SortDocumentLines(pField As enSortField, pASCEND As Boolean)
    On Error GoTo errHandler
    CollectionSort mcolItems, pField, pASCEND
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ch_GDNLine.SortDocumentLines(pField,pASCEND)", Array(pField, pASCEND)
End Sub


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
    ErrorIn "ch_GDNLine.swap(iUP,iDown)", Array(iUP, iDown)
End Sub


