VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "chex_COLDELLAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private strType As String
Private mcolItems As Collection
Private mcolDeleted As Collection
Dim MaxKeyUsed As Long
Private mlngEditing As Long
Public Event Valid(pResult As String)
Type SummType
  PID As String
  QtyAlloc As Long
  QtyOnHand As Long
End Type


Private Sub Class_Initialize()
  Set mcolItems = New Collection
  Set mcolDeleted = New Collection
End Sub

Public Function Count() As Long
  Count = mcolItems.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mcolItems.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As a_COLAllocation
Attribute Item.VB_UserMemId = 0
  Set Item = mcolItems(Index)
End Function

Public Function Add() As a_COLAllocation
Dim obj As a_COLAllocation
Dim lngEdit As Long

    If mlngEditing = 0 Then Err.Raise 445
    
    Set obj = New a_COLAllocation
    With obj
        .Initialize Me
        For lngEdit = 1 To mlngEditing
            .BeginEdit
        Next
    End With
    MaxKeyUsed = MaxKeyUsed + 1
    obj.Key = MaxKeyUsed & "k"
    Set Add = obj
    Set obj = Nothing
End Function

Friend Sub AddCOLAllocation(Child As a_COLAllocation)
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
Dim obj As a_COLAllocation

    If mcolDeleted.Count > 0 Then
        IsDirty = True
    Else
        For Each obj In mcolItems
            If obj.IsDirty Then
                IsDirty = True
                Exit For
            End If
        Next
    End If
End Function

Public Sub BeginEdit()
Dim obj As a_COLAllocation

    For Each obj In mcolItems
        obj.BeginEdit
    Next
    For Each obj In mcolDeleted
        obj.BeginEdit
    Next
    mlngEditing = mlngEditing + 1
End Sub

Public Sub ApplyEdit(pError As String)
Dim obj As a_COLAllocation
    If mlngEditing = 0 Then Err.Raise 445
    
    For Each obj In mcolItems
        obj.ApplyEdit
    Next
    For Each obj In mcolDeleted
        obj.ApplyEdit
    Next
    pError = ""
    Save pError, "NORMAL"
    mlngEditing = mlngEditing - 1
End Sub

Public Sub CancelEdit()
Dim obj As a_COLAllocation
Dim lngIndex As Long

    If mlngEditing = 0 Then Err.Raise 445
    
    For lngIndex = mcolItems.Count To 1 Step -1
        Set obj = mcolItems(lngIndex)
        If obj.IsNew Then
            obj.CancelEdit
            mcolItems.Remove lngIndex
        Else
            obj.CancelEdit
        End If
    Next
    For lngIndex = mcolDeleted.Count To 1 Step -1
        Set obj = mcolDeleted(lngIndex)
        With obj
            If Not .IsNew Then
                .CancelEdit
                If Not .IsDeleted Then
                    mcolItems.Add obj
                    mcolDeleted.Remove lngIndex
                End If
            Else
                mcolDeleted.Remove lngIndex
            End If
        End With
        Set obj = Nothing
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
Dim obj As a_COLAllocation

    Set objPB = New PropertyBag
    With objPB
        .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
        For Each obj In mcolItems
            lngCount = lngCount + 1
            .WriteProperty "Item" & CStr(lngCount), obj.GetSuperState
        Next
        For Each obj In mcolDeleted
            lngCount = lngCount + 1
            .WriteProperty "Item" & CStr(lngCount), obj.GetSuperState
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
Dim obj As a_COLAllocation

    Set objPB = New PropertyBag
    
    Set mcolItems = Nothing
    Set mcolItems = New Collection
    Set mcolDeleted = Nothing
    Set mcolDeleted = New Collection
    
    With objPB
        arBuffer = buffer
        .Contents = arBuffer
        
        For lngIndex = 1 To .ReadProperty("Count")
        
            Set obj = New a_COLAllocation
            obj.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
            obj.Key = lngIndex & "k"
            For lngEdit = 1 To mlngEditing
                obj.BeginEdit
            Next
            If Not obj.IsDeleted Then
                mcolItems.Add obj, obj.Key
            Else
                mcolDeleted.Add obj
            End If
            Set obj = Nothing
            
        Next
        
    End With
    MaxKeyUsed = lngIndex
    Set objPB = Nothing
End Sub
Public Property Get IsEditing() As Boolean
    IsEditing = (mlngEditing > 0)
End Property
Public Function GenerateCOLAllocationset(Optional pDELID As Long)
Dim obj As a_COLAllocation_P
    Set obj = New a_COLAllocation_P
    If pDELID > 0 Then strType = "DELIVERY"
    obj.GenerateCOLAllocationset pDELID, strType
    Set obj = Nothing
End Function

Public Function Load(Optional DELID As Long)
Dim objPersist As a_COLAllocation_P
Dim od As a_COLAllocation
    Set objPersist = New a_COLAllocation_P
    SetSuperState objPersist.Fetch(DELID)
    Set objPersist = Nothing
'    For Each od In mcolItems
'        od.Initialize Me
'    Next
End Function

Public Function Save(pError As String, pType As String)
Dim strBuffer As String
Dim objPersist As a_COLAllocation_P
    Set objPersist = New a_COLAllocation_P
    strBuffer = objPersist.Save(GetSuperState, strType)
    If strBuffer = "" Then
        pError = "Duplicate values in COLALLOC"
    End If

End Function

Private Function LimitExceeded(pResult As String) As Boolean
Dim summ As New Collection
Dim summItem As z_TEMP1
Dim od As a_COLAllocation
Dim i As Long
On Error GoTo ERRH2
    Set summ = New Collection
    LimitExceeded = False
    For Each od In mcolItems
        Set summItem = New z_TEMP1
        summItem.PID = od.PID
        summItem.QtyAlloc = od.AllocatedQty
        summItem.QtyOnHand = od.QtyOnHand
        summItem.Title = od.Title
        On Error GoTo ERRH
        summ.Add summItem, od.PID
    Next
    For i = 1 To summ.Count
        If summ(i).QtyAlloc > summ(i).QtyOnHand Then
            LimitExceeded = True
            pResult = pResult & IIf(Len(pResult) > 0, vbCrLf, "") & Left(summ(i).Title, 25)
        End If
    Next i
    Exit Function

ERRH:
    summ(od.PID).QtyAlloc = summ(od.PID).QtyAlloc + summItem.QtyAlloc
    Resume Next
    Exit Function
ERRH2:
    MsgBox Error
   ' Resume
End Function

Public Function GetStatus() As Boolean
Dim strresult As String
    If LimitExceeded(strresult) Then
        RaiseEvent Valid(strresult)
    Else
        RaiseEvent Valid("")
    End If
End Function

