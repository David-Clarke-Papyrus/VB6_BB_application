VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Staff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Col As Collection
Private colDeleted As Collection
Private flgEditing As Boolean

Public Sub Load(pCOMPID As Long)
  Fetch pCOMPID
End Sub
Private Sub Class_Initialize()
  Set Col = New Collection
  Set colDeleted = New Collection
End Sub

Public Function Count() As Long
  Count = Col.Count
End Function

Public Function NewEnum() As IUnknown
  Set NewEnum = Col.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As a_Staff
  Set Item = Col(Index)
End Function
Public Sub BeginEdit()
  Dim obj As a_Staff
  
  If flgEditing Then Err.Raise 445
  
  For Each obj In Col
    obj.ChildBeginEdit
  Next
  flgEditing = True
End Sub

Public Sub CancelEdit()
  Dim obj As a_Staff
  Dim lngIndex As Long
  
  If Not flgEditing Then Err.Raise 445
  
  flgEditing = False
  For lngIndex = Col.Count To 1 Step -1
    Set obj = Col.Item(lngIndex)
    With obj
      If .IsNew Then _
        Col.Remove lngIndex
      .ChildCancelEdit
    End With
  Next lngIndex
End Sub
Public Property Get IsEditing() As Boolean
    IsEditing = flgEditing
End Property

Public Sub ApplyEdit()
  Dim obj As a_Staff
  Dim lngIndex As Long
  
  If Not flgEditing Then Err.Raise 445

  For lngIndex = Col.Count To 1 Step -1
    Set obj = Col.Item(lngIndex)
    With obj
      If .IsDeleted Then Col.Remove lngIndex
      .ChildApplyEdit
    End With
  Next lngIndex
  flgEditing = False
End Sub

Public Function AddProposedStaff() As a_Staff
  Dim obj As a_Staff
  If Not flgEditing Then Err.Raise 445
  
  Set obj = New a_Staff
  obj.SetAsChild Me
  obj.ChildBeginEdit
  Set AddProposedStaff = obj
  Set obj = Nothing
End Function
Public Function MyChildLoad(PID As Long) As a_Staff
Dim lngResult As Long
Dim oStaff As a_Staff
    For Each oStaff In Col
        If oStaff.ID = PID Then
            Set MyChildLoad = oStaff
      '      oStaff.ChildLoad
        End If
    Next
End Function

Private Sub Fetch(pCOMPID As Long)
  Dim rs As ADODB.Recordset
  Dim strSQL As String
  Dim strWHERE As String
  Dim oStaff As a_Staff
  Dim fs As FileSystemObject
  
  Set rs = New ADODB.Recordset
  Set fs = New FileSystemObject
  Set oStaff = New a_Staff
  rs.Open "GetStaffPerCompany", oPC.CO, adOpenDynamic, adLockReadOnly
  
  Do While Not rs.EOF
    With oStaff
        .ID = rs("STAFF_ID")
        .StaffName = FNS(rs.Fields("STAFF_Name"))
        .COMPID = FNS(rs.Fields("STAFF_CO_ID"))
        .LoadColMember
        Col.Add oStaff
        Set oStaff = Nothing
        rs.MoveNext
    End With
  Loop
  rs.Close
  Set rs = Nothing
  Set fs = Nothing
  Set oStaff = Nothing
End Sub
Public Function MarkForRemoval(ByVal Index As Variant) As Boolean
Dim obj As a_Staff

    If Not flgEditing Then Err.Raise 445
    MarkForRemoval = True
    Set obj = Col.Item(Index)

 '   obj.BeginEdit
    obj.Delete
 '   obj.ApplyEdit lngResult

End Function

Friend Function AddToCol(pStaff As a_Staff) As Boolean
    AddToCol = True
    Col.Add Item:=pStaff
End Function
Public Function DataForLists() As Collection
Dim oStaff As a_Staff
Dim Col As Collection
Dim ar As Variant
Dim tmpID As Long
Dim tmpName As String
    ar = Array(2)
    For Each oStaff In Col
        tmpID = oStaff.ID
        tmpName = oStaff.StaffName
        ar = Array(tmpID, tmpName)
        Col.Add ar
    Next
    Set DataForLists = Col
End Function

Private Sub Class_Terminate()
  Set Col = Nothing
  Set colDeleted = Nothing
End Sub
