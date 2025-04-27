VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_OOSLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Col As Collection
Private colDeleted As Collection
Private flgEditing As Boolean
Event PersonLegAdded()


Private Sub Class_Initialize()
  Set Col = New Collection
  Set colDeleted = New Collection
End Sub

Public Function Item(ByVal Index As Variant) As a_OOSLine
Attribute Item.VB_UserMemId = 0
  Set Item = Col(Index)
End Function

Public Function Count() As Long
  Count = Col.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = Col.[_NewEnum]
End Function

Public Sub BeginEdit()
  Dim obj As a_OOSLine
  
  If flgEditing Then Err.Raise 445
  
  For Each obj In Col
    obj.ChildBeginEdit
  Next
  flgEditing = True
End Sub

Public Sub CancelEdit()
  Dim obj As a_OOSLine
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

Public Sub ApplyEdit(ID As Long)
  Dim obj As a_OOSLine
  Dim lngIndex As Long
  
  If Not flgEditing Then Err.Raise 445

  For lngIndex = Col.Count To 1 Step -1
    Set obj = Col.Item(lngIndex)
    With obj
      If .IsDeleted Then Col.Remove lngIndex
      .ChildApplyEdit ID
    End With
  Next lngIndex
  flgEditing = False
End Sub

Public Sub Delete()
  Dim obj As a_OOSLine
  
  If Not flgEditing Then Err.Raise 445
  
  ' code to delete the object's data goes here
  For Each obj In Col
    With obj
      .BeginEdit
      .Delete
      .ApplyEdit lngResult
    End With
  Next
End Sub

Public Function Add() As a_OOSLine
  Dim obj As a_OOSLine
  If Not flgEditing Then Err.Raise 445
  
  Set obj = New a_OOSLine
  obj.SetAsChild Me
  obj.ChildBeginEdit
  Set Add = obj
  Set obj = Nothing
End Function

Public Sub Load(PID As Long)
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strWHERE As String
    Dim obj As a_OOSLine
    Dim parADO As New ADODB.Parameter
    Dim cmdADO As New ADODB.Command
    With cmdADO
        .CommandText = "GetOOSLines"
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oPC.CO
      .Parameters.Append .CreateParameter("[EnterOOSID]", adInteger, adParamInput) 'needs Size to work
       cmdADO.Parameters("[EnterOOSID]") = PID
       rs.Open cmdADO, , adOpenStatic, adLockReadOnly
    End With
  
  Do While Not rs.EOF
    Set obj = New a_OOSLine
    With obj
      .SetAsChild Me
      .ChildLoad rs("OOSL_ID")
    End With
    Col.Add Item:=obj
    rs.MoveNext
  Loop
  rs.Close

End Sub

Public Function Remove(ByVal Index As Variant) As Boolean  'added by Dom 5-6-7 early
Dim obj As a_OOSLine

    If Not flgEditing Then Err.Raise 445
    
    Remove = True
    Set obj = Col.Item(Index)
        With Col(Index)
            .BeginEdit
            .Delete
            .ApplyEdit lngResult
        End With
End Function
Friend Function AddOOSL(Child As a_OOSLine) As Boolean
    AddOOSL = True
    Col.Add Item:=Child
End Function

