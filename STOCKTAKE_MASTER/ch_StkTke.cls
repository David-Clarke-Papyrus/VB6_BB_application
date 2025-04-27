VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ch_Stktke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private colStktkeLines As Collection
Private colDeleted As Collection

Private flgEditing As Boolean


Private Sub Class_Initialize()
  Set colStktkeLines = New Collection
End Sub

Public Function Item(ByVal Index As Variant) As a_StktkeLine
  Set Item = colStktkeLines(Index)
End Function

Public Function Count() As Long
  Count = colStktkeLines.Count
End Function

Public Function NewEnum() As IUnknown
  Set NewEnum = colStktkeLines.[_NewEnum]
End Function

Public Sub BeginEdit()
  Dim oStktkeLine As a_StktkeLine
  
  If flgEditing Then Err.Raise 445
  
  For Each oStktkeLine In colStktkeLines
    oStktkeLine.ChildBeginEdit
  Next
  flgEditing = True
End Sub

Public Function Add() As a_StktkeLine
  Dim oStktkeLine As a_StktkeLine
  
  Set oStktkeLine = New a_StktkeLine
  If flgEditing Then oStktkeLine.ChildBeginEdit
  colStktkeLines.Add Item:=oStktkeLine
  Set Add = oStktkeLine
End Function

Public Sub Delete()
  Dim oStktkeLine As a_StktkeLine
  
  If Not flgEditing Then Err.Raise 445
  'code to delete the object's data goes here
  For Each oStktkeLine In colStktkeLines
    With oStktkeLine
      .ChildBeginEdit
      .ChildDelete
      .ChildApplyEdit
    End With
  Next
End Sub

Public Sub ApplyEdit(StktkeID As Long)
  Dim oStktkeLine As a_StktkeLine
  Dim lngIndex As Long
  
  If Not flgEditing Then Err.Raise 445
  
  For lngIndex = colStktkeLines.Count To 1 Step -1
    Set oStktkeLine = colStktkeLines.Item(lngIndex)
    With oStktkeLine
      If .IsDeleted Then colStktkeLines.Remove lngIndex
      .ChildApplyEdit 'StktkeID
    End With
  Next lngIndex
  flgEditing = False
End Sub

Public Sub CancelEdit()
  Dim oStktkeLine As a_StktkeLine
  Dim lngIndex As Long
  
  If Not flgEditing Then Err.Raise 445
  
  flgEditing = False
  For lngIndex = colStktkeLines.Count To 1 Step -1
    Set oStktkeLine = colStktkeLines.Item(lngIndex)
    With oStktkeLine
      If .IsNew Then colStktkeLines.Remove lngIndex
      .ChildCancelEdit
    End With
  Next lngIndex
End Sub

Friend Sub Load(StktkeID As Long)
  If flgEditing Then Err.Raise 445
  FetchOrderLines StktkeID
  
End Sub

Private Sub FetchOrderLines(StktkeID As Long)
  Dim rs As ADODB.Recordset
  Dim cmd As New ADODB.Command
  Dim oStktkeLine As a_StktkeLine
  Dim strSQL As String
  
  strSQL = "SELECT StockAdjust_Line_ID FROM " & _
           "[StockAdjusment Line] WHERE SAL_SA_ID = " & StktkeID
  Set cmd.ActiveConnection = oADODBConn
  Set rs = New ADODB.Recordset
  cmd.CommandText = strSQL
  rs.Open cmd, , adOpenDynamic, adLockReadOnly
  Do While Not rs.EOF
    Set oStktkeLine = New a_StktkeLine
    oStktkeLine.ChildLoad rs![StockAdjusment Line]
    colStktkeLines.Add Item:=oStktkeLine
    rs.MoveNext
    Set oStktkeLine = Nothing
  Loop
  rs.Close
  Set rs = Nothing
  Set cmd = Nothing
End Sub


Private Sub Remove(ByVal Index As Variant)
  Dim oStktkeLine As a_StktkeLine
  
  If Not flgEditing Then Err.Raise 445
  Set oStktkeLine = colStktkeLines.Item(Index)
  oStktkeLine.ChildDelete
  colDeleted.Add Item:=oStktkeLine
  colStktkeLines.Remove Index
    
End Sub



