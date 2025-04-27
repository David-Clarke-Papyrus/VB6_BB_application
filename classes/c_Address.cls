VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Address"
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

Public Function Item(ByVal Index As Variant) As a_Address
  Set Item = mcolItems.Item(Index)
End Function

Public Function Add() As a_Address
  Dim oAddress As a_Address
  Dim lngEdit As Long
  
  If mlngEditing = 0 Then Err.Raise 445
  
  Set oAddress = New a_Address
  With oAddress
    .Initialize Me
    For lngEdit = 1 To mlngEditing
      .BeginEdit
    Next
  End With
  Set Add = oAddress
  Set oAddress = Nothing
End Function

Friend Sub AddTask(Child As a_Address)

  mcolItems.Add Child

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
  Dim oAddress As a_Address
  
  If mcolDeleted.Count > 0 Then
    IsDirty = True
  Else
    For Each oAddress In mcolItems
      If oAddress.IsDirty Then
        IsDirty = True
        Exit For
      End If
    Next
  End If
End Function

Public Sub BeginEdit()
  Dim oAddress As a_Address

  For Each oAddress In mcolItems
    oAddress.BeginEdit
  Next
  For Each oAddress In mcolDeleted
    oAddress.BeginEdit
  Next
  
  mlngEditing = mlngEditing + 1
End Sub

Public Sub ApplyEdit()
  Dim oAddress As a_Address

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oAddress In mcolItems
    oAddress.ApplyEdit
  Next
  For Each oAddress In mcolDeleted
    oAddress.ApplyEdit
  Next
  
  mlngEditing = mlngEditing - 1
End Sub

Public Sub CancelEdit()
  Dim lngIndex As Long
  Dim oAddress As a_Address

  If mlngEditing = 0 Then Err.Raise 445
  
  For Each oAddress In mcolItems
    oAddress.CancelEdit
  Next
  For lngIndex = mcolDeleted.Count To 1 Step -1
    Set oAddress = mcolDeleted(lngIndex)
    With oAddress
      If Not .IsNew Then
        .CancelEdit
        If Not .IsDeleted Then
          mcolItems.Add oAddress
          mcolDeleted.Remove lngIndex
        End If
      Else
        mcolDeleted.Remove lngIndex
      End If
    End With
    Set oAddress = Nothing
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
  Dim oAddress As a_Address
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "Count", mcolItems.Count + mcolDeleted.Count
    For Each oAddress In mcolItems
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oAddress.GetSuperState
    Next
    For Each oAddress In mcolDeleted
      lngCount = lngCount + 1
      .WriteProperty "Item" & CStr(lngCount), oAddress.GetSuperState
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
  Dim oAddress As a_Address
  
  Set objPB = New PropertyBag
  
  Set mcolItems = Nothing
  Set mcolItems = New Collection
  Set mcolDeleted = Nothing
  Set mcolDeleted = New Collection
  
  With objPB
    arBuffer = Buffer
    .Contents = arBuffer
    For lngIndex = 1 To .ReadProperty("Count")
      Set oAddress = New a_Address
      oAddress.SetSuperState .ReadProperty("Item" & CStr(lngIndex))
      For lngEdit = 1 To mlngEditing
        oAddress.BeginEdit
      Next
      If Not oAddress.IsDeleted Then
        mcolItems.Add oAddress
      Else
        mcolDeleted.Add oAddress
      End If
      Set oAddress = Nothing
    Next
  End With
  Set objPB = Nothing
End Sub



'=================
Public Sub Load(pDefaultCompany As a_Company, pDefaultCompanyID As Long)
  Fetch pDefaultCompany, pDefaultCompanyID
End Sub

Public Function AddProposedComp() As a_Company
  Dim obj As a_Company
  If Not flgEditing Then Err.Raise 445
  
  Set obj = New a_Company
  obj.SetAsChild Me
  obj.ChildBeginEdit
  Set AddProposedComp = obj
  Set obj = Nothing
End Function

Private Sub Fetch(pDefaultCompany As a_Company, pDefaultCompanyID As Long, Optional nLevels As Integer)
  Dim rs As ADODB.Recordset
  Dim strSQL As String
  Dim strWHERE As String
  Dim oComp As a_Company
  Dim fs As FileSystemObject
  
  Set rs = New ADODB.Recordset
  Set fs = New FileSystemObject
  rs.Open "GetCompanies", oADODBConn, adOpenDynamic, adLockReadOnly
  
        Do While Not rs.EOF
          Set oComp = New a_Company
          With oComp
            .SetAsChild Me
            .ChildLoad rs("Company_ID")
            If oComp.ID = pDefaultCompanyID Then Set pDefaultCompany = oComp
              Col.Add oComp
          End With
          Set oComp = Nothing
          rs.MoveNext
        Loop
  
  rs.Close
  Set rs = Nothing
  Set fs = Nothing
End Sub
Public Function MarkForRemoval(ByVal Index As Variant) As Boolean
Dim obj As a_Company

    If Not flgEditing Then Err.Raise 445
    MarkForRemoval = True
    Set obj = Col.Item(Index)

 '   obj.BeginEdit
    obj.Delete
 '   obj.ApplyEdit lngResult

End Function

Friend Function AddToCol(pComp As a_Company) As Boolean
    AddToCol = True
    Col.Add Item:=pComp
End Function
Public Function FindByCode(pCode As String) As a_Company
Dim oComp As a_Company
    For Each oComp In Col
        If oComp.CompanyCode = pCode Then
            Set FindByCode = oComp
            Exit For
        End If
    Next
End Function
Public Function FindByID(ID As Long) As a_Company
Dim oComp As a_Company
    For Each oComp In Col
        If oComp.ID = ID Then
            Set FindByID = oComp
            Exit For
        End If
    Next
End Function

Private Sub Class_Terminate()
  Set Col = Nothing
  Set colDeleted = Nothing

End Sub
