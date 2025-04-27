VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProductsColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private colDisplay As Collection


Private Sub Class_Initialize()
  Set colDisplay = New Collection
End Sub

Public Function Count() As Long
  Count = colDisplay.Count
End Function


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colDisplay.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As ProductDisplay
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function

Public Sub Load(Optional ISBN As String, Optional Title As String, _
                Optional Author As String)
  'load data from database
  Fetch ISBN, Title, Author
End Sub

Public Sub LoadBookFind(Optional ISBN As String, Optional Title As String, _
                Optional Author As String)
  'Load from database
  FetchBF ISBN, Title, Author
                
End Sub

Private Sub FetchBF(ISBN As String, Title As String, Author As String)
  Dim intSearchResult As Long
  Dim i As Long
  Dim sField As String
  Dim strSearch As String
  'check if BF is accessible
  If Not flgBookFindOK Then PrepareBookfind
  If Not flgBookFindOK Then Exit Sub
  
  If Len(ISBN) > 0 Then
    sField = "BN"
    strSearch = ISBN
  End If
  If Len(Title) > 0 Then
    sField = "TI"
    strSearch = Title
  End If
    
  If Len(Author) > 0 Then
    sField = "AU"
    strSearch = Author
  End If
  
On Error GoTo ErrorHandler
  intSearchResult = bfSearch(strSearch, sField)
  If intSearchResult = 0 Then Exit Sub
  
  Dim objDisplay As ProductDisplay
    
  For i = 1 To intSearchResult
    GetRecord (i)
    LoadProdRec (i)
    If iERR <> 0 Then
       MsgBox "Error in bookfind GET"
       Exit Sub
    End If
    Set objDisplay = New ProductDisplay
    With objDisplay
      .ISBN = BFRec.ISBN
      .Author = Trim$(BFRec.MainAuthor)
      .Title = Trim$(BFRec.Title)
      .Price = Val(BFRec.LocalPrice)
      .Publisher = Trim$(BFRec.PublisherName)
      colDisplay.Add objDisplay
      Set objDisplay = Nothing
    End With
  Next
      
MyExit:
  Exit Sub
ErrorHandler:
  MsgBox "Error Searching on BookFind!", vbOKOnly, "ERROR!"
GoTo MyExit
End Sub

Private Sub Fetch(ISBN As String, Title As String, Author As String)
  Dim rsProducts As ADODB.Recordset
  Dim cmd As ADODB.Command
  Dim strSQL As String
  Dim strWhere As String
  Dim objDisplay As ProductDisplay
    
On Error GoTo ErrorHandler
  If Not flgDbObjConnected Then
    MsgBox "Databse is NOT connected!", vbOKOnly, "WARNING!"
    Exit Sub
  End If
  
  strWhere = ""
  If Trim$(ISBN) <> "" Then
      strWhere = "P_ISBN='" & Trim$(ISBN) & "'"
  End If
  If Trim$(Author) <> "" Then
  If strWhere <> "" Then strWhere = strWhere & " AND "
    strWhere = strWhere & "P_MainAuthor LIKE '%" & Trim$(Author) & "%'"
  End If
  If Trim$(Title) <> "" Then
    If strWhere <> "" Then strWhere = strWhere & " AND "
    strWhere = strWhere & "P_Title LIKE '%" & Trim$(Title) & "%'"
  End If
  strSQL = "SELECT Product_ID, P_ISBN, P_Title, P_MainAuthor, P_SAPrice, P_StockBalance FROM Product WHERE " & strWhere
  'strSQL = "Select P_ISBN, P_Title, P_MainAuthor FROM Product WHERE P_MainAuthor LIKE 'Morrison%'"
  Set cmd = New ADODB.Command
  Set cmd.ActiveConnection = dbObj
  cmd.CommandText = strSQL
  Set rsProducts = New ADODB.Recordset
  rsProducts.Open cmd, , adOpenStatic, adLockOptimistic
  Do While Not rsProducts.EOF
    Set objDisplay = New ProductDisplay
    With objDisplay
      .BookID = rsProducts!Product_ID
      .ISBN = NZS(rsProducts!P_ISBN)
      .Title = NZS(rsProducts!P_Title)
      .Author = NZS(rsProducts!P_MainAuthor)
      .Price = NZ(rsProducts!P_SAPrice)
      .Stock = NZ(rsProducts!P_StockBalance)
      colDisplay.Add objDisplay
      Set objDisplay = Nothing
      rsProducts.MoveNext
    End With
  Loop
MyExit:
  Set cmd = Nothing
  rsProducts.Close
  Set rsProducts = Nothing
  Exit Sub
ErrorHandler:
  MsgBox "Error: " & dbObj.Errors.Item(0), vbOKOnly, "ERROR!"
GoTo MyExit
End Sub

