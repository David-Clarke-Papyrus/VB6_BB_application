VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Product"
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

Public Function Item(ByVal Index As Variant) As d_Product
Attribute Item.VB_UserMemId = 0
  Set Item = colDisplay(Index)
End Function

Public Sub Load(Optional ISBN As String, Optional Title As String, _
                Optional Author As String, Optional pSTarts As String, Optional MaxResult As Long, Optional pbIncludeObsolete As Boolean)
  'load data from database
  Fetch ISBN, Title, Author, pSTarts, MaxResult, pbIncludeObsolete
End Sub

Public Sub LoadBookFind(Optional ISBN As String, Optional Title As String, _
                Optional Author As String, Optional pSTarts As String, Optional MaxReturn As Long)
  'Load from database
   oPC.CO.BF.DestroyResultsList
  FetchBF ISBN, Title, Author, pSTarts, MaxReturn
                
End Sub

Private Sub FetchBF(ISBN As String, Optional Title As String, _
                    Optional Author As String, Optional pSTarts As String, Optional MaxReturn As Long)
Dim intSearchResult, oldSearchResults As Long
Dim i As Integer
Dim strSearch As String, strFirstPart As String
Dim sTab, sNull As String
  'check if BF is accessible
 ' If Not flgBookFindOK Then objPapyCon.BF.PrepareBookfind
  'If Not flgBookFindOK Then Exit Sub
  
  sTab = Chr(9)
  sNull = Chr(0)
  
  strFirstPart = "FIND" & sTab
  If Len(ISBN) > 0 Then
    strSearch = strFirstPart & "BN" & sTab & ISBN
    GoTo Search
  End If
  If Len(Trim$(Title)) > 0 Then strSearch = "TI" & sTab & Title
  If Len(strSearch) > 0 Then strSearch = strSearch & sNull & "AND" & sTab
  If Len(Trim$(Author)) > 0 Then strSearch = strSearch & "AU" & sTab & Author
  If Len(strSearch) > 0 Then strSearch = strSearch & sNull & "AND" & sTab
  If Len(Trim$(pSTarts)) > 0 Then strSearch = strSearch & "TI" & sTab & pSTarts
  
  
Search:
On Error GoTo ErrorHandler
  iERR = oPC.CO.BF.Search(strFirstPart & strSearch, intSearchResult)
  If intSearchResult = 0 Then
    
    Exit Sub
  Else
    If intSearchResult > MaxReturn And MaxReturn > 0 Then
      oldSearchResults = intSearchResult
      intSearchResult = MaxReturn
    End If
  End If
  
  Dim objDisplay As d_Product
    
  For i = 1 To intSearchResult
    iERR = oPC.CO.BF.GetRecordByNum(i)
    oPC.CO.BF.LoadProdRec i
    If iERR <> 0 Then
           MsgBox "Error in bookfind c_Prodcut:FetchBF : iERR = " & iERR & vbCrLf & "Code: =" & ISBN & vbCrLf & "iREcsFound = " & intSearchResult
       Exit Sub
    End If
    Set objDisplay = New d_Product
    With objDisplay
      .Code = BFRec.Code
      .Author = Trim$(BFRec.MainAuthor)
      .Title = Trim$(BFRec.Title)
      .Price = val(BFRec.LocalPrice)
      .Publisher = Trim$(BFRec.PublisherName)
      colDisplay.Add objDisplay
      Set objDisplay = Nothing
    End With
  Next i
      
MyExit:
  Exit Sub
ErrorHandler:
  MsgBox "Error Searching on BookFind!", vbOKOnly, "ERROR!"
GoTo MyExit
End Sub

Private Sub Fetch(ISBN As String, Title As String, Author As String, pSTarts As String, MaxResult As Long, pbIncludeObsolete As Boolean)
  Dim rsProducts As ADODB.Recordset
  Dim cmd As ADODB.Command
  Dim strSQL As String
  Dim strWHERE As String
  Dim objDisplay As d_Product
  Dim i As Long
  Dim bMustClose As Boolean
  
    
On Error GoTo ErrorHandler
  If MaxResult = 0 Then MaxResult = 1000
  strWHERE = ""
  If Trim$(pSTarts) <> "" Then
    If strWHERE <> "" Then strWHERE = strWHERE & " AND "
      strWHERE = strWHERE & "P_Title LIKE '" & Trim$(pSTarts) & "%'"
    End If
  If Trim$(ISBN) <> "" Then
      strWHERE = "P_ISBN='" & Trim$(ISBN) & "'"
  End If
  If Trim$(Author) <> "" Then
  If strWHERE <> "" Then strWHERE = strWHERE & " AND "
    strWHERE = strWHERE & "P_MainAuthor LIKE '%" & Trim$(Author) & "%'"
  End If
  If Trim$(Title) <> "" Then
    If strWHERE <> "" Then strWHERE = strWHERE & " AND "
    strWHERE = strWHERE & "P_Title LIKE '%" & Trim$(Title) & "%'"
  End If
  If pbIncludeObsolete = False Then
    If strWHERE <> "" Then strWHERE = strWHERE & " AND "
    strWHERE = strWHERE & "P_Obsolete = False"
  End If
  strSQL = "SELECT P_ID, P_Code, P_Title, P_MainAuthor, P_Price2, P_QtyOnHand FROM Product WHERE " & strWHERE
  'strSQL = "Select P_ISBN, P_Title, P_MainAuthor FROM Product WHERE P_MainAuthor LIKE 'Morrison%'"
  Set cmd = New ADODB.Command
  Set cmd.ActiveConnection = oPC.CO
  cmd.CommandText = strSQL
  Set rsProducts = New ADODB.Recordset
  rsProducts.Open cmd, , adOpenStatic, adLockOptimistic
  bMustClose = True
  i = 1
  Do While Not rsProducts.EOF And i <= MaxResult
  Set objDisplay = New d_Product
    With objDisplay
      .PID = rsProducts!P_ID
      .Code = FNS(rsProducts!P_Code)
      .Title = FNS(rsProducts!P_Title)
      .Author = FNS(rsProducts!P_MainAuthor)
      .Price = FNS(rsProducts!P_Price2)
      .Stock = FNS(rsProducts!P_QtyOnHand)
      colDisplay.Add objDisplay
      Set objDisplay = Nothing
      rsProducts.MoveNext
      i = i + 1
    End With
  Loop
MyExit:
  Set cmd = Nothing
  bMustClose = True
  rsProducts.Close
  Set rsProducts = Nothing
  Exit Sub
ErrorHandler:
    If bMustClose Then
        rsProducts.Close
        Set rsProducts = Nothing
    End If
  MsgBox "Error: " & oPC.CO.errors.Item(0), vbOKOnly, "ERROR!"
GoTo MyExit
End Sub

