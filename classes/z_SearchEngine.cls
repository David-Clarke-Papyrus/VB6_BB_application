VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_SearchEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim strSQL As String
Dim strcriteria As String
Dim lngretval As Long
Dim colobjs As Collection
Dim strstock As String
Dim intstock As Integer
Dim intWithCopies As Integer
Dim rs As ADODB.Recordset
Dim lngrowcount As Long
Dim inttemptbl As Integer
Dim test As New z_SQL
Dim intsearchcount As Integer
Private colDisplay As Collection
Public Enum enSearchType
    enSearchByCatalogue = 1
    enSearchNormal = 2
End Enum
Public Property Get rows() As Long
    rows = lngrowcount
End Property

Public Function instock(val As Integer)
    Select Case val
        Case 0
            intstock = 0
        Case 1
            intstock = 1
    End Select
End Function
Public Function WithCopies(val As Integer)
    Select Case val
        Case 0
            intWithCopies = 0
        Case 1
            intWithCopies = 1
    End Select
End Function

Public Function execute(MaxResult As Long)
           
    On Error GoTo ERRHANDLR
    
    Dim objDisplay As d_Product
    Dim i As Long
             
    Set colDisplay = Nothing
    Set colDisplay = New Collection
    Set test = New z_SQL
'    lngretval = test.RunSQL("EXEC sp_dboption 'PBKS','select into/bulkcopy','true'")
              
        
    lngretval = test.RunSQL(strSQL & strcriteria)
              
    If inttemptbl = 0 Then
        inttemptbl = 1
    End If
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenDynamic
    lngretval = test.RunGetRecordset("select P_ID, P_Code, P_Title, P_MainAuthor, P_Publisher, QtyOnHand from temppsearch" & inttemptbl & "", enText, Array(), "", rs)
               
    rowsreturned
    If lngrowcount > 0 Then
      i = 1
      rs.MoveFirst
      Do While Not rs.EOF And i <= MaxResult
        '  Set objdisplay = Nothing
          Set objDisplay = New d_Product
          With objDisplay
          .PID = rs!P_ID
          .Code = FNS(rs!P_Code)
          .Title = FNS(rs!P_Title)
          .Author = FNS(rs!P_MainAuthor)
          '.Price = FNN(rs!P_SAPrice)
          .Stock = FNN(rs!QtyOnHand)
          .Publisher = FNS(rs!P_Publisher)
          colDisplay.Add objDisplay '(.BookID)
          Set objDisplay = Nothing
          rs.MoveNext
          i = i + 1
          End With
    Loop
  End If
  rs.Close
  Set rs = Nothing
  
  
  Set test = Nothing
  
  Exit Function
                    
ERRHANDLR:
   MsgBox Error
   'Resume Next
   
 Exit Function
      Resume
End Function

Public Property Get getcols() As Collection

    Set getcols = colDisplay
    Set colDisplay = Nothing

End Property

Private Sub rowsreturned()
    
    On Error GoTo ERRHNDLER
    
    lngrowcount = 0
    
 '   rs.MoveFirst
    Do Until rs.EOF = True
        lngrowcount = lngrowcount + 1
        rs.MoveNext
    Loop
    
    'MsgBox "" & lngrowcount & " rows found", , "Search Complete!"
        
    Exit Sub
    
ERRHNDLER:
  
   MsgBox Error

End Sub

Private Sub droptemp()

    Set test = New z_SQL
    
    lngretval = test.DropTable("temppsearch1", "")
    lngretval = test.DropTable("temppsearch2", "")
    
    Set test = Nothing

End Sub

Public Function lookforobject(pAQ As Boolean, SearchType As enSearchType, IncludeAnt As Boolean)
On Error GoTo ERRH
    If intsearchcount = 0 Then
        droptemp
        If pAQ Then
            If SearchType = enSearchNormal Then
                If IncludeAnt Then
                    strSQL = "SELECT DISTINCT P_ID, P_Code, P_Title, P_MainAuthor, P_Publisher," _
                            & "P_QtyCopiesOnHand as QtyOnHand,  P_SupplierID " _
                            & "INTO temppsearch1 FROM tProduct Join tPI ON PI_P_ID = P_ID WHERE PI_ID is not null AND "
                Else
                    strSQL = "SELECT DISTINCT P_ID, P_Code, P_Title, P_MainAuthor, P_Publisher," _
                            & "P_QtyCopiesOnHand as QtyOnHand,  P_SupplierID " _
                            & "INTO temppsearch1 FROM tProduct WHERE "
                End If
            Else
                strSQL = "SELECT DISTINCT P_ID, P_Code, P_Title, P_MainAuthor, P_Publisher," _
                            & "P_QtyCopiesOnHand as QtyOnHand,  P_SupplierID " _
                            & "INTO temppsearch1 FROM vFindProductsbyCatAQ WHERE "
            End If
        Else
            strSQL = "SELECT DISTINCT P_ID, P_Code, P_Title, P_MainAuthor, P_Publisher," _
                        & "P_QtyOnHand as QtyOnHand,  P_SupplierID  INTO temppsearch1 FROM vFindProducts WHERE "
        End If
    Else
        Select Case intsearchcount
            Case 1
                inttemptbl = 2
                Set test = New z_SQL
                lngretval = test.DropTable("temppsearch2", "")
                strSQL = "SELECT temppsearch" & intsearchcount & ".P_ID," _
                 & "temppsearch" & intsearchcount & ".P_Code," _
                 & "temppsearch" & intsearchcount & ".P_Title," _
                 & "temppsearch" & intsearchcount & ".P_MainAuthor," _
                 & "temppsearch" & intsearchcount & ".P_Publisher," _
                 & "temppsearch" & intsearchcount & ".QtyOnHand," _
                 & "temppsearch" & intsearchcount & ".P_SupplierID " _
                 & "INTO temppsearch" & inttemptbl & " FROM temppsearch" & intsearchcount & " WHERE "
'            '     & "temppsearch" & intsearchcount & ".CAT_Description," _
               '             & "temppsearch" & intsearchcount & ".CAT_Description," _
                 & "temppsearch" & intsearchcount & ".P_CAT_ID," _
                '            & "temppsearch" & intsearchcount & ".P_CAT_ID,"
            Case 2
                inttemptbl = 1
                Set test = New z_SQL
                lngretval = test.DropTable("temppsearch1", "")
                strSQL = "SELECT temppsearch" & intsearchcount & ".P_ID," _
                            & "temppsearch" & intsearchcount & ".P_Code," _
                            & "temppsearch" & intsearchcount & ".P_Title," _
                            & "temppsearch" & intsearchcount & ".P_MainAuthor," _
                            & "temppsearch" & intsearchcount & ".P_Publisher," _
                            & "temppsearch" & intsearchcount & ".QtyOnHand," _
                            & "temppsearch" & intsearchcount & ".P_SupplierID " _
                            & "INTO temppsearch" & inttemptbl & " FROM temppsearch" & intsearchcount & " WHERE "
        End Select
        Set test = Nothing
    End If
    Exit Function
ERRH:
    MsgBox Error
    Exit Function
End Function

Public Sub selectcriteria(t As String, c As String)
    
    Select Case t
        Case "2-Code"
            strcriteria = "P_Code like '" & c & "'"
        Case "5-Author Contains"
            strcriteria = "P_MainAuthor like '%" & c & "%'"
        Case "4-Author Starts With"
            strcriteria = "P_MainAuthor like '" & c & "%'"
        Case "3-Title Contains"
            strcriteria = "P_Title like '%" & c & "%'"
        Case "1-Title Starts With"
            strcriteria = "P_Title like '" & c & "%'"
        Case "7-Publisher"
            strcriteria = "P_Publisher like '" & c & "%'"
        Case "9-BIC"
            strcriteria = "P_BIC like '" & c & "%' OR P_BIC LIKE '," & c & "%'  OR P_BIC LIKE '%," & c & ",%'"
        Case "8-Category"
            Select Case intsearchcount
                Case 0
                    strcriteria = "Category.Category_ID = " & c & ""
                    If intstock = 1 Then
                        strcriteria = strcriteria & " and Product.P_Stockbalance > 0"
                        intstock = 0
                    End If
                Case Else
                    strcriteria = "P_C_ID = " & c & ""
            End Select
        Case "6-Supplier"
            Select Case intsearchcount
                Case 0
                    strcriteria = "[Trading Partner].Trading_Partner_ID = " & c & ""
                    If intstock = 1 Then
                        strcriteria = strcriteria & " and P_QtyCopiesOnHand > 0"
                        intstock = 0
                    End If
                Case Else
                    strcriteria = "P_LastSupplierUsedID = " & c & ""
            End Select
        Case "Catalogue"
            Select Case intsearchcount
                Case 0
                    strcriteria = "CATAL_Serial = " & CInt(c) & ""
                    If intstock = 1 Then
                        strcriteria = strcriteria & " and P_QtyCopiesOnHand > 0"
                        intstock = 0
                    End If
                Case Else
                    strcriteria = "CAT_SerialNum = " & c & ""
            End Select
        End Select
        
        If intstock = 1 Then
            strcriteria = strcriteria & " and P_QtyCopiesOnHand > 0"
        End If

End Sub

Public Sub prisearch()

    intsearchcount = 0
    inttemptbl = 0

End Sub

Public Sub secsearch()

    Select Case intsearchcount
        Case 0
            intsearchcount = 1
        Case 1
            intsearchcount = 2
        Case 2
            intsearchcount = 1
    End Select

End Sub

Private Sub Class_Initialize()
Debug.Print "Initialize SearchEngine"
    addcols
   
End Sub

Private Sub Class_Terminate()
Debug.Print "Terminate SearchEngine"
    Set colobjs = Nothing
    Set test = Nothing
    Set colDisplay = Nothing

End Sub

Public Function addcols() As Collection

    Set colobjs = New Collection
    
    With colobjs
        .Add ("1-Title Starts With")
        .Add ("2-Code")
        .Add ("3-Title Contains")
        .Add ("4-Author Starts With")
        .Add ("5-Author Contains")
        .Add ("6-Supplier")
        .Add ("7-Publisher")
        .Add ("8-Category")
        .Add ("9-BIC")
    End With
    Set addcols = colobjs
    Set colobjs = Nothing

End Function




