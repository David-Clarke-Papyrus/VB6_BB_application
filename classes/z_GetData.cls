VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_GetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private Type tSalesPers
'    ID As Long
'    Name As String * 30
'    Pass As String * 10
'End Type
'
'Dim udtSalePers As tSalesPers

'Dim cSalesPers As Collection
Dim rsSalesPers As ADODB.Recordset

Dim lSalesPersID As Long

Private Sub LoadSPers()
    'On Error Resume Next
    Set rsSalesPers = New ADODB.Recordset
    With rsSalesPers
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .Open "SELECT * FROM tblSalesPerson", gDB.DBConn, adOpenForwardOnly, adLockReadOnly
        If Not .EOF Then Set .ActiveConnection = Nothing
    End With
        
End Sub
Public Function GetProductByISBN(sISBN As String) As ADODB.Recordset
Dim rs As ADODB.Recordset

Dim sTmp As String

    
    
    sTmp = "SELECT * FROM Product WHERE P_ISBN ='" & sISBN & "'"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockReadOnly
    rs.Open sTmp, gDB.DBConn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
'        GetProductByISBN = NZS(rs!P_ISBN) & ";" & NZS(rs!P_Title) & ";" & NZS(rs!P_MainAuthor) & ";" & NZ(rs!P_SAPrice)
         Set GetProductByISBN = rs
    End If
    
MEX:
'    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
Exit Function
EH:
    
   GoTo MEX
End Function

Public Function LoadDB() As Boolean
Dim iR As Integer
    If gDB Is Nothing Then
        Set gDB = New z_DatabasePersist
        iR = gDB.dbConnect("Admin", "") '"admin", "sru")
    End If
    LoadSPers
    LoadDB = (iR = 0)
    
End Function

Public Function CloseDB() As Boolean
    If Not gDB Is Nothing Then
        CloseDB = (gDB.dbCloseConnect = 0)
        Set gDB = Nothing
    End If
End Function

Private Sub Class_Initialize()
'    Set oDB = New z_DatabasePersist
End Sub

Private Sub Class_Terminate()
    CloseDB
End Sub

Public Function SalesPersonList() As ADODB.Recordset
    Set SalesPersonList = rsSalesPers
End Function

Public Function GetCustomer(Optional CustCode As String, Optional Address As String, Optional Phone As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sSQL As String
Dim sWhere As String
    
    On Error GoTo EH
    If Len(CustCode) > 0 Then
        sWhere = "tblCustomer.C_Code ='" & CustCode & "'"
    ElseIf Len(Address) > 0 Then
        sWhere = "tblCustomer.C_Address LIKE '" & Address & "%'"
    ElseIf Len(Phone) > 0 Then
        sWhere = "tblCustomer.C_Phone LIKE '" & Phone & "%'"
    Else
        Exit Function
    End If
    sSQL = "SELECT tblCustomer.* " & _
           "FROM tblCustomer WHERE " & sWhere
    With rs
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockBatchOptimistic
        .Open sSQL, gDB.DBConn, adOpenDynamic, adLockBatchOptimistic
        
        If Not rs.EOF Then
            Set GetCustomer = rs
        End If
    End With
MEX:
    Set rs = Nothing
    Exit Function
EH:
    Debug.Print "Error in GetCustomer: " & Error
    GoTo MEX
    
End Function

Public Function CheckPassword(sUserName As String, sPassword As String) As Boolean
    With rsSalesPers
        .MoveFirst
         Do While Not .EOF
            If NZS(!SP_Code) = sUserName And NZS(!SP_Pass) = sPassword Then
                CheckPassword = True
                lSalesPersID = !SalesPerson_ID
                GoTo MEX
            End If
            .MoveNext
        Loop
MEX:
        .MoveFirst
    End With
End Function

Public Property Get SalesPersonID() As Long
    SalesPersonID = lSalesPersID
End Property



Public Function IsValidZOpPass(Pass As String) As Boolean
Dim rs As New ADODB.Recordset
Dim sSQL As String
        
    On Error GoTo EH
    sSQL = "SELECT * FROM tblSalesPerson WHERE tblSalesPerson.SP_Pass = '" & Pass & "'"
    rs.Open sSQL, gDB.DBConn, adOpenDynamic, adLockReadOnly
    If Not rs.EOF Then
        IsValidZOpPass = rs!SP_ZAllowed
    End If
    
    
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Exit Function
EH:
    MsgBox "Can't access Database!"
    GoTo MEX
End Function

Friend Function GetNextFileNum() As Long
Dim rs As New ADODB.Recordset
Dim sSQL As String

    'sSQL = "SELECT Max([SaleFileNum]) AS MaxNum FROM tblSaleFileNum"
    sSQL = "SELECT * FROM tblSaleFileNum"
    rs.Open sSQL, gDB.DBConn, adOpenDynamic, adLockOptimistic
    On Error Resume Next
    If rs.EOF Then
        rs.AddNew
        rs!SaleFileNum = 1
    Else
        rs!SaleFileNum = rs!SaleFileNum + 1
    End If
'        rs.AddNew
    GetNextFileNum = rs!SaleFileNum
    rs.Update
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Function
