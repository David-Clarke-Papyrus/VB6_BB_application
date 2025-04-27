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
Dim rsSalesPers As ADODB.Recordset
'Dim lngOPSessionID As Variant
Dim lSalesPersID As Long

Public Sub LoadSPers()
    On Error GoTo errHandler
    Set rsSalesPers = New ADODB.Recordset
    With rsSalesPers
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .Open "SELECT * FROM tStaffMembers", oPC.DBLocalConn, adOpenForwardOnly, adLockReadOnly
        If Not .EOF Then Set .ActiveConnection = Nothing
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GetData.LoadSPers"
End Sub


Public Property Get DBCOnnection() As z_POSCLIConnection
    Set DBCOnnection = oPC
End Property
Public Function CloseDB() As Boolean
    On Error GoTo errHandler
    If Not oPC Is Nothing Then
        CloseDB = (oPC.dbCloseLocalConnect = 0)
        Set oPC = Nothing
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GetData.CloseDB"
End Function

Public Function SalesPersonList() As ADODB.Recordset
    On Error GoTo errHandler
  LoadSPers
  Set SalesPersonList = rsSalesPers
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GetData.SalesPersonList"
End Function

Public Function GetCustomer(Optional CustCode As String, Optional Address As String, Optional Phone As String) As ADODB.Recordset
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim sSQL As String
Dim sWhere As String
    
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
        .Open sSQL, oPC.DBLocalConn, adOpenDynamic, adLockBatchOptimistic
        
        If Not rs.EOF Then
            Set GetCustomer = rs
        End If
    End With
MEX:
    Set rs = Nothing
'EH:
'    Debug.Print "Error in GetCustomer: " & Error
'    GoTo MEX
'
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GetData.GetCustomer(CustCode,Address,Phone)", Array(CustCode, Address, Phone)
End Function

Public Function CheckPassword(sUserName As String, sPassword As String) As Boolean
    On Error GoTo errHandler
Dim iOff As Integer
Dim sName As String
    With rsSalesPers
        iOff = InStr(sUserName, "(")
        sName = Trim$(Left$(sUserName, iOff - 1))
        .MoveFirst
         Do While Not .EOF
            
            
            If NZS(!SM_Name) = sName And NZS(!SM_Password) = sPassword Then
                CheckPassword = True
                lSalesPersID = !SM_ID
                GoTo MEX
            End If
            .MoveNext
        Loop
MEX:
        .MoveFirst
    End With
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GetData.CheckPassword(sUserName,sPassword)", Array(sUserName, sPassword)
End Function

Public Property Get SalesPersonID() As Long
    On Error GoTo errHandler
    SalesPersonID = lSalesPersID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GetData.SalesPersonID"
End Property



Public Function IsValidZOpPass(Pass As String, ZA_level As Integer) As Boolean
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim sSQL As String
        
    sSQL = "SELECT * FROM tblStaffMembers WHERE tblStaffMembers.SM_Password = '" & Pass & "'"
    rs.Open sSQL, oPC.DBLocalConn, adOpenDynamic, adLockReadOnly
    If Not rs.EOF Then
        'true if level >= 2
        IsValidZOpPass = (rs!SM_Level >= ZA_level)
    End If
    
    
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
'EH:
'    MsgBox "Can't access Database!"
'    GoTo MEX
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GetData.IsValidZOpPass(Pass,ZA_level)", Array(Pass, ZA_level)
End Function

Friend Function GetNextFileNum() As Long
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim sSQL As String

    'sSQL = "SELECT Max([SaleFileNum]) AS MaxNum FROM tblSaleFileNum"
    sSQL = "SELECT * FROM tblSaleFileNum"
    rs.Open sSQL, oPC.DBLocalConn, adOpenDynamic, adLockOptimistic
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GetData.GetNextFileNum"
End Function

Public Function ClearStaffMembers()
On Error GoTo errHandler
Dim strSQL As String
    strSQL = "DELETE * FROM tblStaffMembers"
    oPC.DBLocalConn.Execute strSQL
    ClearStaffMembers = True
    Exit Function
errHandler:
    ErrPreserve
    ClearStaffMembers = False
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GetData.GetNextFileNum"
End Function

Public Function GetProduct(pCODE As String) As ADODB.Recordset
On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oProdCode As New z_ProdCode_Short
Dim sTmp As String
Dim bFound As Boolean

    oProdCode.Load pCODE
    If Not oProdCode.CanHandle Then
        Exit Function
    End If
    bFound = False
    If oProdCode.IsEAN Then
        sTmp = "SELECT * FROM tProduct WHERE P_EAN ='" & oProdCode.EAN & "'"
      '  sTmp = "SELECT tProduct.*, FROM tProduct LEFT JOIN tMarketing ON P_PTID = M_PTID and P_SECID = M_SECTION_ID WHERE P_EAN ='" & oProdCode.EAN & "'"
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenKeyset
        rs.LockType = adLockReadOnly
        rs.Open sTmp, oPC.DBLocalConn, adOpenForwardOnly, adLockReadOnly
        bFound = Not rs.EOF
        If Not bFound Then
            rs.Close
        Else
            Set GetProduct = rs
        End If
    End If
    If Not bFound Then
        If oProdCode.ISBNExists Or oProdCode.IsCode Then
            sTmp = "SELECT * FROM tProduct WHERE P_Code ='" & oProdCode.Code & "'"
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.CursorType = adOpenKeyset
            rs.LockType = adLockReadOnly
            rs.Open sTmp, oPC.DBLocalConn, adOpenForwardOnly, adLockReadOnly
            bFound = Not rs.EOF
            If Not bFound Then
                rs.Close
            Else
                Set GetProduct = rs
            End If
        End If
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GetData.GetProduct(PCode)", pCODE
End Function


