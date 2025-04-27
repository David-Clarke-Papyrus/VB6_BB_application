VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Lookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    

Public Function GetProduct(pEAN As String, pCODE As String, pCTSet As String, Optional bSuccess As Boolean) As ADODB.Recordset
    On Error GoTo errHandler
10        On Error GoTo errHandler
      Dim iResult As Long
      Dim cmd As ADODB.Command
      Dim par As ADODB.Parameter
      Dim strPos As String
      Dim rs As New ADODB.Recordset
      Dim RetryCount As Integer

      '-------------------------------
20        oPC.OpenLocalDatabase
      '-------------------------------
30        RetryCount = 0
TOP:
40        Set cmd = New ADODB.Command
50        cmd.CommandText = "sp_GetProduct"
60        cmd.CommandType = adCmdStoredProc
70        Set par = cmd.CreateParameter("@EAN", adVarChar, adParamInput, 15, Left(pEAN, 13))
80        cmd.Parameters.Append par
90        Set par = cmd.CreateParameter("@CODE", adVarChar, adParamInput, 15, Left(pCODE, 13))
100       cmd.Parameters.Append par
110       Set par = cmd.CreateParameter("@CTID", adVarChar, adParamInput, 100, pCTSet)
120       cmd.Parameters.Append par
          
130       cmd.ActiveConnection = oPC.ServerConnectionRef
140       rs.CursorLocation = adUseClient
150       rs.Open cmd, , adOpenStatic, adLockOptimistic
160       Set GetProduct = rs

      'Here we have a recordset of all possible rules (including the null rule) that apply to this product/customer
      'We need to determine which to use. It could vary depending on the total value of the sale
  '    MsgBox "pos 1"
240       bSuccess = (rs.RecordCount > 0)
170       If rs.RecordCount = 0 And RetryCount < 2 Then
180           If GetProductFromMasterDB(pEAN, RetryCount) = False Then
                RetryCount = 99
  '               MsgBox "pos 2"
              Else
                RetryCount = RetryCount + 1
200             rs.Close
210             GoTo TOP
              End If
220       End If
230       Set cmd = Nothing

250       Exit Function
errHandler:
260       If ErrMustStop Then Debug.Assert False: Resume
270       ErrorIn "z_Lookup.GetProduct(pEAN,pCODE,pCTSet)", Array(pEAN, pCODE, pCTSet), , , "Line number", Array(Erl())
End Function

Private Function GetProductFromMasterDB(pArg As String, RetryCount As Integer) As Boolean

    On Error GoTo errHandler
Dim mPID As String
Dim mCode As String
Dim mEAN As String
Dim mPublisher As String
Dim mSeriesTitle As String
Dim mAuthor As String
Dim mTitle As String
Dim mSP As Long
Dim mSSP As Long
Dim mVATRate As Double
Dim mPTID As Long
Dim mSECID As Long
Dim mNDA As Boolean
    
    If Not IsISBN13(pArg) Then
        GetProductFromMasterDB = False
        Exit Function
    End If

    oPC.dbConnectMain
    If bIsConnectedtoServer = True Then
        If RetryCount = 1 Then
            oPC.DBLocalConn.Execute "INSERT INTO PBKSFD.dbo.tPRODUCT (P_ID,P_CODE,P_EAN,P_TITLE,P_PUBLISHER,P_SERIESTITLE,P_MAINAUTHOR,P_SAPRICE,P_SPPRICE,P_VATRATE,P_PTID,P_SECID,P_NDA,P_MultibuyCode) " _
                    & " Select A.* From OPENROWSET('MSDASQL','Driver={SQL SERVER};SERVER=" & oPC.MainSQLServerName & ";UID=sa;PWD=car', " _
                    & " 'SELECT P_ID,P_CODE,P_EAN,P_TITLE,P_PUBLISHER,P_SERIESTITLE,P_MAINAUTHOR,P_SP,P_SPECIAL,P_VATRATE,P_PRODUCTTYPE_ID,PBKS.dbo.[MainSectionID](P_ID),P_NDA,PBKS.dbo.vMultibuyCode.DICT_System " _
                    & " FROM PBKS.dbo.tPRODUCT LEFT JOIN PBKS.dbo.vMultibuyCode ON P_ID = PBKS.dbo.vMultibuyCode.PSEC_P_ID WHERE PBKS.dbo.tPRODUCT.P_CODE = ''" & pArg & "''') as A"
        Else
            oPC.DBLocalConn.Execute "INSERT INTO PBKSFD.dbo.tPRODUCT (P_ID,P_CODE,P_EAN,P_TITLE,P_PUBLISHER,P_SERIESTITLE,P_MAINAUTHOR,P_SAPRICE,P_SPPRICE,P_VATRATE,P_PTID,P_SECID,P_NDA,P_MultibuyCode) " _
                    & " Select A.* From OPENROWSET('MSDASQL','Driver={SQL SERVER};SERVER=" & oPC.MainSQLServerName & ";UID=sa;PWD=car', " _
                    & " 'SELECT P_ID,P_CODE,P_EAN,P_TITLE,P_PUBLISHER,P_SERIESTITLE,P_MAINAUTHOR,P_SP,P_SPECIAL,P_VATRATE,P_PRODUCTTYPE_ID,PBKS.dbo.[MainSectionID](P_ID),P_NDA,PBKS.dbo.vMultibuyCode.DICT_System " _
                    & " FROM PBKS.dbo.tPRODUCT LEFT JOIN PBKS.dbo.vMultibuyCode ON P_ID = PBKS.dbo.vMultibuyCode.PSEC_P_ID WHERE PBKS.dbo.tPRODUCT.P_EAN = ''" & pArg & "''') as A"
        End If
        oPC.dbMainDisConnect
        GetProductFromMasterDB = True
    Else
        GetProductFromMasterDB = False
    End If
            
MEX:
    Exit Function
errHandler:
    ErrPreserve
    GetProductFromMasterDB = False

    MsgBox "Please ask Papyrus support to enable ad-hoc distributed queries", vbInformation, "Ad-hoc distributed queries disabled"
    If ErrMustStop Then Debug.Assert False: Resume
    ErrSaveToFile
    GoTo MEX
End Function
Public Function GetAlertFromMasterDB(TPACNO As String) As ADODB.Recordset
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset

    oPC.dbConnectMain
    If bIsConnectedtoServer = True Then
    rs.CursorLocation = adUseClient
        rs.Open "SELECT * FROM tALERT WHERE AL_TPACNO = '" & TPACNO & "' AND AL_READSTATUS = 'U'", oPC.DBMainConn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Do While Not rs.EOF
                rs.Fields("AL_READSTATUS") = "R"
                rs.Update
                rs.MoveNext
            Loop
            rs.MoveFirst
        End If
        Set rs.ActiveConnection = Nothing
        Set GetAlertFromMasterDB = rs
        oPC.dbMainDisConnect
    End If
            
MEX:
    Exit Function
errHandler:
    ErrorIn "z_Lookup.GetAlertFromMasterDB(TPACNO)", TPACNO
    Exit Function
End Function
Public Sub GetCustomerFromMasterDB(pArg As String, RecordsAffected As Integer)
    On Error GoTo errHandler

    oPC.dbConnectMain
    If bIsConnectedtoServer = True Then
            oPC.DBMainConn.Execute "INSERT INTO PBKSFD.dbo.tCustomer (Customer_ID,C_NAME,C_PHONE,C_IDNUMBER,C_ACNO,C_DEFAULTDISCOUNT,C_VATABLE,C_INITIALS,C_TITLE,C_TYPE,C_BALANCE,C_BALANCES,C_CREDITLIMIT,C_TERMS) " _
                    & " SELECT TPID,NAME,PHONE,IDNUMBER,ACNO,DEFAULTDISCOUNT,VATABLE,INITIALS,TITLE,TYPE,BALANCE,BALANCES,CREDITLIMIT,TERMS " _
                    & " FROM PBKS.dbo.vGetCustomerForPOS WHERE PBKS.dbo.vGetCustomerForPOS.ACNO = '" & Replace(pArg, "'", "''") & "'", RecordsAffected
        oPC.dbMainDisConnect
    End If

MEX:
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrSaveToFile
    GoTo MEX
End Sub


