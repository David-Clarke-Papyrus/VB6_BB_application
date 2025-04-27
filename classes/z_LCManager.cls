VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_LCManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function InsertLoyaltyCustomer(pTPID As Long, pACNO As String, pLastName As String, pTitle As String, pFirstname As String, _
pInitials As String, pADDRESSEE As String, pADDL1 As String, pADDL2 As String, pADDL3 As String, pADDL4 As String, pADDL5 As String, pADDL6 As String, _
pCountryName As String, pPCode As String, pPhone1 As String, pPhone2 As String, pMobile As String, pEmail As String, _
pBooksale As Boolean, pBookpromotion As Boolean, pBookLaunch As Boolean, pStoreCode As String) As Long   ', pIDNUM As String
    On Error GoTo errHandler

Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

Dim OpenResult As Integer
Dim prm As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_InsertLoyaltyCustomer"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@ACNO", adVarChar, adParamInput, 10)
    cmd.Parameters.Append par
    par.Value = Left(pACNO, 10)
    Set par = cmd.CreateParameter("@Lastname", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = Left(pLastName, 50)
    Set par = cmd.CreateParameter("@Title", adVarChar, adParamInput, 20)
    cmd.Parameters.Append par
    par.Value = Left(pTitle, 20)
    Set par = cmd.CreateParameter("@Firstname", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = Left(pFirstname, 50)
    Set par = cmd.CreateParameter("@Initials", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = Left(pInitials, 50)
    Set par = cmd.CreateParameter("@ADDRESSEE", adVarChar, adParamInput, 150)
    cmd.Parameters.Append par
    par.Value = Left(pADDRESSEE, 150)
    Set par = cmd.CreateParameter("@ADDL1", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = Left(pADDL1, 50)
    Set par = cmd.CreateParameter("@ADDL2", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = Left(pADDL2, 50)
    Set par = cmd.CreateParameter("@ADDL3", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = Left(pADDL3, 50)
    Set par = cmd.CreateParameter("@ADDL4", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = Left(pADDL4, 50)
    Set par = cmd.CreateParameter("@ADDL5", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = Left(pADDL5, 50)
    Set par = cmd.CreateParameter("@ADDL6", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = Left(pADDL6, 50)
    Set par = cmd.CreateParameter("@CountryName", adVarChar, adParamInput, 30)
    cmd.Parameters.Append par
    par.Value = Left(pCountryName, 30)
    Set par = cmd.CreateParameter("@PCode", adVarChar, adParamInput, 10)
    cmd.Parameters.Append par
    par.Value = Left(pPCode, 10)
    Set par = cmd.CreateParameter("@Phone1", adVarChar, adParamInput, 20)
    cmd.Parameters.Append par
    par.Value = Left(Replace(pPhone1, " ", ""), 20)
    Set par = cmd.CreateParameter("@Phone2", adVarChar, adParamInput, 20)
    cmd.Parameters.Append par
    par.Value = Left(Replace(pPhone2, " ", ""), 20)
    Set par = cmd.CreateParameter("@Mobile", adVarChar, adParamInput, 20)
    cmd.Parameters.Append par
    par.Value = Left(Replace(pMobile, " ", ""), 20)
    Set par = cmd.CreateParameter("@Email", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = Left(pEmail, 50)
    Set par = cmd.CreateParameter("@Booksale", adBoolean, adParamInput)
    cmd.Parameters.Append par
    par.Value = IIf(pBooksale, 1, 0)
    Set par = cmd.CreateParameter("@Bookpromotion", adBoolean, adParamInput)
    cmd.Parameters.Append par
    par.Value = IIf(pBookpromotion, 1, 0)
    Set par = cmd.CreateParameter("@Booklaunch", adBoolean, adParamInput)
    cmd.Parameters.Append par
    par.Value = IIf(pBookLaunch, 1, 0)
    
    Set par = cmd.CreateParameter("@Source", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = pStoreCode
    Set par = cmd.CreateParameter("@pErrCode", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.Execute
    If IsNull(cmd.Parameters("@TPID").Value) Then
        pTPID = 0
    Else
        pTPID = cmd.Parameters("@TPID").Value
    End If
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Function
errHandler:
    ErrorIn "z_StockManager.InsertLoyaltyCustomer(pTPID,pACNO,pLastName,pTitle,pFirstName,pAddressee,pADDL1," & _
        "pADDL2,pADDL3,pADDL4,pADDL5,pADDL6,pCountryName,pPCode,pPhone1,pPhone2,pMobile,pEmail,pBooksale," & _
        "pBookpromotion,pBookLaunch)", Array(pTPID, pACNO, pLastName, pTitle, pFirstname, pADDL1, pADDL2, _
         pADDL3, pADDL4, pADDL5, pADDL6, pCountryName, pPCode, pPhone1, pPhone2, pMobile, pEmail, pBooksale, _
         pBookpromotion, pBookLaunch)
    Exit Function
    Resume
End Function

Public Function InsertLoyaltyCustomerSale(pACNO As String, pCode As String, pSP As Long, _
pPrice As Long, pQty As Long, pDate As Date, rRegNO As Long, pTrackingNo As Long, pSuccess As Long) As Long
    On Error GoTo errHandler

Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_InsertLoyaltyCustomerSaleLine"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@ACNO", adVarChar, adParamInput, 10)
    cmd.Parameters.Append par
    par.Value = pACNO
    Set par = cmd.CreateParameter("@pCode", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = pCode
    Set par = cmd.CreateParameter("@pSP", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = pSP
    Set par = cmd.CreateParameter("@pPrice", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = pPrice
    Set par = cmd.CreateParameter("@pQty", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = pQty
    Set par = cmd.CreateParameter("@pDate", adDate, adParamInput)
    cmd.Parameters.Append par
    par.Value = pDate
    Set par = cmd.CreateParameter("@pRegNo", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = rRegNO
    Set par = cmd.CreateParameter("@pTrackingNo", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = pTrackingNo
    Set par = cmd.CreateParameter("@pSuccess", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.Execute
    pSuccess = cmd.Parameters("@pSuccess").Value
    Set cmd = Nothing
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.InsertLoyaltyCustomerSale(pACNO,pCode,pSP,pPrice,pQty,pDate,rRegNO," & _
        "pTrackingNo)", Array(pACNO, pCode, pSP, pPrice, pQty, pDate, rRegNO, pTrackingNo)
End Function


