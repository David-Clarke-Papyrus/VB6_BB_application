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
Attribute VB_Exposed = True
Option Explicit
    
    
'''Public Function GetProduct_Obsolete(pEAN As String, pCOde As String, pPID As String, pTitle As String, pAuthor As String, _
'''                        pVATRATE As Long, pPrice As Long, pDiscount As Long, pLoyaltyDiscount As Long, _
'''                        pIdentifyCustomer As Boolean, pNodiscountallowable As Boolean, pDiscountRule As String)
'''    On Error GoTo errHandler
''''On Error GoTo errHandler
'''Dim iresult As Long
'''Dim cmd As ADODB.Command
'''Dim par As ADODB.Parameter
'''Dim strPos As String
'''    Set cmd = New ADODB.Command
'''    cmd.CommandText = "sp_GetProduct"
'''    cmd.CommandType = adCmdStoredProc
'''strPos = "Pos 0"
'''
'''    Set par = cmd.CreateParameter("@EAN", adVarChar, adParamInputOutput, 15, Left(pEAN, 13))
'''    cmd.Parameters.Append par
'''    Set par = cmd.CreateParameter("@CODE", adVarChar, adParamInputOutput, 15, Left(pCOde, 13))
'''    cmd.Parameters.Append par
'''    Set par = cmd.CreateParameter("@PID", adGUID, adParamOutput, , pPID)
'''    cmd.Parameters.Append par
'''    Set par = cmd.CreateParameter("@TITLE", adVarChar, adParamOutput, 100, pTitle)
'''    cmd.Parameters.Append par
'''    Set par = cmd.CreateParameter("@AUTHOR", adVarChar, adParamOutput, 50, pAuthor)
'''    cmd.Parameters.Append par
'''strPos = "Pos 01"
'''    Set par = cmd.CreateParameter("@VATRATE", adInteger, adParamOutput, , pVATRATE)
'''    cmd.Parameters.Append par
'''    Set par = cmd.CreateParameter("@PRICE", adInteger, adParamOutput, , pPrice)
'''    cmd.Parameters.Append par
'''    Set par = cmd.CreateParameter("@DISCOUNT", adInteger, adParamOutput, , pDiscount)
'''    cmd.Parameters.Append par
'''strPos = "Pos 02"
'''    Set par = cmd.CreateParameter("@LOYALTYDISCOUNT", adInteger, adParamOutput, , pLoyaltyDiscount)
'''    cmd.Parameters.Append par
'''    Set par = cmd.CreateParameter("@IDENTIFYCUSTOMER", adBoolean, adParamOutput, , pIdentifyCustomer)
'''    cmd.Parameters.Append par
'''    Set par = cmd.CreateParameter("@NODISCOUNTALLOWABLE", adBoolean, adParamOutput, , pNodiscountallowable)
'''    cmd.Parameters.Append par
'''    Set par = cmd.CreateParameter("@DISCOUNTRULE", adVarChar, adParamOutput, 30, pDiscountRule)
'''    cmd.Parameters.Append par
'''    cmd.ActiveConnection = oPC.DBLocalConn
'''strPos = "Pos 1"
'''    cmd.Execute
'''strPos = "Pos 2"
'''    pEAN = FNS(cmd.Parameters(0).Value)
'''strPos = "Pos 3"
'''    pCOde = FNS(cmd.Parameters(1).Value)
'''strPos = "Pos 4"
'''    pPID = FNS(cmd.Parameters(2).Value)
'''strPos = "Pos 5"
'''
'''    pTitle = FNS(cmd.Parameters(3).Value)
'''strPos = "Pos 6"
'''    pAuthor = FNS(cmd.Parameters(4).Value)
'''    pVATRATE = FNN(cmd.Parameters(5).Value)
'''strPos = "Pos 7"
'''    pPrice = FNN(cmd.Parameters(6).Value)
'''    pDiscount = FNN(cmd.Parameters(7).Value)
'''strPos = "Pos 8"
'''    pLoyaltyDiscount = FNN(cmd.Parameters(8).Value)
'''    pIdentifyCustomer = FNB(cmd.Parameters(9).Value)
'''strPos = "Pos 9"
'''    pNodiscountallowable = FNB(cmd.Parameters(10).Value)
'''    pDiscountRule = FNS(cmd.Parameters(11).Value)
'''strPos = "Pos 10"
'''    Set cmd = Nothing
'''
'''
'''    Exit Function
'''errHandler:
'''    If ErrMustStop Then Debug.Assert False: Resume
'''    ErrorIn "z_Lookup.GetProduct(pEAN,pCode,pPID,pTitle,pAuthor,pVATRATE,pPrice,pDiscount," & _
'''        "pLoyaltyDiscount,pIdentifyCustomer,pNodiscountallowable)", Array(pEAN, pCOde, pPID, pTitle, pAuthor, _
'''         pVATRATE, pPrice, pDiscount, pLoyaltyDiscount, pIdentifyCustomer, pNodiscountallowable), , , "StrPos", Array(strPos)
'''End Function

Public Function GetProduct(pEAN As String, pCOde As String, pCT As String, Optional bSuccess As Boolean) As ADODB.Recordset
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strPos As String
Dim rs As New ADODB.Recordset

'-------------------------------
    oPC.OpenLocalDatabase
'-------------------------------
    
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_GetProduct"
    cmd.CommandType = adCmdStoredProc
    Set par = cmd.CreateParameter("@EAN", adVarChar, adParamInput, 15, Left(pEAN, 13))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@CODE", adVarChar, adParamInput, 15, Left(pCOde, 13))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@CTID", adVarChar, adParamInput, 5, pCT)
    cmd.Parameters.Append par
    
  '  MsgBox "oPC.ServerConnectionRef.ConnectionString=" & oPC.ServerConnectionRef.ConnectionString
  '  MsgBox "State = " & oPC.ServerConnectionRef.State
    cmd.ActiveConnection = oPC.ServerConnectionRef
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockOptimistic
    Set GetProduct = rs
'Here we have a recordset of all possible rules (including the null rule) that apply to this product/customer
'We need to determine which to use. It could vary depending on the total value of the sale
    Set cmd = Nothing
    bSuccess = True

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Lookup.GetProduct(pEAN,pCODE,pCT)", Array(pEAN, pCOde, pCT)
End Function

