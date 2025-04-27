VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_QuickProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim oProdCode As New z_ProdCode
Dim rs As ADODB.Recordset
Dim lResult As Long
Dim lCSLID As Long
Dim lngVATRATE As Long
Dim strTimer As String
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim stror As String
Dim str As String

Function HandleCode(pPID As String, pCSID As Long, pCode As String, pTitle As String, PPrice As Long, pCSLID As Long, pCodeF As String, pPriceF As String) As Integer
    On Error GoTo errHandler
Dim bFound As Boolean
Dim oProd As a_Product
Dim strPos As String
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

TOP:

Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set oProd = New a_Product
    If oProd.Load("", 0, Trim(pCode)) = 0 Then   'product  found
        pCode = oProd.code    'FNS(rs.Fields("P_Code"))
        pTitle = oProd.Title     'FNS(rs.Fields("P_Title"))
        PPrice = oProd.SP   'FNN(rs.Fields("P_SP"))
        pPriceF = oProd.SPF
        pPID = oProd.PID    'rs.Fields("P_ID")
        pCodeF = oProd.CodeF     'FNS(rs.Fields("CodeF"))
        lngVATRATE = oProd.VATRate          'FNN(rs.Fields("P_VATRATE"))
        Set cmd = Nothing
        Set cmd = New ADODB.Command
        cmd.CommandTimeout = 0
        Set cmd.ActiveConnection = oPC.COShort
        cmd.CommandText = "sp_CSL"
        cmd.CommandType = adCmdStoredProc

    '1
        strPos = strPos & "a"

            Set prm = cmd.CreateParameter("@TRID", adInteger, adParamInput, , pCSID)
            cmd.Parameters.Append prm
            Set prm = Nothing
    '2
        strPos = strPos & "b"
            Set prm = cmd.CreateParameter("@PID", adGUID, adParamInput, , pPID)
            cmd.Parameters.Append prm
            Set prm = Nothing
    '3
        strPos = strPos & "c"
            Set prm = cmd.CreateParameter("@Qty", adInteger, adParamInput, , 1)
            cmd.Parameters.Append prm
            Set prm = Nothing
    '4
        strPos = strPos & "d"
            Set prm = cmd.CreateParameter("@Price", adInteger, adParamInput, , PPrice)
            cmd.Parameters.Append prm
            Set prm = Nothing
    '5
            Set prm = cmd.CreateParameter("@PriceAlteration", adInteger, adParamInput, , 0)
            cmd.Parameters.Append prm
            Set prm = Nothing

         strPos = strPos & "e"
           Set prm = cmd.CreateParameter("@DiscountRate", adNumeric, adParamInput, , 0)
            prm.NumericScale = 5
            prm.Precision = 5
            cmd.Parameters.Append prm
            Set prm = Nothing
         strPos = strPos & "f"

            Set prm = cmd.CreateParameter("@Discount", adInteger, adParamInput, , 0)
            cmd.Parameters.Append prm
            Set prm = Nothing

        strPos = strPos & "g"
            Set prm = cmd.CreateParameter("@VRate", adInteger, adParamInput, , CInt(lngVATRATE))
            cmd.Parameters.Append prm
            Set prm = Nothing

        strPos = strPos & "h"
            Set prm = cmd.CreateParameter("@STID ", adInteger, adParamInput, , oPC.Configuration.DefaultStoreID)
            cmd.Parameters.Append prm
            Set prm = Nothing

        strPos = strPos & "i"
            Set prm = cmd.CreateParameter("@TPID ", adInteger, adParamInput, , oPC.Configuration.CSCustomerID)
            cmd.Parameters.Append prm
            Set prm = Nothing

        strPos = strPos & "j"
            Set prm = cmd.CreateParameter("@EXCHID ", adGUID, adParamInput, , Null)
            cmd.Parameters.Append prm
            Set prm = Nothing

        strPos = strPos & "j2"
            Set prm = cmd.CreateParameter("@COUNTERFOIL", adVarChar, adParamInput, 30, "")
            cmd.Parameters.Append prm
            Set prm = Nothing

        strPos = strPos & "j3"
            Set prm = cmd.CreateParameter("@DISCOUNTDESCRIPTION", adVarChar, adParamInput, 20, "")
            cmd.Parameters.Append prm
            Set prm = Nothing

        strPos = strPos & "k"
            Set prm = cmd.CreateParameter("@ID", adInteger, adParamOutput)
            cmd.Parameters.Append prm
            Set prm = Nothing

        strPos = strPos & "l"
            Set prm = cmd.CreateParameter("@pCode", adInteger, adParamOutput)
            cmd.Parameters.Append prm
            Set prm = Nothing

        strPos = strPos & "m"
            Set prm = cmd.CreateParameter("@POSITION", adInteger, adParamOutput)
            cmd.Parameters.Append prm
            Set prm = Nothing

            strPos = strPos & "  Before Execute"

            cmd.execute
            strPos = strPos & "  After Execute"

          '  Check cmd.Parameters("@pCode") = 0, EXC_GENERAL, "or in Stored procedure sp_CSL or code is " & cmd.Parameters("@pCode") & " in position: " & cmd.Parameters("@POSITION")

        If FNN(cmd.Parameters("@ID")) <> 0 Then
            strPos = strPos & "  14"
            pCSLID = FNN(cmd.Parameters("@ID"))
            strPos = strPos & "  15"
        End If
        Set cmd = Nothing
       ' rs.Close
        HandleCode = 0
    Else
        HandleCode = 99
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_QuickProduct.HandleCode(pPID,pCSID,pCODE,pTitle,PPrice,pCSLID,pCodeF)", _
         Array(pPID, pCSID, pCode, pTitle, PPrice, pCSLID, pCodeF), , , "strPOS", Array(strPos)
End Function

Private Sub Class_Initialize()
    On Error GoTo errHandler
 Set rs = New ADODB.Recordset
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_QuickProduct.Class_Initialize"
End Sub
Public Function DeleteRow(pCSL As Long)
Dim OpenResult As Integer
    On Error GoTo errHandler
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    Set cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_CancelSaleLine"
    cmd.CommandType = adCmdStoredProc

'1
    Set prm = New ADODB.Parameter
    prm.Type = adInteger
    prm.Direction = adParamInput
    prm.Value = pCSL
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = Nothing
    
'2
    Set prm = Nothing
    Set prm = New ADODB.Parameter
    prm.Type = adInteger
    prm.Direction = adParamInput
    prm.Value = oPC.Configuration.DefaultStoreID
    cmd.Parameters.Append prm
    Set prm = Nothing
''-------------------------------
    OpenResult = oPC.OpenDBSHort
''-------------------------------
    cmd.execute
''---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_QuickProduct.DeleteRow(pCSL)", pCSL
End Function

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set rs = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_QuickProduct.Class_Terminate"
End Sub

'Function QuickProduct(pString As String, pPID As String, pEAN As String, pCODE As String, Optional pTitle As String, Optional PPrice As Long, Optional pCodeF As String) As Integer
'    On Error GoTo ErrHandler
'Dim bFound As Boolean
'Dim oProd As a_Product
'Dim strPos As String
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'Dim OpenResult As Integer
'Top:
'    QuickProduct = 0
' '   oProdCode.Load pString, True 'Assume it is a book
' '   If Not oProdCode.CanHandle Then
' '       GoTo EXIT_HANDLER
' '   End If
'    bFound = False
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    If Len(FNS(pEAN)) = 13 Then
'        Set cmd = New ADODB.Command
'        Set par = cmd.CreateParameter
'        par.Type = adVarChar
'        par.Size = 20
'        par.Value = oProdCode.EAN
'        cmd.Parameters.Append par
'        cmd.CommandText = "q_GetPIDbyEAN"
'        cmd.CommandType = adCmdStoredProc
'
'        cmd.ActiveConnection = oPC.COShort
'        Set rs = cmd.execute
'        Set cmd = Nothing
'        bFound = Not rs.EOF
'        If Not bFound Then
'            rs.Close
'        End If
'
'    End If
'    If Not bFound Then
'        Set cmd = New ADODB.Command
'        Set par = cmd.CreateParameter
'        par.Type = adVarChar
'        par.Size = 20
'        par.Value = oProdCode.Code
'        cmd.Parameters.Append par
'        cmd.CommandText = "q_GetPIDbyCode"
'        cmd.CommandType = adCmdStoredProc
'
'        cmd.ActiveConnection = oPC.COShort
'        Set rs = cmd.execute
'        Set cmd = Nothing
'        bFound = Not rs.EOF
'        If Not bFound Then
'            rs.Close
'        End If
'    End If
'    If Not bFound Then
'        Set oProd = New a_Product
'        lResult = oProd.LoadNew(FNS(pCODE), FNS(pEAN)) 'loads the record from BF if not on database
'        If lResult = 0 Then
'            pCODE = oProd.Code
'            GoTo Top
'        Else
'            bFound = False
'        End If
'
'    End If
'    If bFound Then
'        pEAN = FNS(rs.Fields("P_EAN"))
'        pCODE = FNS(rs.Fields("P_Code"))
'        pTitle = FNS(rs.Fields("P_Title"))
'        PPrice = FNN(rs.Fields("P_SP"))
'        pPID = rs.Fields("P_ID")
'        pCodeF = FNS(rs.Fields("CodeF"))
'        lngVATRATE = FNN(rs.Fields("P_VATRATE"))
'        QuickProduct = 1
'    End If
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'EXIT_Handler:
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_QuickProduct.QuickProduct(pString,pPID,pEAN,pCODE,pTitle,PPrice,pCodeF)", _
'         Array(pString, pPID, pEAN, pCODE, pTitle, PPrice, pCodeF)
'End Function



