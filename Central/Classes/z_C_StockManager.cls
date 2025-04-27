VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_StockManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Event eMax(Max As Long)
Event eProgress(Progress As Long)

Public Sub SendCustomerChanges()
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "_SendCQ"
    cmd.CommandTimeout = 0
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.COShort
    
    Set par = cmd.CreateParameter("@RES", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@ERRMSG", adVarChar, adParamOutput, 200)
    cmd.Parameters.Append par
    
    cmd.execute
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Set cmd = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.SendCustomerChanges"
End Sub

Public Function InsertScript(guid As String, Optional XMLScript As String, Optional FreeScript As String)
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.Open "SELECT * FROM tScripts", oPC.COShort, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs.Fields(1) = XMLScript
    rs.Update
    rs.Bookmark = rs.Bookmark
    guid = rs.Fields(0)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.InsertScript(guid,XMLScript,FreeScript)", Array(guid, XMLScript, _
         FreeScript)
End Function


Public Function GetSalesPatterns(pFrom As Date, pTo As Date, rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "SalesPattern_2"
    cmd.CommandTimeout = 0
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.COShort
    Set par = cmd.CreateParameter("@FROM", adDate, adParamInput, , pFrom)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adDate, adParamInput, , pTo)
    cmd.Parameters.Append par
    rs.Open cmd, , adOpenKeyset, adLockOptimistic
    
  '  rs.Close
    Set rs.ActiveConnection = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Set cmd = Nothing

    Set GetSalesPatterns = rs
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.GetSalesPatterns(pFrom,pTo,rs)", Array(pFrom, pTo, rs)
End Function

Public Function GetSalesPerformance(pFrom As Date, pTo As Date, rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "SalesPerformance_1"
    cmd.CommandTimeout = 0
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.COShort
    Set par = cmd.CreateParameter("@FROM", adDate, adParamInput, , pFrom)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adDate, adParamInput, , pTo)
    cmd.Parameters.Append par
    rs.Open cmd, , adOpenKeyset, adLockOptimistic
    
  '  rs.Close
    Set rs.ActiveConnection = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Set cmd = Nothing

    Set GetSalesPerformance = rs
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.GetSalesPerformance(pFrom,pTo,rs)", Array(pFrom, pTo, rs)
End Function


Public Function CustomerIndexClashes(lngTPID As Long, strACNo As String) As Boolean
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_CustomerClash"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = lngTPID
    Set par = cmd.CreateParameter("@ACNO", adVarChar, adParamInput, 20)
    cmd.Parameters.Append par
 '   If strACNo = "" Then
        par.Value = strACNo
 '   End If
    Set par = cmd.CreateParameter("@R", adInteger, adParamOutput)
    cmd.Parameters.Append par
    par.Value = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    CustomerIndexClashes = cmd.Parameters(2)
    Set cmd = Nothing

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------


End Function

Public Sub MergePTs(pPTOld As Long, pPTNew As Long)
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "Update tProduct SET P_PRODUCTTYPE_ID = " & pPTNew & " WHERE P_PRODUCTTYPE_ID = " & pPTOld
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.MergePTs(pPTOld,pPTNew)", Array(pPTOld, pPTNew), , oPC.COShort
End Sub
Public Sub DeleteUnusedPTs()
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "DELETE dbo.tPT FROM dbo.tPT LEFT OUTER JOIN dbo.tProduct ON dbo.tProduct.P_ProductType_ID = dbo.tPT.PT_ID Where (dbo.tProduct.P_ProductType_ID Is Null)"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.DeleteUnusedPTs", , , oPC.COShort
End Sub
Public Sub DeleteRelatedDocs(pTP As Long)
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "DELETE FROM dbo.tTR Where TR_TP_ID = pTP"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.DeleteRelatedDocs(pTP)", pTP
End Sub


'Public Function GetLabelsToPrint(pType As String, pStart As Boolean, Optional PID As Long, Optional pCODE As String) As XArrayDB
'    On Error GoTo errHandler
'Dim rs As New ADODB.Recordset
'Static XA As New XArrayDB
'
'
'    rs.CursorLocation = adUseClient
'    If pType = "D" Then
'        rs.Open "SELECT P_CODE,P_EAN,P_TITLE,P_SP,SUM(DELL_QTYFIRM) as QTYFIRM,SUM(DELL_QTYSS) as QTYSS FROM tDELL JOIN tProduct ON DELL_P_ID = P_ID WHERE DELL_TR_ID = " & PID & " GROUP BY P_CODE,P_TITLE,P_EAN,P_SP", oPC.CO, adOpenStatic, adLockReadOnly
'    ElseIf pType = "T" Then
'        rs.Open "SELECT P_CODE,P_EAN,P_TITLE,P_SP,SUM(TFRL_QTY) as QTYFIRM,0 as QTYSS FROM tTFRL JOIN tProduct ON TFRL_P_ID = P_ID WHERE TFRL_TR_ID = " & PID & " GROUP BY P_CODE,P_TITLE,P_EAN,P_SP", oPC.CO, adOpenStatic, adLockReadOnly
'    ElseIf pType = "P" Then
'        rs.Open "SELECT P_CODE,P_EAN,P_TITLE,P_SP,1 as QTYFIRM,0 as QTYSS FROM tProduct WHERE P_Code = '" & pCODE & "' GROUP BY P_CODE,P_TITLE,P_EAN,P_SP", oPC.CO, adOpenStatic, adLockReadOnly
'    End If
'  '  XA.ReDim 1, rs.RecordCount + Marker, 1, 7
'    If pStart Then XA.ReDim 1, 0, 1, 7
' '   tmpMarker = Marker
'    Do While Not rs.EOF
'        XA.InsertRows XA.UpperBound(1) + 1, 1
'        XA(XA.UpperBound(1), 1) = FormatCode(FNS(rs!P_Code), True)
'        XA(XA.UpperBound(1), 2) = FNS(rs!P_Title)
'        XA(XA.UpperBound(1), 3) = Format(FNN(rs!P_SP) / oPC.Configuration.DefaultCurrency.Divisor, "####0.00")
'        XA(XA.UpperBound(1), 4) = CStr(FNN(rs!QtySS) + FNN(rs!QtyFirm))
'        XA(XA.UpperBound(1), 5) = False
'        XA(XA.UpperBound(1), 6) = FNS(rs!P_EAN)
'        XA(XA.UpperBound(1), 7) = FNN(rs!P_SP)
'  '      XA(rs.AbsolutePosition + tmpMarker, 6) = FNN(rs!P_SP)
'        rs.MoveNext
'  '      Marker = Marker + 1
'    Loop
'    rs.Close
'    Set GetLabelsToPrint = XA
'    Exit Function
'errHandler:
'    RlsObjs rs
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_StockManager.GetLabelsToPrint(pType,pStart,PID,pCODE)", Array(pType, pStart, PID, _
'         pCODE)
'End Function
Public Function CountProductsPerPTID(pPTID As Long) As Long
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_CountProductsPerPT"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PTID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = pPTID
    Set par = cmd.CreateParameter("@COUNT", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    CountProductsPerPTID = cmd.Parameters(1)
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.CountProductsPerPTID(pPTID)", pPTID, , cmd
End Function


