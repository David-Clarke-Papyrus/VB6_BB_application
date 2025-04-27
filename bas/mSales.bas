Attribute VB_Name = "mSales"
Option Explicit
Dim startingQTYOH As Integer
Dim iLastStatus As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

Public Function WeeklySalesPerPID(pPID As String, XSALES_CY As XArrayDB, XSALES_LY As XArrayDB, XOOS As XArrayDB, XOOS_LY As XArrayDB, pQTYOH As Long, Optional xHeadings As XArrayDB)
    On Error GoTo errHandler
Dim lngCount As Long
Dim pblnNoRecsReturned As Boolean
Dim lngID As Long
Dim objPB As PropertyBag
Dim strSQL As String
Dim i As Integer
Dim rs As ADODB.Recordset
Dim rsOOS As ADODB.Recordset
Dim rsOOSLY As ADODB.Recordset
Dim OpenResult As Integer

    startingQTYOH = pQTYOH
    Set cmd = New ADODB.Command
    cmd.CommandText = "PRODUCTSALESANNUAL"
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandType = adCmdStoredProc
    Set par = cmd.CreateParameter("@PID", adGUID, adParamInput)
    cmd.Parameters.Append par
    par.Value = pPID
    Set rs = cmd.execute
    
    XSALES_CY.ReDim 1, 1, 1, 53
    For i = 1 To 53
        XSALES_CY(1, i) = FNINT(rs.fields(i - 1))
    Next
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set par = Nothing
    
    Set cmd = New ADODB.Command
    cmd.CommandText = "PRODUCTSALESANNUAL_LY"
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandTimeout = 360
    cmd.CommandType = adCmdStoredProc
    Set par = cmd.CreateParameter("@PID", adGUID, adParamInput)
    cmd.Parameters.Append par
    par.Value = pPID
    Set rs = cmd.execute
    
    XSALES_LY.ReDim 1, 1, 1, 53
    For i = 1 To 53
        XSALES_LY(1, i) = FNINT(rs.fields(i - 1))
    Next
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set par = Nothing
    
    
    
    
    If Not IsMissing(xHeadings) Then
        Set cmd = New ADODB.Command
        cmd.CommandText = "MonthNamesperWeekNum"
        cmd.ActiveConnection = oPC.COShort
        cmd.CommandType = adCmdStoredProc
        Set par = cmd.CreateParameter("@ThisYear", adInteger, adParamInput)
        cmd.Parameters.Append par
        par.Value = Year(Date)
        Set rs = cmd.execute
        xHeadings.ReDim 1, 1, 1, 53
        For i = 1 To 53
            xHeadings(1, i) = FNS(rs.fields(i - 1))
          '  i = i + 1
        Next
        rs.Close
        Set rs = Nothing
        Set cmd = Nothing
        Set par = Nothing
    End If
    
    MarkOOS pPID, "CY", XSALES_CY
    MarkOOS pPID, "LY", XSALES_LY
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mSales.WeeklySalesPerPID(pPID,XSALES_CY,XSALES_LY,XOOS,XOOS_LY,pQTYOH)", Array(pPID, _
         XSALES_CY, XSALES_LY, XOOS, XOOS_LY, pQTYOH)
End Function
Private Sub MarkOOS(PID As String, Series As String, x As XArrayDB, Optional iRow As Integer)
Dim rs As ADODB.Recordset
Dim iYear As Integer
Dim i As Integer
Dim strCurrentSymbol As String
Dim iNextChange As Integer
    If Series = "CY" Then
        iYear = Year(Date)
    ElseIf Series = "LY" Then
        iYear = Year(Date) - 1
    End If

    If iRow = 0 Then iRow = 1  'in case iRow is not supplied
    
    Set rs = New ADODB.Recordset
    rs.open "Select *,dbo.udf_DT_ISOWeekNum_forYear(OOS_DATE," & iYear & ") dte FROM tOOS WHERE  dbo.udf_DT_ISOWeekNum_forYear(OOS_DATE," & iYear & ") <>0 and OOS_P_ID = '" & PID & "' ORDER BY OOS_DATE DESC,OOS_ID", oPC.COShort, adOpenForwardOnly, adLockReadOnly
    
    i = 53
    If Series = "CY" Then
        iLastStatus = IIf(startingQTYOH > 0, 1, 0)
    End If
    If Not rs.eof Then
        Do While i >= iNextChange And i >= 1  'Do each week until next stock outage or stock replenishment
            If iNextChange <> i Then   'If stock is in at any point in week it must show as IN
                If iLastStatus = 0 Then
                    If x(iRow, i) = 0 Then
                        x(iRow, i) = "*"
                    End If
                Else
                End If
            End If
            If Not rs.eof Then
                If (rs!dte) = i Then   'Special case: we are at the week where stock status changes
                    iLastStatus = IIf(rs!OOS_Type = "I", 0, 1)
                    rs.MoveNext
                    If Not rs.eof Then
                        iNextChange = rs!dte
                    Else
                        iNextChange = 0
                    End If
                End If
            End If
            If i > 0 Then i = i - 1
        Loop
    End If
    rs.Close
    Set rs = Nothing
    
     For i = i To 1 Step -1
        If iLastStatus = 0 Then
            If x(iRow, i) = 0 Then
                x(iRow, i) = "*"
            End If
        End If
    Next i
End Sub


Public Sub WeeklySalesSet(AuthorName As String, XSALES_CY As XArrayDB, XSALES_LY As XArrayDB, XOOS As XArrayDB, XOOS_LY As XArrayDB, pQTYOH As Long)
    On Error GoTo errHandler
Dim lngCount As Long
Dim pblnNoRecsReturned As Boolean
Dim lngID As Long
Dim objPB As PropertyBag
Dim strSQL As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim rs As ADODB.Recordset
Dim rsOOS As ADODB.Recordset
Dim rsOOSLY As ADODB.Recordset
Dim OpenResult As Integer

    XSALES_CY.ReDim 0, 0, 1, 56
    startingQTYOH = pQTYOH
    Set cmd = New ADODB.Command
    cmd.CommandText = "PRODUCTSALESSetANNUAL"
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandType = adCmdStoredProc
    Set par = cmd.CreateParameter("@AUthorname", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = AuthorName
    Set rs = cmd.execute
    j = 1
    Do While Not rs.eof
        XSALES_CY.ReDim 1, j, 1, 56
        For i = 1 To 53
            XSALES_CY(j, i) = FNINT(rs.fields(i - 1))
        Next
        XSALES_CY(j, 54) = FNS(rs.fields(53))
        XSALES_CY(j, 55) = FNS(rs.fields(54))
        XSALES_CY(j, 56) = FNS(rs.fields(55))
        rs.MoveNext
        j = j + 1
    Loop
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set par = Nothing
    
    
    Set cmd = New ADODB.Command
    cmd.CommandText = "PRODUCTSALESSetANNUAL_LY"
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    Set par = cmd.CreateParameter("@AUthorname", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = AuthorName
    Set rs = cmd.execute
    j = 1
    Do While Not rs.eof
        XSALES_LY.ReDim 1, j, 1, 53
        For i = 1 To 53
            XSALES_LY(j, i) = FNINT(rs.fields(i - 1))
        Next
    '    XSALES_LY(j, 54) = FNINT(rs.Fields(54))
        rs.MoveNext
        j = j + 1
    Loop
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set par = Nothing
    
    k = j - 1
    For j = 1 To k
 '   MsgBox XSALES_CY(j, 54)
        startingQTYOH = XSALES_CY(j, 56)
        MarkOOS XSALES_CY(j, 55), "CY", XSALES_CY, j
        MarkOOS XSALES_CY(j, 55), "LY", XSALES_LY, j
    Next j

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "mSales.WeeklySalesSet(XSALES_CY,XSALES_LY,XOOS,XOOS_LY,pQTYOH)", Array(XSALES_CY, XSALES_LY, XOOS, XOOS_LY, pQTYOH)
End Sub


