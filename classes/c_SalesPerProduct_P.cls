VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_SalesPerProduct_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(pPID As String, XSALES As XArrayDB, XSALES_LY As XArrayDB, XOOS As XArrayDB, XOOS_LY As XArrayDB, pQTYOH As Long)
Dim lngCount As Long
Dim udtProps As SPCProps
Dim udtData As SPCData
Dim pblnNoRecsReturned As Boolean
Dim lngID As Long
Dim objPB As PropertyBag
Dim strSQL As String
Dim OpenResult As Integer
Dim i As Integer
Dim rs As ADODB.Recordset
Dim rsOOS As ADODB.Recordset
Dim rsOOSLY As ADODB.Recordset
Dim iWeek As Integer
Dim iLastStatus As Integer
Dim cmd As New ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.CommandText = "PRODUCTSALESANNUAL"
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandType = adCmdStoredProc
    Set par = cmd.CreateParameter("@PID", adGUID, adParamInput)
    cmd.Parameters.Append par
    par.Value = pPID
    Set rs = cmd.execute
    
    XSALES.ReDim 1, 1, 1, 53
    For iWeek = 1 To 53
        XSALES(1, iWeek) = FNINT(rs.fields(iWeek - 1))
    Next
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set par = Nothing

    cmd.CommandText = "PRODUCTSALESANNUAL_LY"
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandType = adCmdStoredProc
    Set par = cmd.CreateParameter("@PID", adGUID, adParamInput)
    cmd.Parameters.Append par
    par.Value = pPID
    Set rs = cmd.execute
    
    XSALES_LY.ReDim 1, 1, 1, 53
    For iWeek = 1 To 53
        XSALES_LY(1, iWeek) = FNINT(rs.fields(iWeek - 1))
    Next
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set par = Nothing
    
'-------------------------------------------------------Current year
    Set rsOOS = New ADODB.Recordset
    rsOOS.open "Select *,dbo.udf_DT_ISOWeekNum_forYear(OOS_DATE,YEAR(GETDATE())) dte FROM tOOS WHERE  dbo.udf_DT_ISOWeekNum_forYear(OOS_DATE,YEAR(GETDATE())) <>0 and OOS_P_ID = '" & pPID & "' ORDER BY OOS_DATE DESC,OOS_ID", oPC.COShort, adOpenForwardOnly, adLockReadOnly
    XOOS.ReDim 1, 1, 1, 53
    For i = 1 To 53
        XOOS(1, i) = 0
    Next
    
    
    i = 53
    iLastStatus = IIf(pQTYOH > 0, 1, 0)
    If Not rsOOS.eof Then  'Make sure the first record we get corresponds to the in-stock state we have set
        If IIf(rsOOS!OOS_Type = "I", 0, 1) <> iLastStatus Then
            rsOOS.MoveNext
        End If
    End If
    If Not rsOOS.eof Then
        Do While i >= (rsOOS!dte) And i >= 1
           ' If i = 25 Then MsgBox "Here"
            If rsOOS!dte = i Then   'If stock is in at any point in week it must show as IN
                XOOS(1, i) = 1
            Else
                XOOS(1, i) = iLastStatus
            End If
            If (rsOOS!dte) = i Then
                iLastStatus = IIf(rsOOS!OOS_Type = "I", 0, 1)
                rsOOS.MoveNext
                If rsOOS.eof Then
                    Exit Do
                End If
            End If
            If rsOOS!dte <> i Then i = i - 1
        Loop
    End If
    rsOOS.Close
    Set rsOOS = Nothing
    
    For i = i To 1 Step -1
        XOOS(1, i) = iLastStatus
    Next i
    
'------------------------------------------------------- Last year
    
    Set rsOOSLY = New ADODB.Recordset
    rsOOSLY.open "Select *,dbo.udf_DT_ISOWeekNum_forYear(OOS_DATE,YEAR(GETDATE())-1) dte FROM tOOS WHERE  dbo.udf_DT_ISOWeekNum_forYear(OOS_DATE,YEAR(GETDATE())-1) <>0 and OOS_P_ID = '" & pPID & "' ORDER BY OOS_DATE DESC,OOS_ID", oPC.COShort, adOpenForwardOnly, adLockReadOnly
    XOOS_LY.ReDim 1, 1, 1, 53
    For i = 1 To 53
        XOOS_LY(1, i) = 0
    Next
    
    i = 53
    If Not rsOOSLY.eof Then
        Do While i >= (rsOOSLY!dte) And i >= 1
       '     If i = 25 Then MsgBox "Here"
            If rsOOSLY!dte = i Then   'If stock is in at any point in week it must show as IN
                XOOS_LY(1, i) = 1
            Else
                XOOS_LY(1, i) = iLastStatus
            End If
            If (rsOOSLY!dte) = i Then
                iLastStatus = IIf(rsOOSLY!OOS_Type = "I", 0, 1)
                rsOOSLY.MoveNext
                If rsOOSLY.eof Then
                    Exit Do
                End If
            End If
            If rsOOSLY!dte <> i Then i = i - 1
        Loop
    End If
    rsOOSLY.Close
    Set rsOOSLY = Nothing
    For i = i To 1 Step -1
        XOOS_LY(1, i) = iLastStatus
    Next i
    
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
    Resume

End Function

