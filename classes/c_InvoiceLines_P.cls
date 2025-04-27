VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_InvoiceLines_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function Fetch(pTPID As Long) As String
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strSQL As String
Dim strWHERE As String
Dim udtProps As COLPOSProps
Dim udtData As COLPOSData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim pblnNoRecsReturned As Long
  
    strSQL = "SELECT P_TITLE,tCOL.* FROM tCOL LEFT JOIN tPRODUCT ON COL_PID = P_ID WHERE COL_TPID = " & CStr(pTPID) & "  ORDER BY COL_DATE DESC"
    Set objPB = New PropertyBag
    rs.Open strSQL, oPC.DBLocalConn, adOpenDynamic, adLockReadOnly
  
    If rs.EOF Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    On Err GoTo ERRH
  
  
    Do While Not rs.EOF
        With udtProps
            .COLID = FNN(rs!COL_COLID)
            .TPID = FNN(rs!COL_TPID)
            .TRID = FNN(rs!COL_TRID)
            .COLDate = FND(rs!COL_Date)
            .Code = FNS(rs!COL_Code)
            
            .PID = FNS(rs!COL_PID)
            .Description = FNS(rs!P_Title)
            .Qty = FNN(rs!COL_Qty)
            .QtyDispatched = FNN(rs!COL_QtyDispatched)
            .Price = FNN(rs!COL_Price)
            .DiscountRate = FNDBL(rs!COL_DiscountRate)
            .Deposit = FNN(rs!COL_Deposit)
            .DepositStatus = FNS(rs!COL_DepositStatus)

            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
        End With
    Loop
    With objPB
        .WriteProperty "Count", lngCount
        Fetch = .Contents
    End With
    Set objPB = Nothing
    
EXIT_HANDLER:
    rs.Close
    Set rs = Nothing
    Exit Function
ERRH:
    MsgBox Error
    GoTo EXIT_HANDLER
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_COLS_P.Fetch(pTPID)", pTPID
End Function






