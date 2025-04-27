VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_InvoicesContaining_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Fetch(pTPID As Long, pCODE As String) As String
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strSQL As String
Dim strWHERE As String
Dim udtProps As InvoicesContainingProps
Dim udtData As InvoicesContainingData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim pblnNoRecsReturned As Long
  
    oPC.OpenLocalDatabase
    
    strSQL = "SELECT DISTINCT TR_ID,TR_CODE,TR_CAPTUREDATE,TR_DATE,I_PAYABLEAMOUNT,IL_P_ID FROM tTR JOIN tInvoice on TR_ID = I_ID JOIN tILINE on IL_TR_ID = TR_ID WHERE IL_P_ID = dbo.FindPIDFromCODE('" & pCODE & "') AND TR_TP_ID = " & pTPID & " ORDER BY TR_CAPTUREDATE DESC"
    Set objPB = New PropertyBag
    rs.Open strSQL, oPC.DBMainConn, adOpenDynamic, adLockReadOnly
  
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
            .DocCode = FNS(rs!TR_CODE)
            .DOCDate = FND(rs!TR_DATE)
            .DocValue = FNN(rs!I_PAYABLEAMOUNT)
            .TRID = FNN(rs!TR_ID)
            .PID = FNS(rs!IL_P_ID)
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
    
EXIT_Handler:
    rs.Close
    Set rs = Nothing
    oPC.CloseLocalDatabase
    Exit Function
ERRH:
    MsgBox Error
    GoTo EXIT_Handler
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_InvoicesContaining_P.Fetch(pTPID)", pTPID
End Function





