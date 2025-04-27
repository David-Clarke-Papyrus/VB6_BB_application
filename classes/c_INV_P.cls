VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_INV_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function Fetch() As String
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strSQL As String
Dim strWHERE As String
Dim udtProps As INVPOSProps
Dim udtData As InvPOSData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim pblnNoRecsReturned As Long
  
    strSQL = "SELECT * FROM tBridge ORDER BY INV_DOCDATE DESC"
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
            .INVID = FNN(rs!INV_ID)
            .DOCDate = FND(rs!INV_DOCDATE)
            .DocCode = FNS(rs!INV_DOCCODE)
            .CustomerName = FNS(rs!INV_CUSTOMERNAME)
            .Payable = FNN(rs!INV_PAYABLE)
            .TPID = FNN(rs!INV_TPID)
            .VAT = FNN(rs!INV_VATVALUE)
            

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
    Exit Function
ERRH:
    MsgBox Error
    GoTo EXIT_Handler
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_COLS_P.Fetch()"
End Function


