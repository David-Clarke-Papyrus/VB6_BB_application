VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_APPLines_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Fetch(pTRID As Long) As String
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strSQL As String
Dim strWHERE As String
Dim udtProps As APPLPOSProps
Dim udtData As APPLPOSData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim pblnNoRecsReturned As Long
  
    oPC.OpenLocalDatabase
    
    strSQL = "SELECT dbo.CODEF(P_CODE,P_EAN,0) as CODEF,tAPPL.APPL_APPLID,tAPPL.APPL_PID,tAPPL.APPL_DATE,tAPPL.APPL_CODE,tAPPL.APPL_QtyOut,tAPPL.APPL_QtyBack,tAPPL.APPL_Price,tAPPL.APPL_DiscountRate,tAPPL.APPL_TPID,tAPPL.APPL_TRID,P_Title,P_MainAuthor,P_EAN,P_CODE,P_VATRATE " _
            & " FROM tAPPL JOIN tProduct ON APPL_PID = P_ID WHERE APPL_TRID = " & CStr(pTRID) & " AND APPL_QTYOUT > APPL_QTYBACK ORDER BY P_Title ASC"
    Set objPB = New PropertyBag
    rs.Open strSQL, oPC.DBLocalConn, adOpenStatic, adLockReadOnly
  
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
            .APPLID = FNS(rs!APPL_APPLID)
            .title = FNS(rs!P_TITLE)
            .Author = FNS(rs!P_MainAuthor)
            .QtyOut = FNN(rs!APPL_Qtyout)
            .QtyBack = FNN(rs!APPL_QtyBack)
            .QtyNett = FNN(rs!APPL_Qtyout) - FNN(rs!APPL_QtyBack)
            .Price = FNN(rs!APPL_PRICE)
            .Code = FNS(rs!P_EAN)
            .CodeF = FNS(rs!CodeF)
            .DiscountRate = FNDBL(rs!APPL_DISCOUNTRATE)
            .VATRate = FNDBL(rs!P_VATRATE)
            .PID = FNS(rs!APPL_PID)
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
    ErrorIn "c_APPLines_P.Fetch(pTRID)", pTRID
End Function





