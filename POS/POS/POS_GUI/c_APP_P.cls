VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_APP_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Fetch(pTPID As Long) As String
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim strWHERE As String
Dim udtProps As APPPOSProps
Dim udtData As APPPOSData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim pblnNoRecsReturned As Long
  
'-------------------------------
    oPC.OpenLocalDatabase
'-------------------------------
  
    strSQL = "SELECT tAPPL.APPL_DATE,tAPPL.APPL_CODE,tAPPL.APPL_TPID,tAPPL.APPL_TRID FROM tAPPL " _
            & " WHERE APPL_TPID = " & CStr(pTPID) & " AND APPL_QTYOUT > APPL_QTYBACK" _
            & " GROUP BY tAPPL.APPL_DATE,tAPPL.APPL_CODE,tAPPL.APPL_TPID,tAPPL.APPL_TRID ORDER BY APPL_DATE DESC"
    Set objPB = New PropertyBag
    rs.CursorLocation = adUseClient
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
            .TPID = FNN(rs!APPL_TPID)
            .APPID = FNN(rs!APPL_TRID)
            .DOCDate = FND(rs!APPL_Date)
            .DocCode = FNS(rs!APPL_CODE)
            
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
'---------------------------------------------------
    oPC.CloseLocalDatabase
'---------------------------------------------------
    Exit Function
ERRH:
    MsgBox Error
    GoTo EXIT_Handler
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_APP_P.Fetch(pTPID)", pTPID
End Function




