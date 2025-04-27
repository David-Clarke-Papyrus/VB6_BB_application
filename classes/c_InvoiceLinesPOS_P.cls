VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_InvoiceLinesPOS_P"
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
Dim udtProps As InvoiceLinesProps
Dim udtData As InvoiceLinesData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim pblnNoRecsReturned As Long
  
    oPC.OpenLocalDatabase
  
    strSQL = "SELECT P_TITLE,P_CODE,P_EAN,IL_DISCOUNTRATE,IL_PRICE,IL_QTY,IL_ID,IL_SALESVALUE,IL_P_ID,QTYCREDITED " _
                & " FROM tILine JOIN tPRODUCT ON IL_P_ID = P_ID  LEFT JOIN vQtyCreditedPerIL ON ILID = IL_ID WHERE IL_TR_ID = " & pTRID & "  ORDER BY P_TITLE DESC"
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
            .Description = FNS(rs!P_TITLE)
            .DiscountRate = FNDBL(rs!IL_DISCOUNTRate)
            .Price = FNN(rs!IL_PRICE)
            .Qty = FNN(rs!IL_QTY)
            .QtyCredited = FNN(rs!QtyCredited)
            .PID = FNS(rs!IL_P_ID)
            .ILID = FNN(rs!IL_ID)
            If FNS(rs!P_CODE) > "" Then
                .LineCode = FNS(rs!P_CODE)
            Else
                .LineCode = FNS(rs!P_EAN)
            End If

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
    ErrorIn "c_InvoiceLinesPOS_P.Fetch(pTPID)", pTRID
End Function






