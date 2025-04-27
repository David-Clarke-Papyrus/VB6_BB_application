VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_DellsPerPIDTP_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event lngProgress(p As Long)
Event lngMax(p As Long)

Public Function Fetch(pTPID As Long, pPID As String) As String
Dim lngCount As Long
Dim udtProps As dDELLProps
Dim udtData As dDELLData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim dPOL As d_POLine
Dim lngID As Long
Dim objPB As PropertyBag
Dim strSQL As String
Dim strAction As String
Dim bTmp As Boolean
Dim lngRC As Long
Dim OpenResult As Integer

    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM vDELLSperPIDTP WHERE DELL_P_ID = '" & pPID & "' AND TR_TP_ID = " & pTPID
    rs.CursorLocation = adUseClient
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open strSQL, oPC.COShort, adOpenForwardOnly
    RaiseEvent lngMax(rs.RecordCount)
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    On Err GoTo H
    Set rs2 = New ADODB.Recordset
    Do While Not rs.eof
        With udtProps
            udtProps.DELLID = FNS(rs!DELL_ID)
            udtProps.TRID = FNS(rs!TR_ID)
            udtProps.Discount = FNDBL(rs!DELL_Discount)
            udtProps.ForeignPrice = FNN(rs!DELL_ForeignPrice)
            udtProps.Price = FNN(rs!DELL_Price)
            udtProps.Qty = FNN(rs!DELL_QtyTotal)
            udtProps.SINVRef = FNS(rs!DEL_SUpplierInvoiceRef)
            udtProps.SINVDate = FND(rs!DEL_SUpplierInvoiceDate)
            udtProps.DocRef = FNS(rs!TR_CODE)
            udtProps.DOCDate = FND(rs!TR_DATE)
            udtProps.CURRID = FNN(rs!DEL_CURR_ID)
                  udtProps.CurrDivisor = FNN(rs!CURR_Divisor)
                  udtProps.CurrFormatString = FNS(rs!CURR_FormatString)
            LSet udtData = udtProps
            lngCount = lngCount + 1
            RaiseEvent lngProgress(lngCount)
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
    Set rs2 = Nothing
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
    Resume

End Function



