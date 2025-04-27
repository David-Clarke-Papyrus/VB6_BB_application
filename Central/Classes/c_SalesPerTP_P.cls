VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_SalesPerTP_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Fetch(TPID As Long, Optional pType As Integer, Optional pRecentOnly As Boolean) As String
Dim lngCount As Long
Dim udtProps As DocsTPProps
Dim udtData As DocsTPData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim objDisplay As d_SalesPerTP
Dim lngID As Long
Dim objPB As PropertyBag
    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
'    If pRecentOnly Then
'        rs.Open "SELECT * FROM vDocsPerTP WHERE TR_TP_ID = " & TPID & "" _
'            & " AND ((TR_CaptureDATE > '" & ReverseDate(oPC.Configuration.LastStockTakeDate) & "') " _
'            & " OR (TR_STATUS <> 4 AND TR_TYPE = 6) or (TR_STATUS <> 4 AND TR_TYPE = 9))", oPC.CO, adOpenForwardOnly
'    Else
'        If pType = 3 Or pType = 6 Or pType = 9 Then
'            rs.Open "SELECT * FROM vDocsPerTP WHERE TR_TP_ID = " & TPID & " AND TR_TYPE = " & pType, oPC.CO, adOpenForwardOnly
'        Else
'            rs.Open "SELECT * FROM vDocsPerTP WHERE TR_TP_ID = " & TPID, oPC.CO, adOpenForwardOnly
'        End If
'    End If
    If rs.EOF Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    On Err GoTo ERRH
    Do While Not rs.EOF
        With udtProps
            udtProps.TRCODE = FNS(rs!TR_Code)
            udtProps.TRDATE = FND(rs!TR_Date)
            udtProps.TRID = FNN(rs!TR_ID)
            udtProps.Type = FNN(rs!TR_Type)
            udtProps.TRSTATUS = FNN(rs!TR_Status)
            udtProps.OrderType = FNN(rs!CO_OrderType)
            Select Case udtProps.Type
            Case 3
                udtProps.TRValue = FNN(rs!IL_Val)
            Case 6
                udtProps.TRValue = FNN(rs!COL_Val)
            Case 9
                udtProps.TRValue = FNN(rs!APP_Val)
            End Select
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
    Resume

End Function


