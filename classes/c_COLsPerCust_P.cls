VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_COLsPerCust_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(TPID As Long) As String
Dim lngCount As Long
Dim udtProps As dCOLProps
Dim udtData As dCOLData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim dPOL As d_POLine
Dim lngID As Long
Dim objPB As PropertyBag
Dim strSQL As String
Dim OpenResult As Integer

    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM COLsOS WHERE TPID = " & TPID
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open strSQL, oPC.COShort, adOpenForwardOnly
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    On Err GoTo H
    Do While Not rs.eof
        With udtProps
            udtProps.Ref = FNS(rs!COL_Ref)
            udtProps.Price = FNN(rs!COL_PRICE)
            udtProps.Qty = FNN(rs!COL_Qty)
            udtProps.QtyDispatched = FNN(rs!COL_QtyDispatched)
            udtProps.Discount = FNDBL(rs!COL_DiscountPercent)
            udtProps.COLID = FNN(rs!COL_ID)
            udtProps.PID = FNS(rs!COL_P_ID)
            udtProps.DOCCode = FNS(rs!TRCODE)
            udtProps.DOCDate = FND(rs!TRDATE)
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
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
H:
    MsgBox Error
    GoTo EXIT_Handler
    Resume

End Function



