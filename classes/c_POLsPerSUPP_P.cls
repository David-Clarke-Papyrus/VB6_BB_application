VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_POLsPerSUPP_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Fetch(TPID As Long) As String
Dim lngCount As Long
Dim udtProps As dPOLProps
Dim udtData As dPOLData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim dPOL As d_POLine
Dim lngID As Long
Dim objPB As PropertyBag
Dim strSQL As String

    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM POLsOS WHERE TPID = " & TPID
    rs.Open strSQL, oPC.CO, adOpenForwardOnly
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
            udtProps.Ref = FNS(rs!POL_REF)
            udtProps.Price = FNN(rs!POL_Price)
            udtProps.QtyFirm = FNN(rs!POL_QtyFirm)
            udtProps.QtySS = FNN(rs!POL_QTySS)
            udtProps.RecSoFar = FNN(rs!POL_QtyReceivedSoFar)
            udtProps.Discount = FND(rs!POL_Discount)
            udtProps.POLID = FNN(rs!POL_ID)
            udtProps.PID = FNS(rs!POL_P_ID)
            udtProps.DocCode = FNS(rs!TRCode)
            udtProps.DocDate = FND(rs!TRDate)
            udtProps.CFormat = FNS(rs!CFormat)
            udtProps.CDivisor = FNN(rs!CDivisor)
            udtProps.CFactor = FND(rs!CFactor)
            udtProps.COLID = FNN(rs!COLID)
            udtProps.QtyFirm = FNN(rs!POL_QtyFirm)
            udtProps.QtySS = FNN(rs!POL_QTySS)
            udtProps.QtyOutstanding = FNN(rs!POL_QTYOutstanding)
            udtProps.Title = FNS(rs!P_Title)
            udtProps.Author = FNS(rs!P_MainAuthor)
            udtProps.Code = FNS(rs!P_Code)
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


