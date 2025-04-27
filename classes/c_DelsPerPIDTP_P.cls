VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_DelsPerPIDTP_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM vDELLSperPIDTP WHERE DELL_P_ID = '" & pPID & "' AND TR_TP_ID = " & pTPID
    rs.CursorLocation = adUseClient
    rs.Open strSQL, oPC.CO, adOpenForwardOnly
    RaiseEvent lngMax(rs.RecordCount)
    If rs.EOF Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    On Err GoTo ERRH
    Set rs2 = New ADODB.Recordset
    Do While Not rs.EOF
        With udtProps
            udtProps.Actions = strAction
            udtProps.Ref = FNS(rs!POL_REF)
            udtProps.Price = FNN(rs!POL_Price)
            udtProps.ForeignPrice = FNN(rs!POL_ForeignPrice)
            udtProps.QtyFirm = FNN(rs!POL_QtyFirm)
            udtProps.QtySS = FNN(rs!POL_QtySS)
            udtProps.QtyReceivedSoFar = FNN(rs!POL_QTYReceivedSoFar)
            udtProps.Discount = FNDBL(rs!POL_Discount)
            udtProps.POLID = FNN(rs!POL_ID)
            udtProps.PID = FNS(rs!POL_P_ID)
            udtProps.DocCode = FNS(rs!TRCode)
            udtProps.DocDate = FND(rs!TRDate)
            udtProps.CFormat = FNS(rs!CFormat)
            udtProps.CDivisor = FNN(rs!CDivisor)
            udtProps.CForeignDivisor = FNN(rs!CDivisor)
            udtProps.CFactor = FNDBL(rs!CFactor)
            udtProps.COLID = FNN(rs!COLID)
            udtProps.QtyFirm = FNN(rs!POL_QtyFirm)
            udtProps.QtySS = FNN(rs!POL_QtySS)
            udtProps.QtyOutstanding = udtProps.QtySS + udtProps.QtyFirm - FNN(rs!POL_QTYReceivedSoFar)
            udtProps.Title = FNS(rs!P_Title)
            udtProps.MainAuthor = FNS(rs!P_MainAuthor)
            udtProps.Code = FNS(rs!P_Code)
            udtProps.EAN = FNS(rs!P_EAN)
            udtProps.Supplier = FNS(rs!TP_Name)
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
    Exit Function
ERRH:
    MsgBox Error
    GoTo EXIT_Handler
    Resume

End Function



