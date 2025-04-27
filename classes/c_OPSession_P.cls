VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_OPSession_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Fetch(pSince As Date) As String
    On Error GoTo errHandler
Dim lngCount As Long
Dim udtProps As ZSessionProps
Dim udtData As ZSessionData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim objDisplay As d_ZSession
Dim lngID As Long
Dim objPB As PropertyBag
    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tZSession WHERE Z_NominalDate > '" & ReverseDate(pSince) & "' ORDER BY Z_NominalDate", oPC.CO, adOpenForwardOnly
    If rs.EOF Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    Do While Not rs.EOF
        With udtProps
            udtProps.ID = FNS(rs!Z_ID)
            udtProps.Cash = FNN(rs!Z_CashInDrawer)
            udtProps.CCVouchers = FNN(rs!Z_CCVouchersInDrawer)
            udtProps.Cheques = FNN(rs!Z_ChequesInDrawer)
            udtProps.EndDate = FND(rs!Z_EndDate)
            udtProps.GiftVouchers = FNN(rs!Z_VouchersInDrawer)
            udtProps.NominalDate = FND(rs!Z_NominalDate)
            udtProps.SMID = FNN(rs!Z_Supervisor)
            udtProps.StartDate = FND(rs!Z_StartDate)
            udtProps.TillPoint = FND(rs!Z_TillPoint)
            udtProps.TotalValueCredit = FNN(rs!Z_ReturnsValue)
            udtProps.TotalValueDiscount = FNN(rs!Z_DiscountsGiven)
            udtProps.TotalValueSales = FNN(rs!Z_SalesValue)
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_ZSession_P.Fetch(pSince)", pSince
End Function



