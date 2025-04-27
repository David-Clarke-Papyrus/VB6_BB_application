VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_ZSession_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(pSince As Date, pExactdate As Date, pExchangeNumber As Long) As String
    On Error GoTo errHandler
Dim lngCount As Long
Dim udtProps As ZSessionProps
Dim udtData As ZSessionData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim objDisplay As d_ZSession
Dim lngID As Long
Dim objPB As PropertyBag
Dim OpenResult As Integer
Dim strZSessionID  As String
Dim arZS() As String
    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If pExchangeNumber > 0 Then
        rs.open "SELECT EXCH_ZSessionID From dbo.tExchange Where (EXCH_Number = " & CStr(pExchangeNumber) & ")", oPC.COShort, adOpenForwardOnly
        If Not (rs.eof And rs.BOF) Then
            strZSessionID = ""
            Do While Not rs.eof
                strZSessionID = strZSessionID & IIf(strZSessionID > "", ",", "") & "'" & rs.fields(0) & "'"
                rs.MoveNext
            Loop
        End If
        rs.Close
    End If
    If oPC.BlindCashup = True Then
        If strZSessionID > "" Then
            rs.open "SELECT *,Status FROM tZSession  LEFT JOIN dbo.vReportableCashups ON dbo.tZSession.Z_ID = dbo.vReportableCashups.ZID LEFT JOIN tStaffMember ON Z_SUPERVISOR = SM_ID LEFT JOIN tPOSClient c ON Z_TILLPOINT = StationName LEFT JOIN tSTORE d ON c.STORE_ID = d.STORE_ID WHERE Z_ID IN (" & strZSessionID & ") ", oPC.COShort, adOpenForwardOnly
        ElseIf pExactdate > CDate("1/1/2000") Then
            rs.open "SELECT *,Status FROM tZSession  LEFT JOIN dbo.vReportableCashups ON dbo.tZSession.Z_ID = dbo.vReportableCashups.ZID LEFT JOIN tStaffMember ON Z_SUPERVISOR = SM_ID LEFT JOIN tPOSClient c ON Z_TILLPOINT = StationName LEFT JOIN tSTORE d ON c.STORE_ID = d.STORE_ID WHERE Z_NominalDate = '" & ReverseDate(pExactdate) & "' ORDER BY Z_NominalDate DESC", oPC.COShort, adOpenForwardOnly
        Else
            rs.open "SELECT *,Status FROM tZSession  LEFT JOIN dbo.vReportableCashups ON dbo.tZSession.Z_ID = dbo.vReportableCashups.ZID LEFT JOIN tStaffMember ON Z_SUPERVISOR = SM_ID LEFT JOIN tPOSClient c ON Z_TILLPOINT = StationName LEFT JOIN tSTORE d ON c.STORE_ID = d.STORE_ID WHERE Z_NominalDate > '" & ReverseDate(pSince) & "' ORDER BY Z_NominalDate DESC", oPC.COShort, adOpenForwardOnly
        End If
    Else
        If strZSessionID > "" Then
            rs.open "SELECT * FROM tZSession LEFT JOIN tStaffMember ON Z_SUPERVISOR = SM_ID LEFT JOIN tPOSClient c ON Z_TILLPOINT = StationName LEFT JOIN tSTORE d ON c.STORE_ID = d.STORE_ID WHERE Z_ID IN (" & strZSessionID & ") ", oPC.COShort, adOpenForwardOnly
        ElseIf pExactdate > CDate("1/1/2000") Then
            rs.open "SELECT * FROM tZSession LEFT JOIN tStaffMember ON Z_SUPERVISOR = SM_ID LEFT JOIN tPOSClient c ON Z_TILLPOINT = StationName LEFT JOIN tSTORE d ON c.STORE_ID = d.STORE_ID  WHERE Z_NominalDate = '" & ReverseDate(pExactdate) & "' ORDER BY Z_NominalDate DESC", oPC.COShort, adOpenForwardOnly
        Else
            rs.open "SELECT * FROM tZSession LEFT JOIN tStaffMember ON Z_SUPERVISOR = SM_ID LEFT JOIN tPOSClient c ON Z_TILLPOINT = StationName LEFT JOIN tSTORE d ON c.STORE_ID = d.STORE_ID  WHERE Z_NominalDate > '" & ReverseDate(pSince) & "' ORDER BY Z_NominalDate DESC", oPC.COShort, adOpenForwardOnly
        End If
    End If
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    Do While Not rs.eof
        With udtProps
            udtProps.ID = FNS(rs!Z_ID)
            udtProps.EndDate = FND(rs!Z_EndDate)
            udtProps.NominalDate = FND(rs!Z_NominalDate)
            udtProps.SMID = FNN(rs!Z_Supervisor)
            udtProps.SupervisorName = FNS(rs!SM_Name)
            udtProps.StartDate = FND(rs!Z_StartDate)
            udtProps.TillPoint = FNS(rs!Z_TillPoint) & IIf(rs!STORE_ID <> oPC.Configuration.DefaultStoreID, "/" & FNS(rs!STORE_CODE), "")
            udtProps.TotalValueCredit = FNN(rs!Z_ReturnsValue)
            udtProps.TotalValueDiscount = FNN(rs!Z_DiscountsGiven)
            udtProps.TotalValueSales = FNN(rs!Z_SalesValue)
            If oPC.BlindCashup = True Then
                udtProps.Reportable = FNINT(rs!Status)
            Else
                udtProps.Reportable = 1
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
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_ZSession_P.Fetch(pSince)", pSince
End Function


