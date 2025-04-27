VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ro_PO_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function Fetch(pTRID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim strSQL As String
Dim udtProps As POProps
Dim udtData As POData
Dim objPersist As ro_POL_P
Dim objPB As PropertyBag

    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tPO JOIN tTR ON PO_ID = TR_ID JOIN tTP ON TR_TP_ID = TP_ID " _
        & " LEFT JOIN tStore ON TR_ADDRESSDOCID = Store_ID LEFT JOIN tStore s2 ON TR_ADDRESSGOODSID = s2.STORE_ID LEFT JOIN tADD ON TP_DEFAULTADDRESSID = ADD_ID WHERE TR_ID = " & pTRID
    rs.CursorLocation = adUseClient
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open "SELECT * FROM PO_FETCH_ro WHERE TR_ID =" & CStr(pTRID), oPC.COShort, adOpenForwardOnly
    If rs.eof Then
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    
    With udtProps
        .TRID = rs!TR_ID
        .TPID = rs!TR_TP_ID
        .TotalVAT = FNN(rs!PO_VATAmount)
        .TotalPayable = FNN(rs!PO_PayableAmount)   'Always store the local amount
        .DelToAddress = FNS(rs!store_delivertoAddress)
        .DELTOStoreID = FNN(rs!TR_AddressDOCID)
        
        .COMPID = FNN(rs!TR_COMP_ID)
        .StaffID = FNN(rs!TR_StaffID)
        .TPPhone = FNS(rs!TP_Phone)
        .TPPhone2 = FNS(rs!ADD_BusPhone)
        .TPACCNum = FNS(rs!TP_ACno)
        .OrderType = FNS(rs!PO_OrderType)
        .TPFax = FNS(rs!ADD_Fax)
        .TPNAME = FNS(rs!TP_Name)
        .Memo = FNS(rs!TR_NOTE)
        .CurrencyID = FNN(rs!PO_CURR_ID)
        .CurrencyRate = FNDBL(rs!PO_CURRRATE)
        .DOCCode = FNS(rs!TR_CODE)
        .Log = Trim(IIf(IsNull(rs!TR_Log), "", rs!TR_Log))
        If rs!TR_DATE > DateSerial(1995, 1, 1) Then .DOCDate = DateSerial(Year(rs!TR_DATE), Month(rs!TR_DATE), Day(rs!TR_DATE))
        .CaptureDate = FND(rs!TR_CaptureDate)
        .Status = FNN(rs!TR_Status)
        .StaffID = FNN(rs!TR_StaffID)
            If .StaffID > 0 Then
                .StaffName = oPC.Configuration.Staff.FindStaffByID(.StaffID).Shortname
                .Signature = oPC.Configuration.Staff.FindStaffByID(.StaffID).Signature
                .StaffEmail = oPC.Configuration.Staff.FindStaffByID(.StaffID).EMail
            End If

        rs.Close
        Set rs = Nothing
        LSet udtData = udtProps
          Set objPersist = New ro_POL_P
        Set objPB = New PropertyBag
        With objPB
          .WriteProperty "State", udtData.buffer
          .WriteProperty "POLS", objPersist.Fetch(udtProps.TRID)
          Fetch = .Contents
        End With
        Set objPB = Nothing
        Set objPersist = Nothing
    End With
    
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
  
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "ro_PO_P.Fetch(pTRID)", pTRID, , rs
End Function


