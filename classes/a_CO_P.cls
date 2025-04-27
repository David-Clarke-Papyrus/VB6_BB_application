VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_CO_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function Fetch(TRID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As COrderProps
Dim udtData As COrderData
Dim objPersist As a_COL_P
Dim objPB As PropertyBag
Dim oUtil As z_SQL
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set oUtil = New z_SQL
    lngResult = oUtil.RunGetRecordset("SELECT * FROM COrder_Fetch WHERE TR_ID = " & TRID, enText, Array(), "", rs)
    If Not rs.eof Then
        With udtProps
            .TRID = rs!TR_ID
            .TPID = FNN(rs!TR_TP_ID)
            .GoodsAddressID = FNN(rs.fields("AddressGOODSID"))
            .BillToAddressID = FNN(rs.fields("AddressDOCID"))
            .OrderNum = FNS(rs!CO_OrderNum)
            .ForAttn = FNS(rs!CO_ForAttn)
            .TPACCNum = FNS(rs!TP_ACno)
            .TPFax = FNS(rs!ADD_Fax)
            .TPPhone = FNS(rs!TP_Phone)
            .TPNAME = FNS(rs!TP_Name)
            .Memo = FNS(rs!TR_NOTE)
            .DOCCode = FNS(rs!TR_CODE)
            .DOCDate = FND(rs!TR_DATE)
            .CaptureDate = FND(rs!TR_CaptureDate)
            .Status = FNStatus(rs!TR_Status)
            .OrderType = FNN(rs!CO_OrderType)
            .StaffID = FNN(rs!TR_StaffID)
            .ORGUID = FNS(rs!CO_ORGUID)
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            Set oUtil = Nothing
            LSet udtData = udtProps
            Set objPersist = New a_COL_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.buffer
              .WriteProperty "COLS", objPersist.Fetch(udtProps.TRID)
              Fetch = .Contents
            End With
            Set objPB = Nothing
            Set objPersist = Nothing
        End With
    Else
        ' force an or
        rs.MoveNext
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Function
errHandler:
    ErrPreserve
    If OpenResult = 0 Then oPC.DisconnectDBShort
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO_P.Fetch(TRID)", TRID, , rs
End Function
Public Function FetchWants(TPID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As COrderProps
Dim udtData As COrderData
Dim objPersist As a_COL_P
Dim objPB As PropertyBag
Dim oUtil As z_SQL
Dim OpenResult As Integer
    Set rs = New ADODB.Recordset
    Set oUtil = New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    lngResult = oUtil.RunGetRecordset("SELECT * FROM COrder_Fetch WHERE TR_TP_ID = " & TPID & " AND CO_ORDERTYPE = 2", enText, Array(), "", rs)
    
    If Not rs.eof Then
        With udtProps
            .TRID = rs!TR_ID
            .TPID = FNN(rs!TR_TP_ID)
            .GoodsAddressID = FNN(rs.fields("AddressGOODSID"))
            .BillToAddressID = FNN(rs.fields("AddressDOCID"))
            .OrderNum = FNS(rs!CO_OrderNum)
            .ForAttn = FNS(rs!CO_ForAttn)
            .TPACCNum = FNS(rs!TP_ACno)
            .TPFax = FNS(rs!ADD_Fax)
            .TPPhone = FNS(rs!TP_Phone)
            .TPNAME = FNS(rs!TP_Name)
            .Memo = FNS(rs!TR_NOTE)
            .DOCCode = FNS(rs!TR_CODE)
            .DOCDate = FND(rs!TR_DATE)
            .CaptureDate = FND(rs!TR_CaptureDate)
            .Status = rs!TR_Status
            .OrderType = FNN(rs!CO_OrderType)
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            Set oUtil = Nothing
            LSet udtData = udtProps
            Set objPersist = New a_COL_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.buffer
              .WriteProperty "COLS", objPersist.Fetch(udtProps.TRID)
              FetchWants = .Contents
            End With
            Set objPB = Nothing
            Set objPersist = Nothing
        End With
    Else
        ' force an or
        FetchWants = ""
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO_P.FetchWants(TPID)", TPID
End Function


Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strBuffer As String
Dim strSQL As String
Dim objCode As New z_Code
Dim udtProps As COrderProps
Dim udtData As COrderData
Dim objPersist As a_COL_P
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim OpenResult As Integer

    Set objPB = New PropertyBag
    arBuffer = buffer
    With objPB
      .Contents = arBuffer
      udtData.buffer = .ReadProperty("State")
    End With
    LSet udtProps = udtData
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tTR WHERE TR_ID =" & udtProps.TRID
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    With udtProps
        If .IsNew Then rs.AddNew
        rs!TR_Type = IIf(IsNull(.OrderType), 1, .OrderType)
        rs!TR_TP_ID = FNN(.TPID)
        rs!TR_DATE = FNDateSerial(.DOCDate)
        rs!TR_CaptureDate = FND(.CaptureDate)
        rs!TR_CODE = FNS(.DOCCode)
        rs!TR_NOTE = FNS(.Memo)
        rs!TR_AddressDOCID = FNN(.BillToAddressID)
        rs!TR_AddressGoodsID = FNN(.GoodsAddressID)
        rs!TR_Status = .Status
        rs!TR_StaffID = FNN(.StaffID)
        rs!TR_SentToCentralAt = Null
        rs!TR_COMP_ID = oPC.Configuration.DefaultCOMPID
        rs.Update
        If .IsNew Then
           rs.MoveLast
           .TRID = rs!TR_ID
        End If
        rs.Close
    End With

    
    strSQL = "SELECT * FROM tCO WHERE CO_ID =" & udtProps.TRID
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    
    With udtProps
        If udtProps.IsNew Then rs.AddNew
        rs!CO_ID = .TRID
        rs!CO_OrderNum = FNS(.OrderNum)
        rs!CO_OrderType = FNN(.OrderType)
        If FNS(.ORGUID) > "" Then
            rs!CO_ORGUID = FNS(.ORGUID)
        End If
        rs.Update
        .IsNew = False
        .IsDirty = False
    End With
    rs.Close
    Set rs = Nothing
    
    Set objPBOut = New PropertyBag
    LSet udtData = udtProps
    Set objPersist = New a_COL_P
      With objPBOut
        .WriteProperty "State", udtData.buffer
        strBuffer = objPersist.Save(objPB.ReadProperty("COLS"), udtProps.TRID)
        If strBuffer > "" Then .WriteProperty "COLS", strBuffer
      End With
      Set objPB = Nothing
      Set objPersist = Nothing
    
      Save = objPBOut.Contents
      Set objPBOut = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    ErrPreserve
    If OpenResult = 0 Then oPC.DisconnectDBShort
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO_P.Save(buffer)", buffer, , rs
End Function


Private Sub DeleteInvoice(ID As Long)
Dim OpenResult As Integer

    On Error GoTo errHandler
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "DELETE FROM tPI WHERE PI_ID=" & ID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    ErrPreserve
    If OpenResult = 0 Then oPC.DisconnectDBShort
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO_P.DeleteInvoice(ID)", ID
End Sub

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler
Dim objPersist As a_COL_P
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.BeginTrans
    oPC.COShort.execute "DELETE FROM tCO WHERE CO_ID=" & ID
  
    Set objPersist = New a_COL_P
    objPersist.DeleteObject ID
    Set objPersist = Nothing
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    ErrPreserve
    If OpenResult = 0 Then oPC.DisconnectDBShort
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CO_P.DeleteObject(ID)", ID
End Sub

