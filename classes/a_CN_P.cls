VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_CN_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function Fetch(TRID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
  
Dim strSQL As String
Dim udtProps As CNProps
Dim udtData As CNData
Dim objPersist As a_CNL_P
Dim objPB As PropertyBag
Dim oUtil As z_SQL

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
    If OpenResult = 1 Then
        Forms(0).mnuExit_Click
        Exit Function
    End If
'-------------------------------
    Set oUtil = New z_SQL
    lngResult = oUtil.RunGetRecordset("SELECT * FROM Fetch_CN WHERE TR_ID = " & TRID, enText, Array(), "", rs)
    
    
    If Not rs.eof Then
        With udtProps
            .TRID = rs!TR_ID
            .TPID = FNN(rs!TR_TP_ID)
            .COMPID = FNN(rs!TR_COMP_ID)
           ' If FNN(rs.fields("AddressDOCID")) > 0 Then
                .BillToAddressID = FNN(rs.fields("AddressDOCID"))
          '  End If
            .TPACCNum = FNS(rs!TP_ACno)
            .TPFax = FNS(rs!ADD_Fax)
            .TPPhone = FNS(rs!TP_Phone)
            .TPNAME = FNS(rs!TP_Name)
            .TPMemo = FNS(rs!TR_NOTE)
            .DOCCode = FNS(rs!TR_CODE)
            .DOCDate = FND(rs!TR_DATE)
            .CaptureDate = FND(rs!TR_CaptureDate)
            .CurrRate = FNDBL(rs!CN_CurrRate)
            .ForeignCurrencyID = FNN(rs!CN_CURR_ID)
            .Status = FNStatus(rs!TR_Status)
            .StaffID = FNN(rs!TR_StaffID)
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            Set oUtil = Nothing
            LSet udtData = udtProps
              Set objPersist = New a_CNL_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.buffer
              .WriteProperty "CNLS", objPersist.Fetch(udtProps.TRID)
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
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN_P.Fetch(TRID)", TRID
End Function

Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
'Dim cmd As New ADODB.Command
Dim strSQL As String
Dim objCode As New z_Code
Dim udtProps As CNProps
Dim udtData As CNData
Dim objPersist As a_CNL_P
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte

    Set objPB = New PropertyBag
    arBuffer = buffer
    With objPB
      .Contents = arBuffer
      udtData.buffer = .ReadProperty("State")
    End With
    LSet udtProps = udtData

    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tTR WHERE TR_ID =" & udtProps.TRID
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    With udtProps
        If .IsNew Then rs.AddNew
        rs!TR_Type = 8
        rs!TR_TP_ID = FNN(.TPID)
        rs!TR_COMP_ID = FNN(.COMPID)
        rs!TR_DATE = FND(.DOCDate)
        rs!TR_CaptureDate = FND(.CaptureDate)
        rs!TR_CODE = FNS(.DOCCode)
        rs!TR_NOTE = FNS(.TPMemo)
        rs!TR_AddressDOCID = FNN(.BillToAddressID)
        rs!TR_AddressGoodsID = FNN(.DelToAddressID)
        rs!TR_Status = .Status
        rs!TR_StaffID = .StaffID
        rs.Update
        If .IsNew Then
           rs.MoveLast
           .TRID = rs!TR_ID
        End If
        rs.Close
    End With

    strSQL = "SELECT * FROM tCN WHERE CN_ID =" & udtProps.TRID
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    
    With udtProps
        If udtProps.IsNew Then rs.AddNew
        rs!CN_ID = .TRID
        rs!CN_CreditValue = CDbl(.TotalPayable) / oPC.Configuration.DefaultCurrency.Divisor
        rs!CN_VATVALUE = CDbl(.TotalVAT) / oPC.Configuration.DefaultCurrency.Divisor
        rs.Update
        .IsNew = False
        .IsDirty = False
    End With
    rs.Close
    Set rs = Nothing
    
    Set objPBOut = New PropertyBag
    LSet udtData = udtProps
    Set objPersist = New a_CNL_P
      With objPBOut
        .WriteProperty "State", udtData.buffer
        .WriteProperty "CNLS", objPersist.Save(objPB.ReadProperty("CNLS"), udtProps.TRID)
      End With
      Set objPB = Nothing
      Set objPersist = Nothing
    
      Save = objPBOut.Contents
      Set objPBOut = Nothing
      
      Exit Function

'H:
'    rs.Close
'    Set rs = Nothing
'    oor.Setor , "a_CN_P: Save " & or, Now, "", "", ""
    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN_P.Save(buffer)", buffer
End Function


Private Sub DeleteCN(ID As Long)
    On Error GoTo errHandler
    oPC.COShort.execute "DELETE FROM tCN WHERE CN_ID=" & ID

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN_P.DeleteCN(ID)", ID
End Sub

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler
  Dim objPersist As a_CNL_P
  
    oPC.COShort.BeginTrans
  
    Set objPersist = New a_CNL_P
    objPersist.DeleteObject ID
    Set objPersist = Nothing
    oPC.COShort.execute "DELETE FROM tCN WHERE CN_ID=" & ID
    oPC.COShort.CommitTrans
    Exit Sub
  
'H:
'    oPC.COShort.RollbackTrans
'    oor.Setor , "a_CN_P: DeleteObject " & or, Now, "", "", ""
    Exit Sub
errHandler:
    ErrPreserve
    Set objPersist = Nothing
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN_P.DeleteObject(ID)", ID, EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Sub

