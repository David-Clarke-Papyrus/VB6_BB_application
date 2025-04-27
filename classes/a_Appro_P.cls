VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_APP_P"
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
Dim udtProps As APPProps
Dim udtData As APPData
Dim objPersist As a_APPL_P
Dim objPB As PropertyBag

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT * FROM APP_Fetch WHERE TR_ID = " & TRID
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.eof Then
        With udtProps
            .TRID = rs!TR_ID
            .TPNAME = FNS(rs!TP_Name)
            .DOCDate = FND(rs!TR_DATE)
            .CaptureDate = FND(rs!TR_CaptureDate)
            .DOCCode = FNS(rs!TR_CODE)
            .TPID = FNN(rs!TR_TP_ID)
            .APPROTOID = FNN(rs!AddressDOCID)
            .COMPID = FNN(rs!TR_COMP_ID)
            .Status = FNStatus(rs!TR_Status)
            .Memo = FNS(rs!TR_NOTE)
            .StaffID = FNN(rs!TR_StaffID)
            .VATable = FNB(rs!TP_Vatable)
            .ShowVAT = FNB(rs!TP_ShowVAT)
            .NonVATDocument = FNB(rs!TR_NonVATDocument)
            
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            LSet udtData = udtProps
            Set objPersist = New a_APPL_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.buffer
              .WriteProperty "ALS", objPersist.Fetch(udtProps.TRID)
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
    ErrorIn "a_APP_P.Fetch(TRID)", TRID, , rs
End Function

Public Function Save(ByVal buffer As String, bMadeChange As Boolean) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim objCode As New z_Code
Dim udtProps As APPProps
Dim udtData As APPData
Dim objPersist As a_APPL_P
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
        If .IsNew Then
            rs.AddNew
            rs!TR_CaptureDate = Now
        End If
        rs!TR_Type = 9
        rs!TR_TP_ID = .TPID
        If CDate(rs!TR_DATE) = 0 And .Status > 2 Then
            rs!TR_DATE = Date
        Else
            rs!TR_DATE = .DOCDate
        End If
        rs!TR_CODE = FNS(.DOCCode)
        rs!TR_NOTE = FNS(.Memo)
        rs!TR_Status = .Status
        rs!TR_AddressDOCID = .APPROTOID
        rs!TR_COMP_ID = .COMPID
        rs!TR_StaffID = .StaffID
        rs.Update
        If .IsNew Then
           rs.MoveLast
           .TRID = rs!TR_ID
        End If
        rs.Close
   End With

    strSQL = "SELECT * FROM tAPP WHERE APP_ID =" & udtProps.TRID
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    
    With udtProps
        If udtProps.IsNew Then rs.AddNew
        rs!APP_ID = .TRID
                
        rs.Update
        .IsNew = False
        .IsDirty = False
    End With
    rs.Close
    Set rs = Nothing
    
    Set objPBOut = New PropertyBag
    LSet udtData = udtProps
    Set objPersist = New a_APPL_P
      With objPBOut
        .WriteProperty "State", udtData.buffer
        .WriteProperty "ALS", _
        objPersist.Save(objPB.ReadProperty("ALS"), udtProps.TRID)
      End With
      Set objPB = Nothing
      Set objPersist = Nothing
      ConsolidateApproLines udtProps.TRID, bMadeChange
'      If bMadeChange Then  'we must reload he lines because some have been consolidated
'        Save = Fetch(udtProps.TRID)
'      Else
        Save = objPBOut.Contents
'      End If
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
    ErrorIn "a_APP_P.Save(buffer)", buffer
End Function

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler
  Dim objPersist As a_APPL_P
  
    oPC.COShort.BeginTrans
  
    Set objPersist = New a_APPL_P
    objPersist.DeleteObject ID
    Set objPersist = Nothing
    oPC.COShort.execute "DELETE FROM tAPP WHERE APP_ID=" & ID
    oPC.COShort.CommitTrans
    Exit Sub
  
    Exit Sub
errHandler:
    ErrPreserve
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP_P.DeleteObject(ID)", ID
End Sub

Private Sub ConsolidateApproLines(TRID As Long, bMadeChange As Boolean)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long

    Set cmd = New ADODB.Command
    cmd.CommandText = "ConsolidateApproLines"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("Return", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TRID", adInteger, adParamInput, , TRID)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    iReturn = cmd.Parameters("Return")
    If iReturn > 1 Then Err.Raise EXC_SP_FAILED, "a_APP_P:ConsolidateLines", "SP failed"
    bMadeChange = (iReturn = 1)
    
    Set cmd = Nothing

Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP_P.ConsolidateLines(TRID,bMadeChange)", Array(TRID, bMadeChange)
End Sub

