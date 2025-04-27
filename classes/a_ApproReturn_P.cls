VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_APPR_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(TRID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
  
Dim strSQL As String
Dim udtProps As APPRProps
Dim udtData As APPRData
Dim objPersist As a_APPRL_P
Dim objPB As PropertyBag
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT * FROM APPR_Fetch WHERE TR_ID = " & TRID
    
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    If Not rs.eof Then
        With udtProps
            .TRID = FNN(rs!TR_ID)
            .TPNAME = FNS(rs!TP_Name)
            .DOCDate = FND(rs!TR_DATE)
            .CaptureDate = FND(rs!TR_CaptureDate)
            .Memo = FNS(rs!TR_NOTE)
            .DOCCode = FNS(rs!TR_CODE)
            .TPID = FNN(rs!TR_TP_ID)
            .Status = FNStatus(rs!TR_Status)
            .StaffID = FNN(rs!TR_StaffID)
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            LSet udtData = udtProps
              Set objPersist = New a_APPRL_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.buffer
              .WriteProperty "ARLS", objPersist.Fetch(udtProps.TRID)
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
    ErrorIn "a_APPR_P.Fetch(TRID)", TRID, , rs
End Function

Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
'Dim cmd As New ADODB.Command
Dim strSQL As String
Dim objCode As New z_Code
Dim udtProps As APPRProps
Dim udtData As APPRData
Dim objPersist As a_APPRL_P
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim code As String
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
    With udtProps
        SaveApproReturn .IsNew, .TRID, .TPID, .DOCDate, FNS(.DOCCode), FNS(.Memo), .Status, .StaffID
        .IsNew = False
        
        .IsDirty = False
    End With

   
    Set objPBOut = New PropertyBag
    LSet udtData = udtProps
    Set objPersist = New a_APPRL_P
      With objPBOut
        .WriteProperty "State", udtData.buffer
        .WriteProperty "ARLS", _
        objPersist.Save(objPB.ReadProperty("ARLS"), udtProps.TRID)
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
    ErrorIn "a_APPR_P.Save(buffer)", buffer
End Function

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler
  Dim objPersist As a_APPRL_P
  Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.BeginTrans
  
    Set objPersist = New a_APPRL_P
    objPersist.DeleteObject ID
    Set objPersist = Nothing
    oPC.COShort.execute "DELETE FROM tAPPR WHERE APPR_ID=" & ID
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    ErrPreserve
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP_P.DeleteObject(ID)", ID
End Sub


