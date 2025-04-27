VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_SCL_P"
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
Dim udtProps As RProps
Dim udtData As RData
Dim objPersist As a_RL_P
Dim objPB As PropertyBag

    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM R_Fetch JOIN tTP ON TR_TP_ID = TP_ID  WHERE TR_ID = " & pTRID
    '--------------
    OpenResult = oPC.OpenDBSHort
    '--------------
    
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.eof Then
        With udtProps
            .TRID = rs!TR_ID
            .DOCDate = FND(rs!TR_DATE)
            .DOCCode = FNS(rs!TR_CODE)
            .ApprovalRef = FNS(rs!R_ApprovalRef)
            .Memo = FNS(rs!TR_NOTE)
            .ApprovalTermDate = FND(rs!R_ApprovalTermDate)
            .TPID = FNN(rs!TR_TP_ID)
            .TPNAME = FNS(rs!TP_Name)
            .TPACCNum = FNS(rs!TP_ACno)
            .Status = FNStatus(rs!TR_Status)
            .CaptureDate = FND(rs!TR_CaptureDate)
            .StaffID = FNN(rs!TR_StaffID)
            .CurrencyID = FNN(rs!R_CURR_ID)
            .RType = FNS(rs!R_Type)
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            LSet udtData = udtProps
              Set objPersist = New a_RL_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.buffer
              .WriteProperty "RLS", objPersist.Fetch(udtProps.TRID)
              Fetch = .Contents
            End With
            Set objPB = Nothing
            Set objPersist = Nothing
            End With
    Else
        ' force an or
        rs.MoveNext
    End If
    '--------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
    '--------------
    Exit Function
EXIT_Handler:
    Exit Function
errHandler:

    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R_P.Fetch(pTRID)", pTRID
End Function

Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim objCode As New z_Code
Dim udtProps As RProps
Dim udtData As RData
Dim objPersist As a_RL_P
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
            udtProps.CaptureDate = Now
            udtProps.DOCDate = Date
        End If
        rs.fields("TR_Code") = FNS(udtProps.DOCCode)
        rs.fields("TR_CaptureDate") = FND(udtProps.CaptureDate)
        rs.fields("TR_Date") = FND(udtProps.DOCDate)
        rs.fields("TR_TP_ID") = FNN(udtProps.TPID)
        rs.fields("TR_Status") = FNN(udtProps.Status)
        rs.fields("TR_StaffID") = FNN(udtProps.StaffID)
        rs.fields("TR_Type") = 11
        
        rs.Update
        rs.MoveLast
        udtProps.TRID = rs.fields("TR_ID")
        rs.Close
    End With
        
    rs.open "SELECT * FROM tR WHERE R_ID = " & udtProps.TRID, oPC.COShort, adOpenDynamic, adLockOptimistic
    If udtProps.IsNew Then rs.AddNew
        rs.fields("R_ID") = udtProps.TRID
        rs.fields("R_TYPE") = FNS(udtProps.RType)
        rs.fields("R_ApprovalRef") = FNS(udtProps.ApprovalRef)
        rs.fields("R_ApprovalTermDate") = FND(udtProps.ApprovalTermDate)
        rs.fields("R_Curr_ID") = FND(udtProps.CurrencyID)
    rs.Update
    rs.Close


    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Set objPBOut = New PropertyBag
    LSet udtData = udtProps
    Set objPersist = New a_RL_P
    
    With objPBOut
        .WriteProperty "State", udtData.buffer
        .WriteProperty "RLS", objPersist.Save(objPB.ReadProperty("RLS"), udtProps.TRID)
    End With
    
    Set objPB = Nothing
    Set objPersist = Nothing
    
    Save = objPBOut.Contents
    Set objPBOut = Nothing
    
    
EXIT_Handler:

    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    Set objPersist = Nothing

    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R_P.Save(buffer)", buffer
End Function
'Public Function GetTILLID(pName As String) As Long
'Dim rsTILL As ADODB.Recordset
'Dim lngTILLID As Long
'    Set rsTILL = New ADODB.Recordset
'    rsTILL.Open "SELECT * FROM tTILL WHERE TILL_NAME = '" & FNS(pName) & "'", oPC.COShort, adOpenKeyset, adLockOptimistic
'    lngTILLID = 0
'    If Not rsTILL.EOF Then
'        lngTILLID = rsTILL!TILL_ID
'    End If
'    rsTILL.Close
'    GetTILLID = lngTILLID
'End Function

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler
  Dim objPersist As a_R_P
  
    oPC.COShort.BeginTrans
  
    Set objPersist = New a_R_P
    objPersist.DeleteObject ID
    Set objPersist = Nothing
    oPC.COShort.execute "DELETE FROM tR WHERE R_ID=" & ID
    oPC.COShort.CommitTrans
    Exit Sub
  
'H:
'    oPC.COShort.RollbackTrans
'    oor.Setor , "a_TF_P: DeleteObject " & or, Now, "", "", ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R_P.DeleteObject(ID)", ID
End Sub




