VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_CS_P"
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
Dim udtProps As CSProps
Dim udtData As CSData
Dim objPersist As a_CSL_P
Dim objPB As PropertyBag

    If pTRID = 0 Then Exit Function
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT * FROM CS_Fetch WHERE TR_ID = " & pTRID
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.eof Then
        With udtProps
            .TRID = rs!TR_ID
            .DateIssued = FND(rs!TR_DATE)
            .DOCCode = FNS(rs!TR_CODE)
            .TPID = FNN(rs!TR_TP_ID)
            .Status = FNS(rs!TR_Status)
            .DateStarted = FND(rs!TR_CaptureDate)
            .TILLID = FNN(rs!CS_TillPoint)
            .StaffID = FNN(rs!TR_StaffID)
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            LSet udtData = udtProps
              Set objPersist = New a_CSL_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.buffer
              .WriteProperty "CSLS", objPersist.Fetch(udtProps.TRID)
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
EXIT_Handler:

    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    Set objPersist = Nothing

    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS_P.Fetch(pTRID)", pTRID
End Function

Public Function FetchOpenTAForThisWorkstation(lngResult As Long, pDate As Date, pTillID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
  
Dim strSQL As String
Dim udtProps As CSProps
Dim udtData As CSData
Dim objPersist As a_CSL_P
Dim objPB As PropertyBag
Dim OpenResult As Integer


    strSQL = "SELECT tCS.*, tTR.*,tTILL.* FROM tCS JOIN tTR on CS_ID = TR_ID JOIN tTILL ON CS_TILLPOINT = TILL_ID WHERE TR_STATUS = 2 AND TILL_ID = " & pTillID '& "'" ' & " AND TR_DATE = '" & ReverseDate(pDate) & "'"
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    If rs.eof Then
        lngResult = 1
        GoTo EXIT_Handler
    Else
        lngResult = 0
    End If
    
    If Not rs.eof Then
        With udtProps
            .TRID = rs!TR_ID
            .DateIssued = FND(rs!TR_DATE)
            .DOCCode = FNS(rs!TR_CODE)
            .TPID = FNN(rs!TR_TP_ID)
            .Status = FNS(rs!TR_Status)
            .DateStarted = FND(rs!TR_CaptureDate)
            .TILLID = FNN(rs!CS_TillPoint)
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            LSet udtData = udtProps
              Set objPersist = New a_CSL_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.buffer
              .WriteProperty "CSLS", objPersist.Fetch(udtProps.TRID)
              FetchOpenTAForThisWorkstation = .Contents
            End With
            Set objPB = Nothing
            Set objPersist = Nothing
            End With
    Else
        ' force an or
        rs.MoveNext
    End If
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS_P.FetchOpenTAForThisWorkstation(lngResult,pDate,pTillID)", Array(lngResult, pDate, _
         pTillID)
    Exit Function
    Resume
End Function




Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim objCode As New z_Code
Dim udtProps As CSProps
Dim udtData As CSData
Dim objPersist As a_CSL_P
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim iDebug As Integer
Dim OpenResult As Integer
    Set objPB = New PropertyBag
    arBuffer = buffer
    With objPB
      .Contents = arBuffer
      udtData.buffer = .ReadProperty("State")
    End With
    LSet udtProps = udtData
    iDebug = 1
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT * FROM tTR WHERE TR_ID =" & udtProps.TRID
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    With udtProps
        If .IsNew Then
            rs.AddNew
            rs!TR_CaptureDate = Now
            rs!TR_DATE = Date
        End If
            rs.fields("TR_Code") = FNS(udtProps.DOCCode)
            rs.fields("TR_CaptureDate") = FND(udtProps.DateStarted)
            rs.fields("TR_Date") = FND(udtProps.DateIssued)
            rs.fields("TR_TP_ID") = FNN(udtProps.TPID)
            rs.fields("TR_Status") = FNN(udtProps.Status)
        rs.Update
        rs.MoveLast
        udtProps.TRID = rs.fields("TR_ID")
        rs.Close
    End With
    iDebug = 2
    rs.open "SELECT * FROM tCS WHERE CS_ID = " & udtProps.TRID, oPC.COShort, adOpenDynamic, adLockOptimistic
    If udtProps.IsNew Then rs.AddNew
        rs.fields("CS_TillPoint") = udtProps.TILLID
        rs.fields("CS_ID") = udtProps.TRID
    rs.Update
    rs.Close
    iDebug = 3
    Set rs = Nothing
    Set objPBOut = New PropertyBag
    LSet udtData = udtProps
    Set objPersist = New a_CSL_P
    With objPBOut
        .WriteProperty "State", udtData.buffer
        .WriteProperty "CSLS", objPersist.Save(objPB.ReadProperty("CSLS"), udtProps.TRID)
    End With
    Set objPB = Nothing
    Set objPersist = Nothing
    Save = objPBOut.Contents
    Set objPBOut = Nothing
    Exit Function
    iDebug = 4
EXIT_Handler:
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS_P.Save(buffer)", buffer
End Function
Public Function GetTILLID(pName As String) As Long
    On Error GoTo errHandler
Dim rsTILL As ADODB.Recordset
Dim lngTILLID As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rsTILL = New ADODB.Recordset
    rsTILL.open "SELECT * FROM tTILL WHERE TILL_NAME = '" & FNS(pName) & "'", oPC.COShort, adOpenKeyset, adLockOptimistic
    lngTILLID = 0
    If Not rsTILL.eof Then
        lngTILLID = rsTILL!TILL_ID
    End If
    rsTILL.Close
    GetTILLID = lngTILLID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    RlsObjs rsTILL
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS_P.GetTILLID(pName)", pName
End Function

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler
  Dim objPersist As a_CSL_P
  
    oPC.COShort.BeginTrans
  
    Set objPersist = New a_CSL_P
    objPersist.DeleteObject ID
    Set objPersist = Nothing
    oPC.COShort.execute "DELETE FROM tCS WHERE CS_ID=" & ID
    oPC.COShort.CommitTrans
    Exit Sub
  
'H:
'    oPC.COShort.RollbackTrans
'    oor.Setor , "a_APP_P: DeleteObject " & or, Now, "", "", ""
    Exit Sub
errHandler:
    ErrPreserve
    Set objPersist = Nothing
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CS_P.DeleteObject(ID)", ID
End Sub

