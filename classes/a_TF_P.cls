VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_TF_P"
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
Dim udtProps As TFProps
Dim udtData As TFData
Dim objPersist As a_TFL_P
Dim objPB As PropertyBag

    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM TFR_Fetch JOIN tSTORE ON TR_TP_ID = STORE_ID  WHERE TR_ID = " & pTRID
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.eof Then
        With udtProps
            .TRID = rs!TR_ID
            .DOCDate = FND(rs!TR_DATE)
            .DOCCode = FNS(rs!TR_CODE)
            .DestID = FNN(rs!TR_TP_ID)
            .DestinationName = FNS(rs!STORE_Name)
            .Status = FNStatus(rs!TR_Status)
            .InOut = FNS(rs!TFR_INOUT)
            .CaptureDate = FND(rs!TR_CaptureDate)
            .Auto = FNB(rs!TFR_Auto)
            .Memo = FNS(rs!TR_NOTE)
            .StaffID = FNN(rs!TR_StaffID)
            .SendersDocDate = FND(rs!TFR_SendersDocDate)
            .SendersDocRef = FNS(rs!TFR_SendersDocRef)
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            LSet udtData = udtProps
              Set objPersist = New a_TFL_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.buffer
              .WriteProperty "TFLS", objPersist.Fetch(udtProps.TRID)
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
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TF_P.Fetch(pTRID)", pTRID
End Function

'Public Function FetchOpenTAForThisWorkstation(lngResult As Long, pDate As Date, pTillID As Long) As String
'On Error GoTo err_Handler
'Dim rs As ADODB.Recordset
'
'Dim strSQL As String
'Dim udtProps As CSProps
'Dim udtData As CSData
'Dim objPersist As a_CSL_P
'Dim objPB As PropertyBag
'
'    strSQL = "SELECT tCS.*, tTR.*,tTILL.* FROM tCS JOIN tTR on CS_ID = TR_ID JOIN tTILL ON CS_TILLPOINT = TILL_ID WHERE TR_STATUS = 2 AND TILL_ID = " & pTillID '& "'" ' & " AND TR_DATE = '" & ReverseDate(pDate) & "'"
'    Set rs = New ADODB.Recordset
'    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
'    If rs.EOF Then
'        lngResult = 1
'        GoTo EXIT_Handler
'    Else
'        lngResult = 0
'    End If
'
'    If Not rs.EOF Then
'        With udtProps
'            .TRID = rs!TR_ID
'            .DateIssued = FND(rs!TR_Date)
'            .DocCode = FNS(rs!TR_Code)
'            .TPID = FNN(rs!TR_TP_ID)
'            .Status = FNS(rs!TR_Status)
'            .DateStarted = FND(rs!TR_CaptureDate)
'            .TILLID = FNN(rs!CS_TillPoint)
'            .IsNew = False
'            .IsDirty = False
'            .IsDeleted = False
'            rs.Close
'            Set rs = Nothing
'            LSet udtData = udtProps
'              Set objPersist = New a_CSL_P
'            Set objPB = New PropertyBag
'            With objPB
'              .WriteProperty "State", udtData.buffer
'              .WriteProperty "CSLS", objPersist.Fetch(udtProps.TRID)
'              FetchOpenTAForThisWorkstation = .Contents
'            End With
'            Set objPB = Nothing
'            Set objPersist = Nothing
'            End With
'    Else
'        ' force an or
'        rs.MoveNext
'    End If
'
'  Exit Function
'EXIT_Handler:
'    Exit Function
'Err_Handler:
'    tmp =
'    tmpor = or
'    oor.Setor tmp, tmpor, Now(), "a_CS", "Fetch", "unknown"
'    Err.Raise tmp
'    GoTo EXIT_Handler
'End Function
'



Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim objCode As New z_Code
Dim udtProps As TFProps
Dim udtData As TFData
Dim objPersist As a_TFL_P
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
        
        
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tTR WHERE TR_ID =" & udtProps.TRID
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
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
        rs.fields("TR_TP_ID") = FNN(udtProps.DestID)
        rs.fields("TR_Status") = FNN(udtProps.Status)
        rs.fields("TR_StaffID") = FNN(udtProps.StaffID)
        rs.fields("TR_Type") = 1
        rs.fields("TR_COMP_ID") = oPC.Configuration.DefaultCOMPID

        
        rs.Update
        If udtProps.IsNew Then
          rs.Bookmark = rs.Bookmark
          udtProps.TRID = rs("TR_ID")
        End If
      '  rs.MoveLast
      '  udtProps.TRID = rs.Fields("TR_ID")
        rs.Close
    End With
        
    rs.open "SELECT * FROM tTFR WHERE TFR_ID = " & udtProps.TRID, oPC.COShort, adOpenDynamic, adLockOptimistic
    If udtProps.IsNew Then rs.AddNew
        rs.fields("TFR_ID") = udtProps.TRID
        rs.fields("TFR_INOUT") = udtProps.InOut
        rs.fields("TFR_SendersDocDate") = FND(udtProps.SendersDocDate)
        rs.fields("TFR_SendersDocRef") = FNS(udtProps.SendersDocRef)
    rs.Update
    rs.Close


    Set rs = Nothing
    
    Set objPBOut = New PropertyBag
    LSet udtData = udtProps
    Set objPersist = New a_TFL_P
    
    With objPBOut
        .WriteProperty "State", udtData.buffer
        .WriteProperty "TFLS", objPersist.Save(objPB.ReadProperty("TFLS"), udtProps.TRID)
    End With
    
    Set objPB = Nothing
    Set objPersist = Nothing
    
    Save = objPBOut.Contents
    Set objPBOut = Nothing
    
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
    
    
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TF_P.Save(buffer)", buffer
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
  Dim objPersist As a_TFL_P
  
    oPC.COShort.BeginTrans
  
    Set objPersist = New a_TFL_P
    objPersist.DeleteObject ID
    Set objPersist = Nothing
    oPC.COShort.execute "DELETE FROM tTFR WHERE TFR_ID=" & ID
    oPC.COShort.CommitTrans
    Exit Sub
  
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TF_P.DeleteObject(ID)", ID
End Sub


