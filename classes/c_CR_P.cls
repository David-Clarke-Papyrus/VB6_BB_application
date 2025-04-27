VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_CR_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Save(ByVal buffer As String, ByVal SMID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As CRProps
Dim udtData As CRData
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
Dim OpenResult As Integer


    Set objPB = New PropertyBag
    Set objPBOut = New PropertyBag
    arBuffer = buffer
    objPB.Contents = arBuffer
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData

        If Not udtProps.IsDeleted Then
            strSQL = "SELECT * FROM tCR WHERE CR_SM_ID = " & SMID & " AND CR_PT_ID = " & udtProps.PTID
            rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then rs.AddNew
            With rs
                .fields("CR_SM_ID") = SMID
                .fields("CR_PT_ID") = udtProps.PTID
                .Update
                udtProps.IsNew = False
                udtProps.IsDirty = False
                udtProps.IsDeleted = False
            End With
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.Close
        Else
          DeleteCR udtProps.SMID, udtProps.PTID
    End If
    Next
    objPBOut.WriteProperty "Count", lngCount
    
    Set objPB = Nothing
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Save = objPBOut.Contents
    Set objPBOut = Nothing
    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    Set objPBOut = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_CR_P.Save(buffer)", buffer
End Function
Public Function Fetch(SMID As Long, PTID As Long) As String
    On Error GoTo errHandler
Dim fs As FileSystemObject
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim mudtProps As CRProps
Dim mudtData As CRData
Dim objPB As PropertyBag
Dim lngCount As Long

    Set rs = New ADODB.Recordset
    rs.open "SELECT *  FROM tCR_IG WHERE CR_SM_ID = " & SMID & " AND CR_PT_ID = " & PTID, oPC.COShort, adOpenDynamic, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.eof
      With rs
            mudtProps.SMID = .fields("CR_SM_ID")
            mudtProps.PTID = .fields("CR_PT_ID")
            mudtProps.CommissionRate = FNDBL(.fields("CR_RATE"))
            mudtProps.IsDirty = False
            mudtProps.IsNew = False
            mudtProps.IsDeleted = False
      End With
      LSet mudtData = mudtProps
      lngCount = lngCount + 1
      objPB.WriteProperty "Item" & CStr(lngCount), mudtData.buffer
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    With objPB
      .WriteProperty "Count", lngCount
      Fetch = .Contents
    End With
    Set objPB = Nothing
    Exit Function

    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "C_CR_P.Fetch"
End Function
Private Sub DeleteCR(SMID As Long, PTID As Long)
    On Error GoTo errHandler
Dim cnClient As Connection
Dim strSQL As String
    strSQL = "DELETE FROM tCR WHERE CR_SM_ID=" & SMID & " AND CR_PT_ID = " & PTID
    oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_CR_P_P.DeleteCR(SMID,PTID)", Array(SMID, PTID)
End Sub







