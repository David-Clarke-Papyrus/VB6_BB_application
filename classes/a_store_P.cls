VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_store_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As StoreProps
Dim udtData As StoreData
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set objPB = New PropertyBag
    Set objPBOut = New PropertyBag
    arBuffer = buffer
    objPB.Contents = arBuffer
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData

        If Not udtProps.IsDeleted Then
            strSQL = "SELECT * FROM tStore WHERE STORE_ID = " & udtProps.STID
            rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then rs.AddNew
            With rs
                .fields("STORE_Name") = FNS(udtProps.Description)
                .fields("STORE_BillToAddress") = FNS(udtProps.BillAddress)
                .fields("STORE_DeliverToAddress") = FNS(udtProps.DelAddress)
 '               .Fields("STORE_VPN_Address") = FNS(udtProps.VPNAddress)
                .fields("STORE_Code") = FNS(udtProps.code)
                .fields("STORE_SystemType") = FNS(udtProps.SystemName)
                .fields("STORE_ISACTIVE") = FNB(udtProps.IsActive)
                .fields("STORE_ISExternal") = FNB(udtProps.IsExternal)
                .Update
                If udtProps.IsNew Then
                  .Bookmark = .Bookmark
                  udtProps.STID = rs("STORE_ID")
                End If
                udtProps.IsNew = False
                udtProps.IsDirty = False
                udtProps.IsDeleted = False
            End With
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.Close
        Else
          DeleteSTORE udtProps.STID
    End If
    Next
    objPBOut.WriteProperty "Count", lngCount
    
    Set objPB = Nothing
    Set rs = Nothing
    
    Save = objPBOut.Contents
    Set objPBOut = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    Set objPBOut = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_store_P.Save(buffer)", buffer
End Function
Public Function Fetch() As String
    On Error GoTo errHandler
Dim fs As FileSystemObject
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim mudtProps As StoreProps
Dim mudtData As StoreData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    rs.open "SELECT *  FROM tStore", oPC.COShort, adOpenDynamic, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.eof
      With rs
            mudtProps.STID = .fields("STORE_ID")
            mudtProps.Description = FNS(.fields("STORE_Name"))
            mudtProps.BillAddress = FNS(.fields("STORE_BillToAddress"))
            mudtProps.DelAddress = FNS(.fields("STORE_DeliverToAddress"))
'            mudtProps.VPNAddress = FNS(.Fields("STORE_VPN_Address"))
            mudtProps.code = FNS(.fields("STORE_CODE"))
            mudtProps.SystemName = FNS(.fields("STORE_SystemType"))
            mudtProps.IsActive = FNB(.fields("STORE_ISACTIVE"))
            mudtProps.IsExternal = FNB(.fields("STORE_ISExternal"))
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
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_store_P.Fetch"
End Function
Private Sub DeleteSTORE(STID As Long)
    On Error GoTo errHandler
Dim cnClient As Connection
Dim strSQL As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "DELETE FROM tStore WHERE STORE_ID=" & STID
    oPC.COShort.execute strSQL
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_store_P.DeleteSTORE(STID)", STID
End Sub

Public Sub DeleteAllStores()
    On Error GoTo errHandler
'Dim cnClient As Connection
Dim strSQL As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "DELETE FROM tStoreP"
    oPC.COShort.execute strSQL

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_store_P.DeleteAllStores"
End Sub



Public Function StoreIndexClashes(strCode As String, pStoreID As Long) As Boolean
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_StoreClash"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@STORECODE", adVarChar, adParamInput, 3)
    cmd.Parameters.Append par
    par.Value = strCode
    Set par = cmd.CreateParameter("@ID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = pStoreID
    Set par = cmd.CreateParameter("@R", adInteger, adParamOutput)
    cmd.Parameters.Append par
    par.Value = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    StoreIndexClashes = (cmd.Parameters(2) > 0)
    Set cmd = Nothing

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function


End Function


