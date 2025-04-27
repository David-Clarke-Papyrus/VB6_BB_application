VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_ProductSection_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Save(ByVal buffer As String, ByVal ProductID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As ProductSectionProps
Dim udtData As ProductSectionData
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long

    Set objPB = New PropertyBag
    Set objPBOut = New PropertyBag
    arBuffer = buffer
    objPB.Contents = arBuffer
    Set rs = New ADODB.Recordset
    oPC.COShort.execute "set context_info 0x1"
    
    For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData
        If udtProps.IsDeleted Then
          DeleteProductSection udtProps.SECID, ProductID
        End If
    Next
    For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData
        If Not udtProps.IsDeleted Then
            strSQL = "SELECT * FROM tProductSection WHERE PSEC_P_ID = '" & ProductID & "' AND PSEC_SEC_ID = " & udtProps.SECID
            rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
   '         If udtProps.IsNew Then rs.AddNew
            If rs.eof Then rs.AddNew
            With rs
                .fields("PSEC_P_ID") = ProductID
                .fields("PSEC_SEC_ID") = udtProps.SECID
                .fields("PSEC_PRIORITY") = IIf(udtProps.Priority = 99, 99, 0)
                .Update
                udtProps.IsNew = False
                udtProps.IsDirty = False
                udtProps.IsDeleted = False
            End With
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.Close
        End If
    Next
    
    oPC.COShort.execute "set context_info 0x0"
    objPBOut.WriteProperty "Count", lngCount
    
    Set objPB = Nothing
    Set rs = Nothing
    
    Save = objPBOut.Contents
    Set objPBOut = Nothing
    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    Set objPBOut = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection_P.Save(buffer)", buffer
    'Resume
End Function
Public Function Fetch(pPID As String, Optional pPSECType As Long) As String
    On Error GoTo errHandler
Dim fs As FileSystemObject
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim mudtProps As ProductSectionProps
Dim mudtData As ProductSectionData
Dim objPB As PropertyBag
Dim lngCount As Long

    Set rs = New ADODB.Recordset
    If oPC.GetProperty("DeliveryStyle") = "BB" Then
        rs.open "SELECT tProductSection.*,DICT_SHORT + ' - ' + DICT_DESC as DICT_DESC,tDictTypes.System  FROM tProductSection JOIN tDICT ON PSEC_SEC_ID =DICT_ID  JOIN tDictTypes ON DICT_TYPE = ID " _
            & "WHERE PSEC_P_ID = '" & Trim$(pPID) & "' AND DICT_TYPE = " & pPSECType & " ORDER BY PSEC_PRIORITY DESC,DICT_DESC", oPC.COShort, adOpenDynamic, adLockReadOnly
    Else
        rs.open "SELECT tProductSection.*,DICT_DESC,tDictTypes.System  FROM tProductSection JOIN tDICT ON PSEC_SEC_ID =DICT_ID  JOIN tDictTypes ON DICT_TYPE = ID " _
            & "WHERE PSEC_P_ID = '" & Trim$(pPID) & "' AND DICT_TYPE = " & pPSECType & " ORDER BY PSEC_PRIORITY DESC,DICT_DESC", oPC.COShort, adOpenDynamic, adLockReadOnly
    End If
    Set objPB = New PropertyBag
    Do While Not rs.eof
      With rs
            mudtProps.SECID = .fields("PSEC_SEC_ID")
            mudtProps.PID = .fields("PSEC_P_ID")
            mudtProps.Priority = FNN(.fields("PSEC_PRIORITY"))
            mudtProps.Description = FNS(.fields("DICT_DESC"))
            mudtProps.typ = FNS(.fields("SYSTEM"))
'            mudtProps.IsLoyalty = FNN(.Fields("TPIG_IG_ID")) = oPC.Configuration.LoyaltyClubTypeID
'            mudtProps.IsLaunch = FNS(.Fields("DICT_SYSTEM")) = "LA"
'            mudtProps.IsLitLunch = FNS(.Fields("DICT_SYSTEM")) = "LL"
'            mudtProps.ISPromotion = FNS(.Fields("DICT_SYSTEM")) = "PR"
'            mudtProps.IsSale = FNS(.Fields("DICT_SYSTEM")) = "SA"
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
errHandler:
    ErrPreserve
    Set objPB = Nothing
    
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection_P.Fetch"
End Function
Private Sub DeleteProductSection(PSECID As Long, PID As String)
    On Error GoTo errHandler
Dim cnClient As Connection
Dim strSQL As String
    strSQL = "DELETE FROM tProductSection WHERE PSEC_SEC_ID=" & PSECID & " AND PSEC_P_ID = '" & PID & "'"
    oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_ProductSection_P.DeleteProductSection(PSECID,PID)", Array(PSECID, PID)
End Sub







