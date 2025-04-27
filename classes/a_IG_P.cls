VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_IG_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Save(ByVal buffer As String, ByVal CustID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As IGProps
Dim udtData As IGData
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
    
    For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData

        If udtProps.IsDeleted Then
          DeleteIG udtProps.IGID, udtProps.TPID
        End If
    Next
    
    For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData
        If udtProps.IGID > 0 Then
            If Not udtProps.IsDeleted Then
                strSQL = "SELECT * FROM tTP_IG WHERE TPIG_TP_ID = " & CustID & " AND TPIG_IG_ID = " & udtProps.IGID
                rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
                If udtProps.IsNew Then rs.AddNew
                With rs
                    .fields("TPIG_TP_ID") = CustID
                    .fields("TPIG_IG_ID") = udtProps.IGID
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
        Else
            MsgBox ("Cannot add this item")
        End If
    Next
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
    ErrorIn "a_IG_P.Save(buffer)", buffer
End Function
Public Function Fetch(pTPID As Long, Optional pIGType As Long) As String
    On Error GoTo errHandler
Dim fs As FileSystemObject
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim mudtProps As IGProps
Dim mudtData As IGData
Dim objPB As PropertyBag
Dim lngCount As Long

    Set rs = New ADODB.Recordset
    rs.open "SELECT *  FROM tTP_IG JOIN tDICT ON TPIG_IG_ID =DICT_ID WHERE TPIG_TP_ID = " & pTPID & " AND DICT_TYPE = " & pIGType, oPC.COShort, adOpenDynamic, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.eof
      With rs
            mudtProps.IGID = .fields("TPIG_IG_ID")
            mudtProps.TPID = .fields("TPIG_TP_ID")
            mudtProps.Description = FNS(.fields("DICT_DESC"))
            mudtProps.IsLoyalty = FNN(.fields("TPIG_IG_ID")) = oPC.Configuration.LoyaltyClubTypeID
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
    ErrorIn "a_IG_P.Fetch"
End Function
Private Sub DeleteIG(IGID As Long, TPID As Long)
    On Error GoTo errHandler
Dim cnClient As Connection
Dim strSQL As String
    strSQL = "DELETE FROM tTP_IG WHERE TPIG_IG_ID=" & IGID & " AND TPIG_TP_ID = " & TPID
    oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_store_P.DeleteIG(IGID,TPID)", Array(IGID, TPID)
End Sub






