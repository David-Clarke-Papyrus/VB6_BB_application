VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_TPAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

    Set objPB = New PropertyBag
    Set objPBOut = New PropertyBag
    arBuffer = buffer
    objPB.Contents = arBuffer
    Set rs = New ADODB.Recordset
    
    For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData

        If Not udtProps.IsDeleted Then
            strSQL = "SELECT * FROM tStore WHERE STORE_ID = " & udtProps.STID
            rs.Open strSQL, oPC.CO, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then rs.AddNew
            With rs
                .Fields("STORE_Name") = FNS(udtProps.Description)
                .Fields("STORE_BillToAddress") = FNS(udtProps.BillAddress)
                .Fields("STORE_DeliverToAddress") = FNS(udtProps.DelAddress)
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
'ERRH:
'  Set objPB = Nothing
'  Set objPBOut = Nothing
'  Set rs = Nothing
'  Err.Raise Err.Number
    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    Set objPBOut = Nothing
    ErrRlsObjs rs
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

    Set rs = New ADODB.Recordset
    rs.Open "SELECT *  FROM tStore", oPC.CO, adOpenDynamic, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.EOF
      With rs
            mudtProps.STID = .Fields("STORE_ID")
            mudtProps.Description = FNS(.Fields("STORE_Name"))
            mudtProps.BillAddress = FNS(.Fields("STORE_BillToAddress"))
            mudtProps.DelAddress = FNS(.Fields("STORE_DeliverToAddress"))
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

'ERRH:
'  rs.Close
'  Set rs = Nothing
'  Err.Raise Err.Number
'  Resume
    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    
    ErrRlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_store_P.Fetch"
End Function
Private Sub DeleteSTORE(STID As Long)
    On Error GoTo errHandler
Dim cnClient As Connection
Dim strSQL As String
    strSQL = "DELETE FROM tStore WHERE STORE_ID=" & STID
    oPC.CO.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_store_P.DeleteSTORE(STID)", STID
End Sub

Public Sub DeleteAllStores()
    On Error GoTo errHandler
'Dim cnClient As Connection
Dim strSQL As String
  
    strSQL = "DELETE FROM tStoreP"
    oPC.CO.execute strSQL

'ERRH:
 ' Set cnClient = Nothing
'  If Not mobjContext Is Nothing Then mobjContext.SetAbort
'  Err.Raise Err.Number
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_store_P.DeleteAllStores"
End Sub






