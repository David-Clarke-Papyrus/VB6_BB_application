VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_StoreP_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Save(ByVal buffer As String, ByVal PID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As StorePProps
Dim udtData As StorePData
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
            strSQL = "SELECT * FROM tStoreP WHERE STP_P_ID = '" & udtProps.PID & "'"
            rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then rs.AddNew
            With rs
                .fields("STP_DateLastStockTake") = FND(udtProps.DateLastStocktake)
                .fields("STP_QtyLastStocktake") = FNN(udtProps.QtyLastStocktake)
                .fields("STP_FirstReceived") = FND(udtProps.FirstReceivedDate)
                .fields("STP_LastReceived") = FND(udtProps.LastDeliveredDate)
                .fields("STP_QtyOnBACKORDER") = FNN(udtProps.QtyOnBackorder)
                .fields("STP_QtyOnHand") = FNN(udtProps.QtyOnHand)
'                .Fields("STP_QtyReserved") = FNN(udtProps.QtyReserved)
                .fields("STP_QtyOnOrder") = FNN(udtProps.QtyonOrder)
'                .Fields("STP_QtyOnAppro") = FNN(udtProps.QtyOnAppro)
                .Update
                udtProps.IsNew = False
                udtProps.IsDirty = False
            End With
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.Close
        Else
          DeleteSTOREP udtProps.PID, udtProps.STID
    End If
    Next
    objPBOut.WriteProperty "Count", lngCount
    
    Set objPB = Nothing
    Set rs = Nothing
    
    Save = objPBOut.Contents
    Set objPBOut = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP_P.Save(buffer,pID)", Array(buffer, PID), , rs
End Function
Public Function FetchByStore(STID As Long) As String
    On Error GoTo errHandler
Dim fs As FileSystemObject
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim mudtProps As StorePProps
Dim mudtData As StorePData
Dim objPB As PropertyBag
Dim lngCount As Long

    Set rs = New ADODB.Recordset
    rs.open "SELECT *  FROM tStoreP WHERE STP_ST_ID = " & STID, oPC.COShort, adOpenDynamic, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.eof
      With rs
            mudtProps.STID = .fields("STP_ST_ID")
            mudtProps.PID = .fields("STP_P_ID")
            mudtProps.DateLastStocktake = FND(.fields("STP_DateLastStockTake"))
            mudtProps.QtyLastStocktake = FNN(.fields("STP_QtyLastStockTake"))
            mudtProps.FirstReceivedDate = FND(.fields("STP_Firstreceived"))
            mudtProps.LastDeliveredDate = FND(.fields("STP_Lastreceived"))
            mudtProps.QtyOnHand = FNN(.fields("STP_QtyOnHand"))
     '       mudtProps.QtyReserved = FNN(.Fields("STP_QtyReserved"))
     '       mudtProps.QtyCopiesOnHand = FNN(.Fields("STP_QtyCopiesOnHand"))
            mudtProps.QtyonOrder = FNN(.fields("STP_QtyOnOrder"))
            mudtProps.QtyOnBackorder = FNN(.fields("STP_QtyOnBackorder"))
    '        mudtProps.QtyOnAppro = FNN(.Fields("STP_QtyOnAppro"))
            mudtProps.StoreName = FNS(.fields("STORE_Name"))
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
      FetchByStore = .Contents
    End With
    Set objPB = Nothing
    Exit Function

'H:
'  rs.Close
'  Set rs = Nothing
'  Err.Raise err.Number
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP_P.FetchByStore(STID)", STID, , rs
End Function
Public Function FetchByProduct(PID As String) As String
    On Error GoTo errHandler
Dim fs As FileSystemObject
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim mudtProps As StorePProps
Dim mudtData As StorePData
Dim objPB As PropertyBag
Dim lngCount As Long

    Set rs = New ADODB.Recordset
   ''' rs.Open "SELECT tStoreP.*,tStore.STORE_Name,tStore.STORE_Code  FROM tStoreP JOIN tStore On STP_ST_ID = STORE_ID WHERE STP_P_ID = '" & PID & "' AND STORE_ID <> " & oPC.Configuration.DefaultStoreID & " AND STORE_ISACTIVE = 1", oPC.COShort, adOpenDynamic, adLockReadOnly
    rs.open "SELECT tStoreP.*,tStore.STORE_Name,tStore.STORE_Code  FROM tStoreP JOIN tStore On STP_ST_ID = STORE_ID WHERE STP_P_ID = '" & PID & "' AND STORE_ISACTIVE = 1 AND ISNULL(STORE_IsExternal,1) = 0 AND STORE_ID <> " & oPC.Configuration.DefaultStoreID, oPC.COShort, adOpenDynamic, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.eof
      With rs
            mudtProps.STID = .fields("STP_ST_ID")
            mudtProps.PID = .fields("STP_P_ID")
            mudtProps.DateLastStocktake = FND(.fields("STP_DateLastStockTake"))
            mudtProps.QtyLastStocktake = FNN(.fields("STP_QtyLastStockTake"))
            mudtProps.FirstReceivedDate = FND(.fields("STP_Firstreceived"))
          '  mudtProps.LastSharedDate = FND(.Fields("STP_DATE"))
            mudtProps.LastDeliveredDate = FND(.fields("STP_Lastreceived"))
            mudtProps.QtyOnHand = FNN(.fields("STP_QTYOnHand")) ' & "(" & FND(.Fields("STP_DATELASTUPDATE")) & ")"
    '        mudtProps.QtyReserved = FNN(.Fields("STP_QtyReserved"))
    '        mudtProps.QtyCopiesOnHand = FNN(.Fields("STP_QtyCopiesOnHand"))
            mudtProps.QtyonOrder = FNN(.fields("STP_QtyOnOrder"))
            mudtProps.QtyOnBackorder = FNN(.fields("STP_QtyOnBackorder"))
            mudtProps.StoreName = FNS(.fields("STORE_Name"))
            mudtProps.StoreCode = FNS(.fields("STORE_Code"))
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
      FetchByProduct = .Contents
    End With
    Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP_P.FetchByProduct(pID)", PID, , rs
End Function
Private Sub DeleteSTOREP(PID As String, STID As Long)
    On Error GoTo errHandler
Dim strSQL As String
    strSQL = "DELETE FROM tStoreP WHERE STP_P_ID='" & PID & "' AND STP_ST_ID = " & STID
    oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP_P.DeleteSTOREP(pID,STID)", Array(PID, STID)
End Sub

Public Sub DeleteAllStorePs(ByVal PID As String)
    On Error GoTo errHandler
Dim strSQL As String
  
    strSQL = "DELETE FROM tStoreP WHERE STP_P_ID='" & PID & "'"
    oPC.COShort.execute strSQL
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP_P.DeleteAllStorePs(pID)", PID
End Sub




