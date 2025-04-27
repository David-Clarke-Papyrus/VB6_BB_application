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
Attribute VB_Exposed = True
Option Explicit


Public Function Save(ByVal buffer As String, ByVal PID As String) As String
    On Error GoTo ErrHandler
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
            rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then rs.AddNew
            With rs
                .Fields("STP_DateLastStockTake") = FND(udtProps.DateLastStocktake)
                .Fields("STP_QtyLastStocktake") = FNN(udtProps.QtyLastStocktake)
                .Fields("STP_FirstReceived") = FND(udtProps.FirstReceivedDate)
                .Fields("STP_LastReceived") = FND(udtProps.LastDeliveredDate)
                .Fields("STP_QtyOnBACKORDER") = FNN(udtProps.QtyOnBackorder)
                .Fields("STP_QtyOnHand") = FNN(udtProps.QtyOnHand)
                .Fields("STP_QtyOnOrder") = FNN(udtProps.QtyonOrder)
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
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP_P.Save(buffer,pID)", Array(buffer, PID), , rs
End Function
Public Function FetchByStore(STID As Long) As String
    On Error GoTo ErrHandler
Dim fs As FileSystemObject
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim mudtProps As StorePProps
Dim mudtData As StorePData
Dim objPB As PropertyBag
Dim lngCount As Long

    Set rs = New ADODB.Recordset
    rs.Open "SELECT *  FROM tStoreP WHERE STP_ST_ID = " & STID, oPC.COShort, adOpenDynamic, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.EOF
      With rs
            mudtProps.STID = .Fields("STP_ST_ID")
            mudtProps.PID = .Fields("STP_P_ID")
            mudtProps.DateLastStocktake = FND(.Fields("STP_DateLastStockTake"))
            mudtProps.QtyLastStocktake = FNN(.Fields("STP_QtyLastStockTake"))
            mudtProps.FirstReceivedDate = FND(.Fields("STP_Firstreceived"))
            mudtProps.LastDeliveredDate = FND(.Fields("STP_LastDeliveredDate"))
            mudtProps.QtyOnHand = FNN(.Fields("STP_QtyOnHand"))
            mudtProps.QtyonOrder = FNN(.Fields("STP_QtyOnOrder"))
            mudtProps.QtyOnBackorder = FNN(.Fields("STP_QtyOnBackorder"))
            mudtProps.TotalSold = FNN(.Fields("STP_TotalSold"))
            mudtProps.LastDeliveredQty = FNN(.Fields("STP_LastQtyDelivered"))
            mudtProps.LastDeliveredDate = FND(.Fields("STP_LastDeliveredDate"))
            mudtProps.LastSoldDate = FND(.Fields("STP_LastDateSold"))
            mudtProps.StoreName = FNS(.Fields("STORE_Name"))
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
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP_P.FetchByStore(STID)", STID, , rs
End Function
Public Function FetchByProduct(PID As String) As String
    On Error GoTo ErrHandler
Dim fs As FileSystemObject
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim mudtProps As StorePProps
Dim mudtData As StorePData
Dim objPB As PropertyBag
Dim lngCount As Long

    Set rs = New ADODB.Recordset
    rs.Open "SELECT tStoreProduct.*,tStore.STORE_Name  FROM tStoreProduct JOIN tStore On STP_ST_ID = STORE_ID WHERE STP_P_ID = '" & PID & "' AND STORE_ISACTIVE = 1", oPC.COShort, adOpenDynamic, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.EOF
      With rs
            mudtProps.STID = .Fields("STP_ST_ID")
            mudtProps.PID = .Fields("STP_P_ID")
            mudtProps.SP = FND(.Fields("STP_SP"))
            mudtProps.QtyLastStocktake = FNN(.Fields("STP_QtyLastStockTake"))
            mudtProps.DateLastStocktake = FND(.Fields("STP_DateLastStocktake"))
            mudtProps.LastDeliveredPrice = FND(.Fields("STP_LastDeliveredPrice"))
            mudtProps.QtyOnHand = FNN(.Fields("STP_QTYOH"))
            mudtProps.QtyonOrder = FNN(.Fields("STP_QtyOnOrder"))
            mudtProps.QtyOnBackorder = FNN(.Fields("STP_QtyOnBackorder"))
            mudtProps.LastQtyFirmOrdered = FNN(.Fields("STP_LastQtyFirmOrdered"))
            mudtProps.LastQtySSOrdered = FNN(.Fields("STP_LastQtySSOrdered"))
            mudtProps.LastOrderedDate = FND(.Fields("STP_LastOrderedDate"))
            
            mudtProps.TotalSold = FNN(.Fields("STP_TotalSold"))
            mudtProps.LastDeliveredQty = FNN(.Fields("STP_LastQtyDelivered"))
            mudtProps.LastDeliveredDate = FND(.Fields("STP_LastDeliveredDate"))
            mudtProps.LastSoldDate = FND(.Fields("STP_LastDateSold"))
            
            
            mudtProps.StoreName = FNS(.Fields("STORE_Name"))
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
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP_P.FetchByProduct(pID)", PID, , rs
End Function
Private Sub DeleteSTOREP(PID As String, STID As Long)
    On Error GoTo ErrHandler
Dim strSQL As String
    strSQL = "DELETE FROM tStoreP WHERE STP_P_ID='" & PID & "' AND STP_ST_ID = " & STID
    oPC.COShort.execute strSQL
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP_P.DeleteSTOREP(pID,STID)", Array(PID, STID)
End Sub

Public Sub DeleteAllStorePs(ByVal PID As String)
    On Error GoTo ErrHandler
Dim strSQL As String
  
    strSQL = "DELETE FROM tStoreP WHERE STP_P_ID='" & PID & "'"
    oPC.COShort.execute strSQL
    Exit Sub
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_StoreP_P.DeleteAllStorePs(pID)", PID
End Sub




