VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_COLDELLAllocation_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GenerateCOLDELLAllocationset(Optional pDELID As Long, Optional pType As String)
Dim strSQL As String
Dim lngRecords As Long

On Error GoTo ERRH
    oPC.CO.BeginTrans
    If pDELID > 0 Then
        oPC.CO.execute "DELETE FROM tCOLDELLAlloc"
        strSQL = "INSERT INTO tCOLDELLAlloc (COLDELL_COL_ID,COLDELL_DELL_ID, COLDELL_Qty,COLDELL_Status,COLDELL_CO_ID) " _
             & " SELECT DISTINCT COL_ID,0,0,0,TR_ID FROM tCOL JOIN tProduct ON P_ID = COL_P_ID JOIN tTR On COL_TR_ID = TR_ID JOIN tDELL ON COL_P_ID = DELL_P_ID JOIN tTP ON TR_TP_ID = TP_ID JOIN tDICT on TP_CT_ID = DICT_ID WHERE " _
             & " COL_QTY > isnull(COL_QtyDispatched,0) and COL_FulFilled = 'OS' AND TR_STATUS = 3 AND P_QTYONHAND > 0 AND (DICT_DESC = '*Casual*' or DICT_DESC = '*Book club*') AND DELL_TR_ID = " & pDELID
    Else
        oPC.CO.execute "DELETE FROM tCOLDELLAllocation"
        strSQL = "INSERT INTO tCOLDELLAllocation (COLDELL_COL_ID,COLDELL_DELL_ID, COLDELL_Qty,COLDELL_Status,COLDELL_CO_ID) " _
             & " SELECT DISTINCT COL_ID,0,ISNULL(COL_QTY,0)-ISNULL(COL_QtyDispatched,0),0,TR_ID FROM tCOL JOIN tProduct ON P_ID = COL_P_ID JOIN tTR On COL_TR_ID = TR_ID WHERE " _
             & " COL_QTY > isnull(COL_QtyDispatched,0) and COL_FulFilled = 'OS' AND TR_STATUS = 3 AND P_QTYONHAND > 0"
    End If
    oPC.CO.execute strSQL, lngRecords
    oPC.CO.CommitTrans
    Exit Function
    '' JOIN tTP ON TR_TP_ID = TP_ID JOIN tDICT on TP_CT_ID = DICT_ID
ERRH:
    oPC.CO.RollbackTrans
  oError.SetError Err, "a_COLDELLAllocation_P:GenerateCOLDELLAllocationSet: " & Error, Now, "", "", ""
End Function

Public Function Save(ByVal buffer As String, pType As String) As String
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As COLDELLAllocationProps
Dim udtData As COLDELLAllocationData
Dim oPB As PropertyBag
Dim oPBOUT As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long

    Set oPB = New PropertyBag
    Set oPBOUT = New PropertyBag
    arBuffer = buffer
    oPB.Contents = arBuffer
    Set rs = New ADODB.Recordset
    
    On Error GoTo ERRH
    oPC.CO.BeginTrans
    For lngIndex = 1 To oPB.ReadProperty("Count")
        udtData.buffer = oPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData

        If Not udtProps.IsDeleted Then
            If pType = "DELIVERY" Then
                strSQL = "SELECT * FROM tCOLDELLAlloc WHERE COLDELL_ID = " & udtProps.ID
            Else
                strSQL = "SELECT * FROM tCOLDELLAllocation WHERE COLDELL_ID = " & udtProps.ID
            End If
            rs.Open strSQL, oPC.CO, adOpenDynamic, adLockOptimistic
            If udtProps.IsNew Then rs.AddNew
            With rs
                .Fields("COLDELL_COL_ID") = udtProps.COLID
                .Fields("COLDELL_Qty") = udtProps.AllocatedQty
                .Fields("COLDELL_Status") = udtProps.Status
                 .Update
                udtProps.IsNew = False
                udtProps.IsDirty = False
            End With
            LSet udtData = udtProps
            lngCount = lngCount + 1
            oPBOUT.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            If pType = "DELIVERY" Then
                oPC.CO.execute "UPDATE tPRODUCT SET P_QTYRESERVED = P_QTYRESERVED + " & udtProps.AllocatedQty & " WHERE P_ID = '" & udtProps.PID & "'"
                oPC.CO.execute "UPDATE tSTOREP SET STP_QTYRESERVED = STP_QTYRESERVED + " & udtProps.AllocatedQty & " WHERE STP_P_ID = '" & udtProps.PID & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
            End If
            oPC.CO.execute "UPDATE tCOL SET COL_QTYDispatched = COL_QtyDispatched + " & udtProps.AllocatedQty & " WHERE COL_ID = " & udtProps.COLID
            oPC.CO.execute "UPDATE tCOL SET COL_Fulfilled = 'FUL' WHERE dbo.NONNEGATIVE(COL_QTY - COL_QtyDispatched) = 0 AND COL_ID = " & udtProps.COLID
            strSQL = "UPDATE tTR Set TR_STATUS = 4 WHERE TR_ID IN " _
                        & " (SELECT TR_ID  FROM tCOL JOIN tTR ON tCOL.COL_TR_ID = tTR.TR_ID WHERE tTR.TR_ID = " & rs.Fields("COLDELL_CO_ID") _
                        & " GROUP BY TR_CODE ,TR_ID " _
                        & " Having (MAX(COL_Fulfilled) = 'CAN') or (MAX(COL_Fulfilled) = 'FUL'))"
            oPC.CO.execute strSQL
            rs.Close
        Else
          DeleteCOLDELL udtProps.ID, pType
    End If
    Next
    oPBOUT.WriteProperty "Count", lngCount
    
    Set oPB = Nothing
    Set rs = Nothing
    
    Save = oPBOUT.Contents
    Set oPBOUT = Nothing
    oPC.CO.CommitTrans
    Exit Function
ERRH:
    oPC.CO.RollbackTrans
    If Err = -2147217873 Then
        Save = ""
    Else
        oError.SetError Err, "a_COLDELLAllocation:Fetch:" & Error, Date, "", "", ""
    End If
  Set oPB = Nothing
  Set oPBOUT = Nothing
  Set rs = Nothing
End Function
Public Function Fetch(Optional TRID As Long, Optional pType As String) As String
Dim lngCount As Long
Dim udtProps As COLDELLAllocationProps
Dim udtData As COLDELLAllocationData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim dPOL As d_POLine
Dim lngID As Long
Dim objPB As PropertyBag
Dim strSQL As String

    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
    If TRID > 0 Then
        strSQL = "SELECT * FROM COLsOSPERDEL JOIN tDELL ON DELL_P_ID = COL_P_ID WHERE DELL_TR_ID = " & TRID
    Else
        strSQL = "SELECT * FROM COLsOS"
    End If
    rs.Open strSQL, oPC.CO, adOpenForwardOnly
    If rs.EOF Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    On Err GoTo ERRH
    Do While Not rs.EOF
        With udtProps
            udtProps.ID = FNN(rs!COLDELL_ID)
            udtProps.Ref = FNS(rs!COL_Ref)
            udtProps.QuotedPrice = FNN(rs!COL_Price)
            udtProps.OrderedQty = FNN(rs!COL_Qty)
            udtProps.DeliveredSoFar = FNN(rs!COL_QtyDispatched)
            udtProps.QtyOnHand = FNN(rs!P_QtyOnHand)
            udtProps.QtyReserved = FNN(rs!P_QtyReserved)
            udtProps.COLID = FNN(rs!COL_ID)
            udtProps.CustomerName = FNS(rs!TP_Name)
            udtProps.PID = FNS(rs!P_ID)
            udtProps.OrderCode = FNS(rs!TRCode)
            udtProps.Title = FNS(rs!P_Title)
            udtProps.OrderDate = FND(rs!TRDate)
            udtProps.AllocatedQty = FNN(rs!COLDELL_Qty)
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
        End With
    Loop
    With objPB
          .WriteProperty "Count", lngCount
          Fetch = .Contents
    End With
    Set objPB = Nothing
    
EXIT_Handler:
    rs.Close
    Set rs = Nothing
    Exit Function
ERRH:
  oError.SetError Err, "a_COLDELLAllocation:Fetch:" & Error, Date, "", "", ""
End Function
Private Sub DeleteCOLDELL(COLDELLAllocationID As Long, pType As String)
Dim strSQL As String
On Error GoTo ERRH
    If pType = "DELIVERY" Then
        strSQL = "DELETE FROM tCOLDELLAlloc WHERE COLDELL_ID=" & COLDELLAllocationID
    Else
        strSQL = "DELETE FROM tCOLDELLAllocation WHERE COLDELL_ID=" & COLDELLAllocationID
    End If
    oPC.CO.execute strSQL
    Exit Sub
ERRH:
  oError.SetError Err, "a_COLDELLAllocation:DeleteCOLDELL:" & Error, Date, "", "", ""
End Sub

Public Sub DeleteAllforDEL(pDELID As Long, pType As String)
Dim strSQL As String
On Error GoTo ERRH
    If pType = "DELIVERY" Then
        strSQL = "DELETE FROM tCOLDELLAlloc JOIN tDELL ON DELL_ID = COLDELL_DELL_ID WHERE DELL_TR_ID = " & pDELID
    Else
        strSQL = "DELETE FROM tCOLDELLAllocation JOIN tDELL ON DELL_ID = COLDELL_DELL_ID WHERE DELL_TR_ID = " & pDELID
    End If
    oPC.CO.execute strSQL
    Exit Sub
ERRH:
  oError.SetError Err, "a_COLDELLAllocation:DeleteAllforDEL:" & Error, Date, "", "", ""
End Sub


