VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_COFF_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByVal COLID As Long, ByVal ILID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As COFFProps
Dim udtData As COFFData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim OpenResult As Integer
    
    Set rs = New ADODB.Recordset
    If COLID > 0 Then
        strSQL = "SELECT * FROM Possible_COL_IL_Matches WHERE IL_ID = " & ILID & " ORDER BY CODate"
    ElseIf ILID > 0 Then
        strSQL = "SELECT * FROM  Possible_IL_COL_Matches WHERE IL_ID = " & ILID ' & "' ORDER BY MC_Description"
    Else
        MsgBox "Missing parameter in a_COFF_P: Fetch"
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open strSQL, oPC.COShort, adOpenStatic, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.eof
    With udtProps
        .COFFCOLID = FNN(rs.fields("COFF_COL_ID"))
        .COFFID = FNN(rs.fields("COFF_ID"))
        .COFFCOLLID = FNN(rs.fields("COFF_COLL_ID"))
        .COFFILID = FNN(rs.fields("COFF_IL_ID"))
        .COFFQTY = FNN(rs.fields("QtyDispatchedThis"))
        .COFFCOCode = FNS(rs.fields("COCode"))
        .COFFINVCode = FNS(rs.fields("INVCode"))
        .COFFCOLQTY = FNN(rs.fields("COL_Qty"))
        .COFFILQTY = FNN(rs.fields("IL_Qty"))
        .COFFCode = FNS(rs.fields("P_Code"))
        .COLQty = FNN(rs.fields("COL_QTY"))
        .COLQtyDispatched = FNN(rs.fields("QtyDispatchedAll"))
        .CODate = FND(rs.fields("CODate"))
        .COCode = FNS(rs.fields("COCode"))
        .COLID = FNN(rs.fields("COL_ID"))

        .Status = FNB(rs.fields("COFF_Status"))
       
        .IsNew = False
        .IsDirty = False
        .IsDeleted = False
       
        End With
        LSet udtData = udtProps
        lngCount = lngCount + 1
        objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
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
    
EXIT_Handler:
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
    
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF_P.Fetch(COLID,ILID)", Array(COLID, ILID)
End Function


Public Function Save(ByVal buffer As String, ByVal pILID As Long, ByVal pCOLID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As COFFProps
Dim udtData As COFFData
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
Dim lngRecords As Long
  
  Set objPB = New PropertyBag
  Set objPBOut = New PropertyBag
  arBuffer = buffer
  objPB.Contents = arBuffer
  Set rs = New ADODB.Recordset
  
    For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData
        If Not udtProps.IsDeleted Then
            strSQL = "SELECT * FROM tCOFF WHERE COFF_ID= " & udtProps.COFFID
            rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
      
            If udtProps.IsNew Then rs.AddNew
            rs.fields("COFF_COL_ID") = udtProps.COFFCOLID
            rs.fields("COFF_COLL_ID") = udtProps.COFFCOLLID
            rs.fields("COFF_IL_ID") = pILID
            rs.fields("COFF_Qty") = udtProps.COFFQTY
            rs.fields("COFF_Status") = udtProps.Status
            rs.Update
            If udtProps.IsNew Then
                rs.Bookmark = rs.Bookmark
                udtProps.COFFID = rs("COFF_ID")
            End If
            udtProps.IsNew = False
            udtProps.IsDirty = False
            
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.Close
        Else
            DeleteCOFF udtProps.COFFID
        End If
    Next
    objPBOut.WriteProperty "Count", lngCount
    
    Set objPB = Nothing
    Set rs = Nothing
    
    Save = objPBOut.Contents
    Set objPBOut = Nothing
    Exit Function
EXIT_Handler:
    Exit Function
errHandler:
    ErrPreserve
    RlsObjs rs
    Set objPB = Nothing
    Set objPBOut = Nothing
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF_P.Save(buffer,pILID,pCOLID)", Array(buffer, pILID, pCOLID)
End Function

Private Sub DeleteCOFF(COFFID As Long)
    On Error GoTo errHandler
  Dim strSQL As String
  Dim lngRecords As Long
  
    strSQL = "DELETE FROM tCOFF WHERE COFF_ID=" & COFFID
    oPC.COShort.execute strSQL, lngRecords
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COFF_P.DeleteCOFF(COFFID)", COFFID
End Sub




