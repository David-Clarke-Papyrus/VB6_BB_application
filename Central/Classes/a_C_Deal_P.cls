VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Deal_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function Save(ByVal buffer As String, pSupplierID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As DealProps
Dim udtData As DealData
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
            strSQL = "SELECT * FROM tDeal WHERE DL_ID = " & udtProps.ID
            rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then rs.AddNew
            With rs
                .Fields("DL_Description") = FNS(udtProps.Description)
                .Fields("DL_Rate") = Round(FNDBL(udtProps.Discount), 2)
                .Fields("DL_TP_ID") = FNN(pSupplierID)
                .Update
                If udtProps.IsNew Then
                    rs.Bookmark = rs.Bookmark
                    udtProps.ID = .Fields("DL_ID")
                End If
                udtProps.IsNew = False
                udtProps.IsDirty = False
            End With
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.Close
        Else
          DeleteDeal udtProps.ID
    End If
    Next
    objPBOut.WriteProperty "Count", lngCount
    
    Set objPB = Nothing
    Set rs = Nothing
    
    Save = objPBOut.Contents
    Set objPBOut = Nothing
'H:
'  Set objPB = Nothing
'  Set objPBOut = Nothing
'  Set rs = Nothing
'  Err.Raise err.Number
    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    Set objPBOut = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal_P.Save(buffer,pSupplierID)", Array(buffer, pSupplierID)
End Function
Public Function Fetch(ID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim mudtProps As DealProps
Dim mudtData As DealData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tDeal WHERE DL_TP_ID = " & ID, oPC.COShort, adOpenDynamic, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.EOF
      With rs
        mudtProps.ID = .Fields("DL_ID")
        mudtProps.Description = FNS(.Fields("DL_Description"))
        mudtProps.Discount = FNDBL(.Fields("DL_Rate"))
        mudtProps.TPID = FNS(.Fields("DL_TP_ID"))
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
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal_P.Fetch(ID)", ID
End Function
Public Sub DeleteDeal(DEALID As Long)
    On Error GoTo errHandler
Dim strSQL As String
  strSQL = "DELETE FROM tDeal WHERE DL_ID=" & DEALID
  oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal_P.DeleteDeal(DEALID)", DEALID
End Sub
Public Sub DeleteAllDeal(TPID As Long)
    On Error GoTo errHandler
Dim strSQL As String
  strSQL = "DELETE FROM tDeal WHERE DL_TP_ID=" & TPID
  oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Deal_P.DeleteAllDeal(TPID)", TPID
End Sub




