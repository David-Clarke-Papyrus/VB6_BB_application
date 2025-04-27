VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_DocumentControl_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As DCProps
Dim udtData As DCData
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
            strSQL = "SELECT * FROM tDocumentControl WHERE DC_ID = " & udtProps.ID
            rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then rs.AddNew
            With rs
              '  .Fields("DC_DOCCODE") = FNS(udtProps.DocCode)
                .fields("DC_PrinterID") = FNN(udtProps.PrinterID)
                .fields("DC_WSID") = FNN(udtProps.WSID)
                .fields("DC_QtyCopies") = FNN(udtProps.QtyCopies)
                .fields("DC_Style") = FNS(udtProps.Style)
                .fields("DC_PreviewPrint") = FNS(udtProps.PreviewPrint)
                .Update
                If udtProps.IsNew Then
                  .Bookmark = .Bookmark
                  udtProps.ID = rs("DC_ID")
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
          DeleteDC udtProps.ID
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
    ErrorIn "a_DocumentControl_P.Save(buffer)", buffer
End Function
Public Function Fetch(Optional pSMID As Long) As String
    On Error GoTo errHandler
Dim fs As FileSystemObject
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim mudtProps As DCProps
Dim mudtData As DCData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    If pSMID > 0 Then
        rs.open "SELECT *  FROM tDocumentControl JOIN tDICT ON DC_DocCode = DICT_System AND DC_WSID = " & pSMID & " AND LEFT OUTER JOIN tPRINTERS ON DC_PRINTERID = PRINT_ID, oPC.COShort, adOpenDynamic, adLockReadOnly"
    Else
        rs.open "SELECT *  FROM tDocumentControl JOIN tDICT ON DC_DocCode = DICT_System LEFT OUTER JOIN tPRINTERS ON DC_PRINTERID = PRINT_ID", oPC.COShort, adOpenDynamic, adLockReadOnly
    End If
    Set objPB = New PropertyBag
    Do While Not rs.eof
      With rs
            mudtProps.ID = .fields("DC_ID")
            mudtProps.PrinterID = FNN(.fields("DC_PrinterID"))
            mudtProps.PrinterName = FNS(.fields("PRINT_NAME"))
            mudtProps.DOCTypeName = FNS(.fields("DICT_DESC"))
            mudtProps.DOCCode = FNS(.fields("DC_DOCCODE"))
            mudtProps.PreviewPrint = FNS(.fields("DC_PreviewPrint"))
            mudtProps.QtyCopies = FNN(.fields("DC_QtyCopies"))
            mudtProps.Style = FNS(.fields("DC_Style"))
            mudtProps.WSID = FNN(.fields("DC_WSID"))
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

'H:
'  rs.Close
'  Set rs = Nothing
'  Err.Raise err.Number
'  Resume
    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DocumentControl_P.Fetch"
End Function
Private Sub DeleteDC(DCID As Long)
    On Error GoTo errHandler
Dim cnClient As Connection
Dim strSQL As String
    strSQL = "DELETE FROM tDocumentControl WHERE DC_ID=" & DCID
    oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DocumentControl_P.DeleteDC(DCID)", DCID
End Sub








