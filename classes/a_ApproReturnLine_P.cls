VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_APPRL_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByVal APPRID As Long) As String
    On Error GoTo errHandler
Dim rsARL As ADODB.Recordset
Dim strSQL As String
Dim udtProps As APPRLProps
Dim udtData As APPRLData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim oProdCode As z_ProdCode

    Set rsARL = New ADODB.Recordset
    strSQL = "SELECT * FROM APPRL_Fetch WHERE APPRL_TR_ID=" & APPRID
    rsARL.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rsARL.eof
        With udtProps
            .APPRLID = rsARL.fields("APPRL_ID")
            .APPLID = rsARL.fields("APPRL_APPL_ID")
            .Qty = FNN(rsARL.fields("APPRL_Qty"))
            .QtyIssued = FNN(rsARL.fields("APPL_Qty"))
            .QtyReturned = FNN(rsARL.fields("APPL_QtyReturned"))
            .Title = FNS(rsARL.fields("P_Title"))
            .PID = FNS(rsARL.fields("APPL_P_ID"))
            .ApproCode = FNS(rsARL.fields("TR_CODE"))
            .ApproDate = FND(rsARL.fields("TR_DATE"))
            .Fulfilled = FNS(rsARL.fields("APPRL_FULFILLED"))
            .code = rsARL.fields("CodeF")
            .CodeF = FNS(rsARL.fields("CODEFEX"))
            .EAN = FNS(rsARL.fields("P_EAN"))
    
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
        End With
        LSet udtData = udtProps
        lngCount = lngCount + 1
        objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
        rsARL.MoveNext
    Loop
    rsARL.Close
    Set rsARL = Nothing
    With objPB
          .WriteProperty "Count", CStr(lngCount)
          Fetch = .Contents
    End With
    Set objPB = Nothing

    Exit Function
errHandler:
    ErrPreserve
    RlsObjs rsARL
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL_P.Fetch(APPRID)", APPRID, , rsARL
End Function

Public Function Save(ByVal buffer As String, ByVal TRID As Long) As String
    On Error GoTo errHandler
  Dim rsARL As ADODB.Recordset
  Dim strSQL As String
  Dim udtProps As APPRLProps
  Dim udtData As APPRLData
  Dim objPB As PropertyBag
  Dim objPBOut As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngCount As Long
  Dim APPLID As Long
    
  Set rsARL = New ADODB.Recordset
    
    
  Set objPB = New PropertyBag
  Set objPBOut = New PropertyBag
  arBuffer = buffer
  objPB.Contents = arBuffer
  Set rsARL = New ADODB.Recordset
    
  For lngIndex = 1 To objPB.ReadProperty("Count")
    udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
    LSet udtProps = udtData
        
    If Not udtProps.IsDeleted Then
      strSQL = "SELECT * FROM tAPPRL WHERE APPRL_ID =" & udtProps.APPRLID
      rsARL.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
      If udtProps.IsNew Then rsARL.AddNew
              
      With udtProps
        rsARL!APPRL_TR_ID = TRID
        rsARL!APPRL_Qty = .Qty
        rsARL!APPRL_APPL_ID = .APPLID
        rsARL.Update
        
        If .IsNew Then
          rsARL.Bookmark = rsARL.Bookmark
          .APPRLID = rsARL("APPRL_ID")
        End If
        .IsNew = False
        .IsDirty = False
      End With
      LSet udtData = udtProps
      lngCount = lngCount + 1
      objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
      rsARL.Close
    Else
      DeleteApproReturnLine udtProps.APPRLID
    End If
  Next
  objPBOut.WriteProperty "Count", lngCount
  
  Set objPB = Nothing
  Set rsARL = Nothing
  
  Save = objPBOut.Contents
  Set objPBOut = Nothing
  Exit Function
  
    Exit Function
errHandler:
    ErrPreserve
    RlsObjs rsARL
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL_P.Save(buffer,TRID)", Array(buffer, TRID)
End Function

Public Sub DeleteObject(ByVal ApproRetID As Long)
    On Error GoTo errHandler
Dim strSQL As String
  
    strSQL = "DELETE FROM tAPPRL WHERE APPRL_TR_ID=" & CStr(ApproRetID)
    oPC.COShort.execute strSQL

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL_P.DeleteObject(ApproRetID)", ApproRetID
End Sub

Private Sub DeleteApproReturnLine(APPRLID As Long)
    On Error GoTo errHandler
Dim strSQL As String
    strSQL = "DELETE FROM tAPPRL WHERE APPRL_ID=" & CStr(APPRLID)
    oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPRL_P.DeleteApproReturnLine(APPRLID)", APPRLID
End Sub

