VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Offer_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByVal PID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As WantProps
Dim udtData As WantData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tWants JOIN tTP on WT_TP_ID = TP_ID WHERE WT_P_ID= '" & PID & "'"
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.eof
        With udtProps
            .ID = rs.fields("WT_ID")
            .PID = rs.fields("WT_P_ID")
            .TPID = rs.fields("WT_TP_ID")
            .RequestDate = FND(rs.fields("WT_REQDATE"))
            .CustomerName = FNS(rs.fields("TP_Name"))
            .Notes = FNS(rs.fields("WT_Note"))
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
          .WriteProperty "Count", CStr(lngCount)
          Fetch = .Contents
    End With
    Set objPB = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Offer_P.Fetch(PID)", PID
End Function



Public Function Save(ByVal buffer As String, ByVal PID As String) As String
    On Error GoTo errHandler
  Dim rs As ADODB.Recordset
  Dim strSQL As String
  Dim udtProps As WantProps
  Dim udtData As WantData
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
      strSQL = "SELECT * FROM tWants WHERE WT_P_ID= '" & udtProps.PID & "'"
      rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
      If udtProps.IsNew Then rs.AddNew
      
      With udtProps
        rs.fields("WT_P_ID") = Trim$(.PID)
        rs.fields("WT_TP_ID") = .TPID
        rs.fields("WT_ReqDate") = .RequestDate
        rs.fields("WT_Note") = Trim$(.Notes)
        rs.Update
        
        If .IsNew Then
          rs.Bookmark = rs.Bookmark
          .ID = rs("WT_ID")
        End If
        .IsNew = False
        .IsDirty = False
      End With
      LSet udtData = udtProps
      lngCount = lngCount + 1
      objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
      rs.Close
    Else
      DeleteWant udtProps.ID
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
    ErrorIn "a_Offer_P.Save(buffer,PID)", Array(buffer, PID)
End Function

Public Sub DeleteObject(ByVal PID As String)
    On Error GoTo errHandler
    oPC.COShort.execute "DELETE FROM tWants WHERE WT_P_ID='" & PID & "'"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Offer_P.DeleteObject(PID)", PID
End Sub

Private Sub DeleteWant(ID As Long)
    On Error GoTo errHandler
    oPC.COShort.execute "DELETE FROM tWants WHERE WT_ID = " & ID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Offer_P.DeleteWant(ID)", ID
End Sub




