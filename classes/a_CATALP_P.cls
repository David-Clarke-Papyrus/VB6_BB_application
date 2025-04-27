VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_CATALP_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByVal PID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As CATALPProps
Dim udtData As CATALPData
Dim objPB As PropertyBag
Dim lngCount As Long

    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tCATALP JOIN tCatalogue on CATALP_CATAL_ID = CATAL_ID WHERE CATALP_P_ID= '" & PID & "'"
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.eof
        With udtProps
            .ID = rs.fields("CATALP_ID")
            .PID = rs.fields("CATALP_P_ID")
            .CATALID = rs.fields("CATALP_CATAL_ID")
            .Price = FNCURR(rs.fields("CATALP_Price"))
            .Serial = FNN(rs.fields("CATAL_Serial"))
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
    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CATALP_P.Fetch(PIID)", PID, , rs
End Function



Public Function Save(ByVal buffer As String, ByVal PID As String) As String
    On Error GoTo errHandler
  Dim rs As ADODB.Recordset
  Dim strSQL As String
  Dim udtProps As CATALPProps
  Dim udtData As CATALPData
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
      strSQL = "SELECT * FROM tCATALP WHERE CATALP_ID= " & udtProps.ID
      rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
      If udtProps.IsNew Then rs.AddNew
      
      With udtProps
        rs.fields("CATALP_P_ID") = PID
        rs.fields("CATALP_CATAL_ID") = .CATALID
        rs.fields("CATALP_Price") = .Price
        rs.Update
        
        If .IsNew Then
          rs.Bookmark = rs.Bookmark
          .ID = rs("CATALP_ID")
        End If
        .IsNew = False
        .IsDirty = False
      End With
      LSet udtData = udtProps
      lngCount = lngCount + 1
      objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
      rs.Close
    Else
      DeleteCATALP udtProps.ID
    End If
  Next
  objPBOut.WriteProperty "Count", lngCount
  
  Set objPB = Nothing
  Set rs = Nothing
  
  Save = objPBOut.Contents
  Set objPBOut = Nothing
  Exit Function
  

    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    Set objPBOut = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CATALP_P.Save(buffer,PID)", Array(buffer, PID), , rs
End Function

Public Sub DeleteObject(PID As String)
    On Error GoTo errHandler
  
    oPC.COShort.execute "DELETE FROM tCATALP WHERE CATALP_P_ID='" & PID & "'"
    Exit Sub
    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CATALP_P.DeleteObject(PID)", PID
End Sub
Private Sub DeleteCATALP(ID As Long)
    On Error GoTo errHandler
    oPC.COShort.execute "DELETE FROM tCATALP WHERE CATALP_ID = " & ID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CATALP_P.DeleteCATALP(ID)", ID
End Sub




