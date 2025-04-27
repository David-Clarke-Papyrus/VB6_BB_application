VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Copy_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByVal PID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As CopyProps
Dim udtData As CopyData
Dim objPB As PropertyBag
Dim objCopyPB As PropertyBag
Dim lngCount As Long
Dim objPersist As a_CATALPI_P

    Set rs = New ADODB.Recordset
    strSQL = "SELECT tPI.*,TP_Name FROM tPI LEFT JOIN tTP ON PI_SOLDTO_ID = TP_ID WHERE PI_P_ID= '" & PID & "' ORDER BY PI_PURCHASEDATE"
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.eof
        With udtProps
            .ID = rs.fields("PI_ID")
            .PID = rs.fields("PI_P_ID")
            .Serial = FNS(rs.fields("PI_Serial"))
            .PurchaseDate = FND(rs.fields("PI_PurchaseDate"))
            .SoldDate = FND(rs.fields("PI_SoldDate"))
            .Comment = FNS(rs.fields("PI_Comment"))
            .Description = FNS(rs.fields("PI_Description"))
            .Cost = FNCURR(rs.fields("PI_Cost"))
            .Price = FNCURR(rs.fields("PI_Price"))
            .SoldTo = FNS(rs.fields("TP_Name"))
            .FlagText = FNS(rs.fields("PI_FlagText"))
    
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
        End With
        LSet udtData = udtProps
        
        
        Set objCopyPB = New PropertyBag
        With objCopyPB
            .WriteProperty "State", udtData.buffer
            Set objPersist = New a_CATALPI_P
            .WriteProperty "Catalogues", objPersist.Fetch(udtProps.ID)
            Set objPersist = Nothing
            lngCount = lngCount + 1
            objPB.WriteProperty "Item" & CStr(lngCount), .Contents
        End With
        
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
    ErrPreserve
    Set objPB = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy_P.Fetch(PID)", PID
End Function



Public Function Save(ByVal buffer As String, ByVal PID As String, pLastSerialUsed As Integer, pQtyOnHand As Integer) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As CopyProps
Dim udtData As CopyData
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim objCopyPB As PropertyBag
Dim objCopyPBOut As PropertyBag
Dim lngIndex As Long
Dim lngCount As Long
Dim objPersist As a_CATALPI_P
Dim iQtyOnHand As Integer

  Set objPB = New PropertyBag
  Set objPBOut = New PropertyBag
  arBuffer = buffer
  objPB.Contents = arBuffer
  Set rs = New ADODB.Recordset
  
  For lngIndex = 1 To objPB.ReadProperty("Count")
    
    Set objCopyPB = New PropertyBag
    With objCopyPB
      arBuffer = objPB.ReadProperty("Item" & CStr(lngIndex))
      .Contents = arBuffer
      udtData.buffer = .ReadProperty("State")
      LSet udtProps = udtData
    End With
    
    iQtyOnHand = 0
    If Not udtProps.IsDeleted Then
      strSQL = "SELECT * FROM tPI WHERE PI_ID=" & udtProps.ID
      rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
      If udtProps.IsNew Then rs.AddNew
      
      With udtProps
        rs.fields("PI_P_ID") = PID
        rs.fields("PI_Cost") = .Cost
        rs.fields("PI_Price") = .Price
        rs.fields("PI_SoldDate") = .SoldDate
        If .SoldDate < "1990-01-01" Then
            iQtyOnHand = iQtyOnHand + 1
        End If
        rs.fields("PI_PurchaseDate") = .PurchaseDate
        If .Serial = 0 Then
            pLastSerialUsed = pLastSerialUsed + 1
            .Serial = pLastSerialUsed
        End If
        rs.fields("PI_Serial") = .Serial
        rs.fields("PI_Comment") = FNS(.Comment)
        rs.fields("PI_Description") = FNS(.Description)
        rs.fields("PI_FlagText") = FNS(.FlagText)
        rs.Update
        
        If .IsNew Then
          rs.Bookmark = rs.Bookmark
          .ID = rs("PI_ID")
        End If
        .IsNew = False
        .IsDirty = False
      End With
      
      Set objCopyPBOut = New PropertyBag
      
        LSet udtData = udtProps
        objCopyPBOut.WriteProperty "State", udtData.buffer
        Set objPersist = New a_CATALPI_P
      objCopyPBOut.WriteProperty "Catalogues", objPersist.Save(objCopyPB.ReadProperty("Catalogues"), udtProps.ID)
      Set objPersist = Nothing
      lngCount = lngCount + 1
      objPBOut.WriteProperty "Item" & CStr(lngCount), objCopyPBOut.Contents
      Set objCopyPB = Nothing
      Set objCopyPBOut = Nothing
      rs.Close
    Else
      DeleteCopy udtProps.ID
    End If
  Next
  objPBOut.WriteProperty "Count", lngCount
  
  Set objPB = Nothing
  Set rs = Nothing
  
  Save = objPBOut.Contents
  Set objPBOut = Nothing
  pQtyOnHand = iQtyOnHand
  Exit Function
  
'H:
'  MsgBox "In a_Copy_P: Save. " & or
'  Exit Function
'  Resume
    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    Set objPBOut = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy_P.Save(buffer,PID,pLastSerialUsed,pQtyOnHand)", Array(buffer, PID, _
         pLastSerialUsed, pQtyOnHand)
End Function

Private Sub DeleteCopy(PIID As Long)
    On Error GoTo errHandler
Dim strSQL As String
    strSQL = "DELETE FROM tPI WHERE PI_ID=" & PIID
    oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy_P.DeleteCopy(PIID)", PIID
End Sub

Public Sub DeleteObject(ByVal PID As String)
    On Error GoTo errHandler
Dim rsCopy As ADODB.Recordset
Dim objPersist As a_CATALPI_P
Dim strSQL As String
  
    Set objPersist = New a_CATALPI_P
    strSQL = "SELECT PI_ID FROM tPI WHERE PI_P_ID ='" & PID & "'"
    Set rsCopy = New ADODB.Recordset
    rsCopy.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Do While Not rsCopy.eof
        objPersist.DeleteObject rsCopy("PI_ID")
        rsCopy.MoveNext
    Loop
    rsCopy.Close
    Set rsCopy = Nothing
    Set objPersist = Nothing
  
    oPC.COShort.execute "DELETE FROM tPI WHERE PI_P_ID='" & PID & "'"
  Exit Sub


errHandler:
    ErrPreserve
    RlsObjs rsCopy
    Set objPersist = Nothing
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Copy_P.DeleteObject(PID)", PID
End Sub


