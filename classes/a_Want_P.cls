VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Want_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByVal PID As String) As String
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As WantProps
Dim udtData As WantData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim strAcno As String
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tCOL JOIN tTR ON COL_TR_ID = TR_ID JOIN tTP on TR_TP_ID = TP_ID WHERE COL_P_ID= '" & PID & "'"
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Set objPB = New PropertyBag
    On Error GoTo H
    Do While Not rs.eof
        With udtProps
            .ID = rs.fields("COL_ID")
            .PID = rs.fields("COL_P_ID")
            .TPID = rs.fields("TR_TP_ID")
            .RequestDate = FND(rs.fields("COL_WANTDATE"))
            If Len(FNS(rs.fields("TP_ACNO"))) > 0 Then
                strAcno = " (" & FNS(rs.fields("TP_ACNO")) & ")"
            Else
                strAcno = ""
            End If
            .CustomerName = FNS(rs.fields("TP_Initials")) & " " & FNS(rs.fields("TP_Name")) & strAcno
            .Notes = FNS(rs.fields("COL_Note"))
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
  
H:
  rs.Close
  Set rs = Nothing
  Err.Raise Err.Number
End Function



Public Function Save(ByVal buffer As String, ByVal PID As String) As String
  Dim rs As ADODB.Recordset
  Dim strSQL As String
  Dim udtProps As WantProps
  Dim udtData As WantData
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
  
  On Error GoTo H
  For lngIndex = 1 To objPB.ReadProperty("Count")
    udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
    LSet udtProps = udtData
    
    If Not udtProps.IsDeleted Then
      strSQL = "SELECT * FROM tWants WHERE WT_ID= '" & FNN(udtProps.ID) & "'"
      rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
      If udtProps.IsNew Then rs.AddNew
      
      With udtProps
        rs.fields("WT_P_ID") = Trim$(PID)
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
  Exit Function
  
H:
MsgBox "a_Want_P:Save: " & Error
Exit Function
Resume
End Function

Public Sub DeleteObject(ByVal PID As String)
    oPC.COShort.execute "DELETE FROM tWants WHERE WT_P_ID='" & PID & "'"
End Sub

Private Sub DeleteWant(ID As Long)
    oPC.COShort.execute "DELETE FROM tWants WHERE WT_ID = " & ID
End Sub




