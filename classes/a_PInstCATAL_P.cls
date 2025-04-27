VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_PInstCATAL_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Fetch(ByVal PIID As Long) As String
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As PICATALProps
Dim udtData As PICATALData
Dim objPB As PropertyBag
Dim lngCount As Long

    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tCATALPI JOIN tCatalogue on CATALPI_CATAL_ID = CATAL_ID WHERE CATALPI_PI_ID= " & PIID
    rs.Open strSQL, oPC.CO, adOpenForwardOnly, adLockReadOnly
    Set objPB = New PropertyBag
    On Error GoTo ERRH
    Do While Not rs.EOF
        With udtProps
            .ID = rs.Fields("CATALPI_ID")
            If HasData(rs.Fields("CATALPI_PI_ID")) Then .PIID = rs.Fields("CATALPI_PI_ID")
            If HasData(rs.Fields("CATALPI_CATAL_ID")) Then .CATALID = rs.Fields("CATALPI_CATAL_ID")
            If HasData(rs.Fields("CATALPI_Price")) Then .Price = rs.Fields("CATALPI_Price")
            If HasData(rs.Fields("CATAL_Serial")) Then .Serial = rs.Fields("CATAL_Serial")
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
        End With
        LSet udtData = udtProps
        lngCount = lngCount + 1
        objPB.WriteProperty "Item" & CStr(lngCount), udtData.Buffer
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
  
ERRH:
  rs.Close
  Set rs = Nothing
  Err.Raise Err.Number
End Function



Public Function Save(ByVal Buffer As String, ByVal PIID As Long) As String
  Dim rs As ADODB.Recordset
  Dim strSQL As String
  Dim udtProps As PICATALProps
  Dim udtData As PICATALData
  Dim objPB As PropertyBag
  Dim objPBOut As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngCount As Long
  
  Set objPB = New PropertyBag
  Set objPBOut = New PropertyBag
  arBuffer = Buffer
  objPB.Contents = arBuffer
  Set rs = New ADODB.Recordset
  
  On Error GoTo ERRH
  For lngIndex = 1 To objPB.ReadProperty("Count")
    udtData.Buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
    LSet udtProps = udtData
    
    If Not udtProps.IsDeleted Then
      strSQL = "SELECT * FROM tCATALPI WHERE CATALPI_ID= " & udtProps.ID
      rs.Open strSQL, oPC.CO, adOpenKeyset, adLockOptimistic
      If udtProps.IsNew Then rs.AddNew
      
      With udtProps
        rs.Fields("CATALPI_PI_ID") = PIID
        rs.Fields("CATALPI_CATAL_ID") = .CATALID
        rs.Fields("CATALPI_Price") = .Price
        rs.Update
        
        If .IsNew Then
          rs.Bookmark = rs.Bookmark
          .ID = rs("CATALPI_ID")
        End If
        .IsNew = False
        .IsDirty = False
      End With
      LSet udtData = udtProps
      lngCount = lngCount + 1
      objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.Buffer
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
  Exit Function
  
ERRH:
  Set objPB = Nothing
  Set objPBOut = Nothing
  Set rs = Nothing
  Err.Raise Err.Number
End Function

Private Sub DeleteCopy(PIID As Long)
Dim strSQL As String
Dim oUtil As z_Batch
    strSQL = "DELETE FROM tCATALPI WHERE CATALPI_ID=" & PIID
    Set oUtil = New z_Batch
    oUtil.RunSQL strSQL
    Set oUtil = Nothing
End Sub




