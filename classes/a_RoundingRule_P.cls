VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_RoundingRule_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Save(ByVal buffer As String) As String
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As RRProps
Dim udtData As RRData
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
            strSQL = "SELECT * FROM tRoundingRule"
            rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then rs.AddNew
            With rs
                .fields("RR_LowerBound") = FNN(udtProps.LowerBound)
                .fields("RR_UpperBound") = FNN(udtProps.UpperBound)
                .fields("RR_RoundTo") = FNN(udtProps.RoundTo)
                .Update
                udtProps.IsNew = False
                udtProps.IsDirty = False
            End With
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.Close
        Else
          DeleteRoundingRule udtProps.ID
    End If
    Next
    objPBOut.WriteProperty "Count", lngCount
    
    Set objPB = Nothing
    Set rs = Nothing
    
    Save = objPBOut.Contents
    Set objPBOut = Nothing
    Exit Function
H:
  Set objPB = Nothing
  Set objPBOut = Nothing
  Set rs = Nothing
  Err.Raise Err.Number
End Function
Public Function Fetch() As String
Dim fs As FileSystemObject
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim mudtProps As RRProps
Dim mudtData As RRData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set rs = New ADODB.Recordset
    rs.open "SELECT *  FROM tRoundingRule", oPC.COShort, adOpenDynamic, adLockReadOnly
    Set objPB = New PropertyBag
    On Error GoTo H
    Do While Not rs.eof
      With rs
            mudtProps.ID = .fields("RR_ID")
            mudtProps.LowerBound = .fields("RR_LowerBound")
            mudtProps.UpperBound = .fields("RR_UpperBound")
            mudtProps.RoundTo = .fields("RR_RoundTo")
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

H:
  rs.Close
  Set rs = Nothing
  Err.Raise Err.Number
End Function
Private Sub DeleteRoundingRule(ID As Long)
Dim cnClient As Connection
Dim strSQL As String
    strSQL = "DELETE FROM tRoundingRule WHERE RR_ID=" & ID
    oPC.COShort.execute strSQL
End Sub

Public Sub DeleteAllROundingRules()
Dim strSQL As String
  
    strSQL = "DELETE FROM tRoundingRule"
    oPC.COShort.execute strSQL

H:
 ' Set cnClient = Nothing
'  If Not mobjContext Is Nothing Then mobjContext.SetAbort
  Err.Raise Err.Number
End Sub






