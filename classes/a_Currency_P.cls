VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Currency_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As CurrencyProps
Dim udtData As CurrencyData
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
      strSQL = "SELECT * FROM tCurrency WHERE CURR_ID = " & udtProps.ID
      rs.open strSQL, oPC.COShort, adOpenDynamic, adLockOptimistic
      If udtProps.IsNew Then rs.AddNew
    '  If udtProps.IsNew Then rs.AddNew
        
      With rs
        .fields("CURR_Description") = FNS(udtProps.Description)
        .fields("CURR_Symbol") = FNS(udtProps.Symbol)
        .fields("CURR_Format") = FNS(udtProps.Format)
        .fields("CURR_Divisor") = FNN(udtProps.Divisor)
        .fields("CURR_ConvertToLocal") = FNDBL(udtProps.Factor)
        .fields("CURR_SYSTEM") = FNS(udtProps.SYS)
        .Update
        If udtProps.IsNew Then
            rs.MoveLast
            udtProps.ID = .fields("CURR_ID")
        End If
        udtProps.IsNew = False
        udtProps.IsDirty = False
      End With
      LSet udtData = udtProps
      lngCount = lngCount + 1
      objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
      rs.Close
    Else
      DeleteCurrency udtProps.ID
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
    ErrorIn "a_Currency_P.Save(buffer)", buffer, EA_DFTRBKCLS, oPC.COShort
End Function
Public Function Fetch() As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As CurrencyProps
Dim udtData As CurrencyData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open "SELECT * FROM tCurrency", oPC.COShort, adOpenStatic, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.eof
    With rs
        udtProps.Description = FNS(.fields("CURR_Description"))
        udtProps.Symbol = FNS(.fields("CURR_Symbol"))
        udtProps.Format = FNS(.fields("CURR_Format"))
        udtProps.Divisor = FNN(.fields("CURR_Divisor"))
        udtProps.Factor = FNDBL(.fields("CURR_ConvertToLocal"))
        udtProps.ID = FNN(.fields("CURR_ID"))
        udtProps.SYS = FNS(.fields("CURR_SYSTEM"))
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

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency_P.Fetch"
End Function

Private Sub DeleteCurrency(CurrencyID As Long)
    On Error GoTo errHandler
Dim strSQL As String

    strSQL = "DELETE FROM tCurrency WHERE CURR_ID=" & CStr(CurrencyID)
    oPC.COShort.execute strSQL
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Currency_P.DeleteCurrency(CurrencyID)", CurrencyID
End Sub

