VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Nielsen_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Save(ByVal buffer As String, pStatus As String) As String
    On Error GoTo errHandler

Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As ConfigProps
Dim udtData As ConfigData
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
Dim strStatus As String

  Set objPB = New PropertyBag
  arBuffer = buffer
  With objPB
    .Contents = arBuffer
    udtData.buffer = .ReadProperty("State")
  End With
  LSet udtProps = udtData
  
  strSQL = "SELECT * FROM tNielsen"
  Set rs = New ADODB.Recordset
  rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
  If udtProps.IsNew Then rs.AddNew
  With rs
        .Fields("N_Query") = FNS(udtProps.Q)
        .Fields("N_LCQuery") = FNS(udtProps.LCQ)
        .Fields("N_FTPPassive") = FNB(udtProps.FTPPassive)
        .Fields("N_FTPAddress") = FNS(udtProps.FTPAddress)
        .Fields("N_Username") = FNS(udtProps.FTPUsername)
        .Fields("N_Password") = FNS(udtProps.FTPPassword)
        .Fields("N_FTPDefaultFolder") = FNS(udtProps.FTPDefaultFolder)
        .Fields("N_LASTDateSalesSent") = FND(udtProps.LastDateSalesSent)
        .Fields("N_CentralFTPPassive") = FNB(udtProps.CentralFTPPassive)
        .Fields("N_CentralFTPAddress") = FNS(udtProps.CentralFTPAddress)
        .Fields("N_CentralUsername") = FNS(udtProps.CentralFTPUsername)
        .Fields("N_CentralPassword") = FNS(udtProps.CentralFTPPassword)
        .Fields("N_CentralFTPDefaultFolder") = FNS(udtProps.CentralFTPDefaultFolder)
        .Fields("N_LASTDateLCExtraction") = FND(udtProps.LastDateLCSent)
        .Fields("N_DUN") = FNS(udtProps.DUN)
        .Fields("N_NielsenActive") = FNB(udtProps.NielsenActive)
        .Fields("N_LoyaltySchemeActive") = FNB(udtProps.LoyaltySchemeActive)
        .Fields("N_StockSharingActive") = FNB(udtProps.StockSharingACtive)
        .Fields("N_AuditingActive") = FNB(udtProps.AuditingActive)
        .Update
        udtProps.IsNew = False
        udtProps.IsDirty = False
    End With
    rs.Close
  Set rs = Nothing
  
  Set objPBOut = New PropertyBag
  
        LSet udtData = udtProps
        objPBOut.WriteProperty "State", udtData.buffer
        

        Set objPB = Nothing
        Save = objPBOut.Contents
        Set objPBOut = Nothing
  Exit Function
        
'ERRH:
'MsgBox Error
'Dim strNewCode
'Resume
'Exit Function
'Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Nielsen_P.Save(Buffer,pStatus)", Array(buffer, pStatus), EA_NORERAISE
    handleError
End Function
Public Function Fetch() As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As ConfigProps
Dim udtData As ConfigData
Dim objPB As PropertyBag
Dim lngCount As Long

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tNielsen", oPC.COShort, adOpenStatic, adLockReadOnly
    Set objPB = New PropertyBag
    If Not rs.EOF Then
        With rs
        udtProps.Q = FNS(.Fields("N_Query"))
        udtProps.FTPPassive = FNB(.Fields("N_FTPPassive"))
        udtProps.FTPAddress = FNS(.Fields("n_FTPAddress"))
        udtProps.FTPUsername = FNS(.Fields("N_Username"))
        udtProps.FTPPassword = FNS(.Fields("N_Password"))
        udtProps.FTPDefaultFolder = FNS(.Fields("N_FTPDefaultFolder"))
        udtProps.LastDateSalesSent = FND(.Fields("N_LASTDateSalesSent"))
        udtProps.LCQ = FNS(.Fields("N_LCQuery"))
        udtProps.CentralFTPPassive = FNB(.Fields("N_CentralFTPPassive"))
        udtProps.CentralFTPAddress = FNS(.Fields("n_CentralFTPAddress"))
        udtProps.CentralFTPUsername = FNS(.Fields("N_CentralUsername"))
        udtProps.CentralFTPPassword = FNS(.Fields("N_CentralPassword"))
        udtProps.CentralFTPDefaultFolder = FNS(.Fields("N_CentralFTPDefaultFolder"))
        udtProps.LastDateLCSent = FND(.Fields("N_LASTDateLCExtraction"))
        udtProps.DUN = FNS(.Fields("N_DUN"))
        udtProps.NielsenActive = FNB(.Fields("N_NielsenActive"))
        udtProps.LoyaltySchemeActive = FNB(.Fields("N_LoyaltySchemeActive"))
        udtProps.StockSharingACtive = FNB(.Fields("N_StockSharingActive"))
        udtProps.DateLastLCEditedReceived = FND(.Fields("N_DateLastLCEditedReceived"))
        udtProps.AuditingActive = FNB(.Fields("N_AuditingActive"))
        LSet udtData = udtProps
        End With
        rs.Close
        Set rs = Nothing
        
        With objPB
          .WriteProperty "State", udtData.buffer
          Fetch = .Contents
        End With
        
        Set objPB = Nothing

'    Else
'        rs.MoveNext
  End If
  Exit Function
        
'ERRH:
'  rs.Close
'  Set rs = Nothing
'  Err.Raise Err.Number
'  Exit Function
'  Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Nielsen_P.Fetch", , EA_NORERAISE
    handleError
End Function

Private Sub DeleteConfiguration()
    On Error GoTo errHandler
Dim strSQL As String

  strSQL = "DELETE FROM tblConfiguration WHERE ID=1"
  oPC.COShort.Execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Nielsen_P.DeleteConfiguration"
End Sub

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler
'  Dim strSQL As String
'  Dim objPersist As ProjectPersist
'
'  On Error GoTo ERRH
'    strSQL = "DELETE FROM tblCOnfiguration WHERE ID=1"
'    oPC.COShort.Execute strSQL
'
'  If mobjContext Is Nothing Then
'    Set objPersist = New a_Company_Persist
'  Else
''    Set objPersist = mobjContext.CreateInstance("TaskServer.a_Company_Persist")
'  End If
'  objPersist.DeleteObject ID
'  Set objPersist = Nothing
'  If Not mobjContext Is Nothing Then mobjContext.SetComplete
'  Exit Sub
'
'ERRH:
'  Set cnClient = Nothing
'  If Not mobjContext Is Nothing Then mobjContext.SetAbort
'  Err.Raise Err.Number
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Nielsen_P.DeleteObject(ID)", ID, EA_NORERAISE
    handleError
End Sub

