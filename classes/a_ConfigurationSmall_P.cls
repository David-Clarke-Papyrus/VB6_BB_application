VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Configuration_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Save(ByVal Buffer As String, pStatus As String) As String

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
  arBuffer = Buffer
  With objPB
    .Contents = arBuffer
    udtData.Buffer = .ReadProperty("State")
  End With
  LSet udtProps = udtData
  
  strSQL = "SELECT * FROM tConfiguration"
  Set rs = New ADODB.Recordset
  rs.Open strSQL, oPC.CO, adOpenKeyset, adLockOptimistic
  On Error GoTo ERRH
  If udtProps.IsNew Then rs.AddNew
  With rs
        .Fields("CF_Q") = Trim$(udtProps.Q)
        .Fields("CF_QQ") = Trim$(udtProps.QQ)
        .Fields("CF_BelowDIP") = Trim$(udtProps.BelowDIP)
        .Fields("CF_AS400ConnectionString") = Trim$(udtProps.AS400COnnectionString)
        .Fields("CF_GrowerFolder") = Trim$(udtProps.GrowerFolder)
        .Fields("CF_RepFolder") = Trim$(udtProps.RepFolder)
        .Fields("CF_PDFFolder") = Trim$(udtProps.PDFFolder)
        .Fields("CF_FTPPassive") = Trim$(udtProps.FTPPassive)
        .Fields("CF_FTPAddress") = Trim$(udtProps.FTPAddress)
        .Fields("CF_FTPUsername") = Trim$(udtProps.FTPUsername)
        .Fields("CF_FTPPassword") = Trim$(udtProps.FTPPassword)
        .Fields("CF_FTPDefaultFolder") = Trim$(udtProps.FTPDefaultFolder)
        .Update
        udtProps.IsNew = False
        udtProps.IsDirty = False
    End With
    rs.Close
  Set rs = Nothing
  
  Set objPBOut = New PropertyBag
  
        LSet udtData = udtProps
        objPBOut.WriteProperty "State", udtData.Buffer
        

        Set objPB = Nothing
        Save = objPBOut.Contents
        Set objPBOut = Nothing
  Exit Function
        
ERRH:
MsgBox Error
Dim strNewCode
'Resume
Exit Function
Resume
End Function
Public Function Fetch() As String
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As ConfigProps
Dim udtData As ConfigData
Dim objPB As PropertyBag
Dim lngCount As Long

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tConfiguration", oPC.CO, adOpenStatic, adLockReadOnly
    Set objPB = New PropertyBag
    On Error GoTo ERRH
    If Not rs.EOF Then
        With rs
        udtProps.Q = FNS(.Fields("CF_Q"))
        udtProps.QQ = FNS(.Fields("CF_QQ"))
        udtProps.BelowDIP = FNS(.Fields("CF_BelowDIP"))
        udtProps.AS400COnnectionString = FNS(.Fields("CF_AS400ConnectionString"))
        udtProps.RepFolder = FNS(.Fields("CF_RepFolder"))
        udtProps.GrowerFolder = FNS(.Fields("CF_GrowerFolder"))
        udtProps.PDFFolder = FNS(.Fields("CF_PDFFolder"))
        udtProps.FTPPassive = FNB(.Fields("CF_FTPPassive"))
        udtProps.FTPAddress = FNS(.Fields("CF_FTPAddress"))
        udtProps.FTPUsername = FNS(.Fields("CF_FTPUsername"))
        udtProps.FTPPassword = FNS(.Fields("CF_FTPPassword"))
        udtProps.FTPDefaultFolder = FNS(.Fields("CF_FTPDefaultFolder"))
        LSet udtData = udtProps
        End With
        rs.Close
        Set rs = Nothing
        
        With objPB
          .WriteProperty "State", udtData.Buffer
          Fetch = .Contents
        End With
        
        Set objPB = Nothing

    Else
        rs.MoveNext
  End If
  Exit Function
        
ERRH:
  rs.Close
  Set rs = Nothing
  Err.Raise Err.Number
  Exit Function
  Resume
End Function

Private Sub DeleteConfiguration()
Dim strSQL As String

  strSQL = "DELETE FROM tblConfiguration WHERE ID=1"
  oPC.CO.Execute strSQL
End Sub

Public Sub DeleteObject(ByVal ID As Long)
'  Dim strSQL As String
'  Dim objPersist As ProjectPersist
'
'  On Error GoTo ERRH
'    strSQL = "DELETE FROM tblCOnfiguration WHERE ID=1"
'    oPC.CO.Execute strSQL
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
End Sub

