VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Catalogue_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Implements ObjectControl

'Private mobjContext As ObjectContext

Public Function Fetch(Optional CATALID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strSQL As String
Dim strWHERE As String
Dim udtProps As CatalogueProps
Dim udtData As CatalogueData
Dim objPB As PropertyBag
Dim OpenResult As Integer
    
  strSQL = "SELECT * FROM tCatalogue WHERE CATAL_ID = " & CATALID

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open strSQL, oPC.COShort, adOpenDynamic, adLockOptimistic
    If Not rs.eof Then
        With rs
          udtProps.ID = !CATAL_ID
          udtProps.DateFirstPrinted = FND(!CATAL_DateFirstPrinted)
          udtProps.Description = FNS(!CATAL_Description)
          udtProps.Serial = FNN(!CATAL_Serial)
          .Close
        End With
        Set rs = Nothing
        LSet udtData = udtProps
        
        Set objPB = New PropertyBag
        With objPB
          .WriteProperty "State", udtData.buffer
          Fetch = .Contents
        End With
        Set objPB = Nothing
  Else
    ' force an or
    rs.MoveNext
  End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    ErrPreserve
    Fetch = ""
    If OpenResult = 0 Then oPC.DisconnectDBShort
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue_P.Fetch(CATALID)", CATALID
End Function

Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim udtProps As CatalogueProps
Dim udtData As CatalogueData
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim OpenResult As Integer
Dim cmd As New ADODB.Command
Dim strSQL As String

    Set objPB = New PropertyBag
    arBuffer = buffer
    With objPB
      .Contents = arBuffer
      udtData.buffer = .ReadProperty("State")
    End With
    LSet udtProps = udtData
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    With udtProps
        strSQL = "SELECT * FROM tCatalogue WHERE CATAL_ID = " & .ID
        rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
        If .IsNew Then rs.AddNew
        rs!CATAL_DateFirstPrinted = .DateFirstPrinted
        rs!CATAL_Description = Trim$(.Description)
        rs!CATAL_Serial = .Serial
        udtProps.dbactionStatus = 0
        rs.Update
        If .dbactionStatus = 0 Then
            If .IsNew Then
              rs.MoveLast
              .ID = rs.fields("CATAL_ID")
            End If
            rs.Close
        End If
        Set rs = Nothing
    End With
        
    Set objPBOut = New PropertyBag
    
    LSet udtData = udtProps
    
    With objPBOut
        .WriteProperty "State", udtData.buffer
    End With
    Set objPB = Nothing
    
    Save = objPBOut.Contents
    Set objPBOut = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
        
EXITH:
  Exit Function

errHandler:
    ErrPreserve
    If OpenResult = 0 Then oPC.DisconnectDBShort
    If Err = -2147217873 Then
        udtProps.dbactionStatus = 22
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue_P.Save(buffer)", buffer
End Function

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler
Dim strSQL As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "DELETE FROM tCatalogue WHERE CATAL_ID=" & ID
    oPC.COShort.execute strSQL
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    ErrPreserve
    If OpenResult = 0 Then oPC.DisconnectDBShort
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Catalogue_P.DeleteObject(ID)", ID
End Sub




