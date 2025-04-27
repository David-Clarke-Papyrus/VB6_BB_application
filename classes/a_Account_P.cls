VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Account_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As CatHeadProps
Dim udtData As CatHeadData
Dim oPB As PropertyBag
Dim oPBOUT As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
Dim OpenResult As Integer

    Set oPB = New PropertyBag
    Set oPBOUT = New PropertyBag
    arBuffer = buffer
    oPB.Contents = arBuffer
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.BeginTrans
    For lngIndex = 1 To oPB.ReadProperty("Count")
        udtData.buffer = oPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData

        If Not udtProps.IsDeleted Then
            strSQL = "SELECT * FROM tCatalHeading WHERE CATHEAD_ID = " & udtProps.ID
            rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then rs.AddNew
            With rs
                .fields("CATHEAD_Description") = FNS(udtProps.Description)
                .fields("CATHEAd_SOrttag") = FNS(udtProps.SortTag)
                .fields("CATHEAD_Parent") = FNN(udtProps.Parent)
                .Update
                udtProps.IsNew = False
                udtProps.IsDirty = False
            End With
            LSet udtData = udtProps
            lngCount = lngCount + 1
            oPBOUT.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.Close
        Else
          Delete udtProps.ID
    End If
    Next
    oPBOUT.WriteProperty "Count", lngCount
    
    Set oPB = Nothing
    Set rs = Nothing
    
    Save = oPBOUT.Contents
    Set oPBOUT = Nothing
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    ErrPreserve
    If Err = -2147217873 Then
        MsgBox "Duplicate value for age next"
        Save = ""
        oPC.COShort.RollbackTrans
        Set oPB = Nothing
        Set oPBOUT = Nothing
        Set rs = Nothing
        Exit Function
    End If
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_cathead_P.Save(buffer)", buffer, EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Function
Public Function Fetch() As String
    On Error GoTo errHandler
Dim fs As FileSystemObject
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim mudtProps As CatHeadProps
Dim mudtData As CatHeadData
Dim oPB As PropertyBag
Dim lngCount As Long

    Set rs = New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open "SELECT M.CATHEAD_ID,M.CATHEAD_DESCRIPTION,M.CATHEAd_sORTTAG,M.CATHEAD_PARENT,P.CATHEAD_DESCRIPTION as ParentHeading FROM tCATALHeading as M LEFT OUTER JOIN tCATALHEADING as P on M.CATHEAD_Parent =  P.CATHEAD_ID ", oPC.COShort, adOpenDynamic, adLockReadOnly
    Set oPB = New PropertyBag
    Do While Not rs.eof
      With rs
            mudtProps.ID = .fields("CATHEAD_ID")
            mudtProps.Description = FNS(.fields("CATHEAD_Description"))
            mudtProps.SortTag = FNS(.fields("CATHEAd_sOrtTag"))
            mudtProps.Parent = FNN(.fields("CATHEAD_Parent"))
            mudtProps.ParentDescription = FNS(.fields("ParentHeading"))
            mudtProps.IsDirty = False
            mudtProps.IsNew = False
            mudtProps.IsDeleted = False
      End With
      LSet mudtData = mudtProps
      lngCount = lngCount + 1
      oPB.WriteProperty "Item" & CStr(lngCount), mudtData.buffer
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    With oPB
      .WriteProperty "Count", lngCount
      Fetch = .Contents
    End With
    Set oPB = Nothing
    Exit Function


    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_cathead_P.Fetch"
End Function
Private Sub Delete(CATHEADID As Long)
    On Error GoTo errHandler
Dim strSQL As String
Dim oUtil As z_SQL
Dim OpenResult As Integer
    strSQL = "DELETE FROM tCATALHEADING WHERE CATHEAD_ID=" & CATHEADID
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute strSQL
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_cathead_P.Delete(CATHEADID)", CATHEADID
End Sub

Public Sub DeleteAll()
    On Error GoTo errHandler
Dim strSQL As String
Dim oUtil As z_SQL
Dim OpenResult As Integer
  
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "DELETE FROM tCATALHEADING"
    oPC.COShort.execute strSQL

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_cathead_P.DeleteAll"
End Sub









