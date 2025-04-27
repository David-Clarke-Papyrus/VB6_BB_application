VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Staff_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Save(ByVal buffer As String, pStatus As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As StaffProps
Dim udtData As StaffData
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long

    Set objPB = New PropertyBag
    Set objPBOut = New PropertyBag
    arBuffer = buffer
    objPB.Contents = arBuffer
Dim OpenResult As Integer
    
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData

        If Not udtProps.IsDeleted Then
            strSQL = "SELECT * FROM tStaffmember WHERE SM_ID = " & udtProps.ID
            rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then rs.AddNew
            If Not rs.eof Then
                With rs
                    .fields("SM_ROLE") = (udtProps.Role)
                    .fields("SM_Name") = FNS(udtProps.StaffName)
                    .fields("SM_ShortName") = FNS(udtProps.Shortname)
                    .fields("SM_Telephone") = FNS(udtProps.StaffTel)
                    .fields("SM_Mobile") = FNS(udtProps.StaffCell)
                    .fields("SM_Password") = FNS(udtProps.Password)
                    .fields("SM_Email") = FNS(udtProps.EMail)
                    .fields("SM_Signature") = FNS(udtProps.Signature)
                '    .fields("SM_Active") = FNB(udtProps.Active)
                    .Update
                    If udtProps.IsNew Then
                      .Bookmark = .Bookmark
                      udtProps.ID = rs("SM_ID")
                    End If
                    udtProps.IsNew = False
                    udtProps.IsDirty = False
                    udtProps.IsDeleted = False
                End With
                LSet udtData = udtProps
                lngCount = lngCount + 1
                objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            End If
            rs.Close
        Else
            DeleteSTAFF udtProps.ID
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
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff_P.Save(buffer,pStatus)", Array(buffer, pStatus)
End Function
Public Function Fetch() As String
    On Error GoTo errHandler
Dim fs As FileSystemObject
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim mudtProps As StaffProps
Dim mudtData As StaffData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim OpenResult As Integer
    
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open "SELECT *  FROM tStaffmember", oPC.COShort, adOpenDynamic, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.eof
      With rs
            mudtProps.ID = .fields("SM_ID")
            mudtProps.StaffName = FNS(.fields("SM_Name"))
            mudtProps.Shortname = FNS(.fields("SM_ShortName"))
            mudtProps.StaffCell = FNS(.fields("SM_Mobile"))
            mudtProps.StaffTel = FNS(.fields("SM_Telephone"))
            mudtProps.Password = FNS(.fields("SM_Password"))
            mudtProps.EMail = FNS(.fields("SM_Email"))
            mudtProps.Signature = FNS(.fields("SM_Signature"))
      '      mudtProps.Active = FNB(.fields("SM_Active"))
            
            If IsNull(.fields("SM_ROLE")) Then
                mudtProps.Role = Space(50)
            Else
                mudtProps.Role = (.fields("SM_ROLE"))
            End If
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
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    With objPB
      .WriteProperty "Count", lngCount
      Fetch = .Contents
    End With
    Set objPB = Nothing
    Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Staff_P.Fetch"
End Function
Private Sub DeleteSTAFF(STID As Long)
Dim cnClient As Connection
Dim strSQL As String
    strSQL = "DELETE FROM tStaff WHERE SM_ID=" & STID
    oPC.COShort.execute strSQL
End Sub







