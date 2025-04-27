VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Company_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As CompanyProps
Dim udtData As CompanyData
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
            strSQL = "SELECT * FROM tCompany WHERE COMP_ID = " & udtProps.ID
            rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then rs.AddNew
            With rs
                .fields("COMP_Name") = FNS(udtProps.CompanyName)
                .fields("COMP_Code") = FNS(udtProps.CompanyCode)
                .fields("COMP_VATNumber") = FNS(udtProps.VatNumber)
                .fields("COMP_LogoFileName") = FNS(udtProps.LogoFilePath)
                .fields("COMP_StreetAdd") = FNS(udtProps.StreetAddress)
                .fields("COMP_PostalAdd") = FNS(udtProps.PostalAddress)
                .fields("COMP_RegistrationNumber") = FNS(udtProps.CoRegistrationNumber)
                .fields("COMP_Pastel") = FNS(udtProps.Pastel)
                .fields("COMP_BankDetails") = FNS(udtProps.BankDetails)
                .Update
                If udtProps.IsNew Then
                    .Bookmark = .Bookmark
                    udtProps.ID = .fields("COMP_ID")
                End If
                udtProps.IsNew = False
                udtProps.IsDirty = False
            End With
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.Close
        Else
          DeleteCompany udtProps.ID
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
    ErrorIn "a_Company_P.Save(buffer)", buffer
End Function
Public Function Fetch() As String
    On Error GoTo errHandler
Dim fs As FileSystemObject
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim mudtProps As CompanyProps
Dim mudtData As CompanyData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open "tCompany", oPC.COShort, adOpenDynamic, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.eof
      With rs
          mudtProps.ID = .fields("COMP_ID")
          mudtProps.CompanyCode = FNS(.fields("COMP_Code"))
          mudtProps.CompanyName = FNS(.fields("COMP_Name"))
          mudtProps.VatNumber = FNS(.fields("COMP_VATNumber"))
          mudtProps.LogoFilePath = FNS(.fields("COMP_LogoFileName"))
          mudtProps.StreetAddress = FNS(.fields("COMP_StreetAdd"))
          mudtProps.PostalAddress = FNS(.fields("COMP_PostalAdd"))
          mudtProps.CoRegistrationNumber = FNS(.fields("COMP_RegistrationNumber"))
          mudtProps.Pastel = FNS(.fields("COMP_Pastel"))
          mudtProps.BankDetails = FNS(.fields("COMP_BankDetails"))
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
errHandler:
    ErrPreserve
    Set objPB = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company_P.Fetch"
End Function
Private Sub DeleteCompany(COMPID As Long)
    On Error GoTo errHandler
Dim strSQL As String
  strSQL = "DELETE FROM tCompany WHERE Comp_ID=" & CStr(COMPID)
  oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company_P.DeleteCompany(COMPID)", COMPID
End Sub

Public Sub DeleteObject()
    On Error GoTo errHandler
  Dim cnClient As Connection
  Dim strSQL As String
  
  strSQL = "DELETE FROM tblCompany"
  oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Company_P.DeleteObject"
End Sub


