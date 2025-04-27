VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Address_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function Fetch(ByVal TPID As Long) As String
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As AddressProps
Dim udtData As AddressData
Dim objPB As PropertyBag
Dim lngCount As Long
'Dim oProdcode As z_ProdCode

    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM vAddresses WHERE ADD_TP_ID=" & TPID & " ORDER BY ADD_ADDRESSEE"
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Set objPB = New PropertyBag
    On Error GoTo H
    Do While Not rs.eof
        With udtProps
            .ID = rs.fields("ADD_ID")
            .TPID = Trim$(rs.fields("ADD_TP_ID"))
            .Description = FNS(rs.fields("Add_Description"))
            .Phone = (FNS(rs.fields("ADD_Phone")))
            .Fax = (FNS(rs.fields("ADD_Fax")))
            .EMail = FNS(rs.fields("Add_Email"))
            .Addressee = FNS(rs.fields("ADD_Addressee"))
            .BusPhone = (FNS(rs.fields("ADD_BusPhone")))
            .PostageType = FNN(rs.fields("ADD_PostageType"))
            .GetsCatalogue = FNB(rs.fields("ADD_GetsCatalogue"))
            .ForMailing = FNB(rs.fields("ADD_ForMailing"))
            .Appro = FNB(rs.fields("ADD_Default"))
            .BillTo = FNB(rs.fields("ADD_BillTO"))
            .DelTo = FNB(rs.fields("ADD_DelTo"))
            .OrderTo = FNB(rs.fields("ADD_OrderTo"))
            .Line1 = FNS(rs.fields("ADD_L1"))
            .Line2 = FNS(rs.fields("ADD_L2"))
            .Line3 = FNS(rs.fields("ADD_L3"))
            .Line4 = FNS(rs.fields("ADD_L4"))
            .Line5 = FNS(rs.fields("ADD_L5"))
            .Line6 = FNS(rs.fields("ADD_L6"))
            .CountryID = FNN(rs.fields("ADD_Cntry_ID"))
            .Category = FNN(rs.fields("ADD_Category"))
            .pCode = FNS(rs.fields("ADD_PCode"))
            .Delivery = FNS(rs.fields("ADD_Delivery"))
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
          .WriteProperty "Count", lngCount
          Fetch = .Contents
    End With
    Set objPB = Nothing
    Exit Function
  
H:
  rs.Close
  Set rs = Nothing
'  If Not mobjContext Is Nothing Then mobjContext.SetAbort
  Err.Raise Err.Number
End Function



Public Function Save(ByVal buffer As String, ByVal CustID As Long) As String
  Dim rs As ADODB.Recordset
  Dim strSQL As String
  Dim udtProps As AddressProps
  Dim udtData As AddressData
  Dim objPB As PropertyBag
  Dim objPBOut As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngCount As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
  Set objPB = New PropertyBag
  Set objPBOut = New PropertyBag
  arBuffer = buffer
  objPB.Contents = arBuffer
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient
  On Error GoTo H
  For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData
    
        If Not udtProps.IsDeleted Then
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = oPC.COShort
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "PersistAddress"
            Set par = cmd.CreateParameter("@NU", adChar, adParamInput, 1)
            cmd.Parameters.Append par
            If udtProps.IsNew Then
                par.Value = "I"
            Else
                par.Value = "U"
            End If
            Set par = cmd.CreateParameter("@ADDID", adInteger, adParamInput, , FNN(udtProps.ID))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_TP_ID", adInteger, adParamInput, , CustID)
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_L1", adVarChar, adParamInput, 50, FNS(udtProps.Line1))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_L2", adVarChar, adParamInput, 50, FNS(udtProps.Line2))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_L3", adVarChar, adParamInput, 50, FNS(udtProps.Line3))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_L4", adVarChar, adParamInput, 50, FNS(udtProps.Line4))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@Add_L5", adVarChar, adParamInput, 50, FNS(udtProps.Line5))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_L6", adVarChar, adParamInput, 50, FNS(udtProps.Line6))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_PCode", adVarChar, adParamInput, 10, FNS(udtProps.pCode))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_Description", adVarChar, adParamInput, 15, FNS(udtProps.Description))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_Phone", adVarChar, adParamInput, 50, FNS(udtProps.Phone))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_BusPhone", adVarChar, adParamInput, 50, FNS(udtProps.BusPhone))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_Fax", adVarChar, adParamInput, 50, FNS(udtProps.Fax))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_Email", adVarChar, adParamInput, 100, FNS(udtProps.EMail))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_Cntry_ID", adInteger, adParamInput, , FNN(udtProps.CountryID))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_Addressee", adVarChar, adParamInput, 80, FNS(udtProps.Addressee))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_GetsCatalogue", adBoolean, adParamInput, , FNB(udtProps.GetsCatalogue))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_ForMailing", adBoolean, adParamInput, , FNB(udtProps.ForMailing))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_PostageType", adInteger, adParamInput, , FNN(udtProps.PostageType))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_Category", adInteger, adParamInput, , FNN(udtProps.Category))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_Default", adBoolean, adParamInput, , FNB(udtProps.Appro))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_BillTo", adBoolean, adParamInput, , FNB(udtProps.BillTo))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_DelTo", adBoolean, adParamInput, , FNB(udtProps.DelTo))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ADD_OrderTo", adBoolean, adParamInput, , FNB(udtProps.OrderTo))
            cmd.Parameters.Append par
            Set par = cmd.CreateParameter("@ID", adInteger, adParamOutput)
            cmd.Parameters.Append par
            
            cmd.execute
        
'      strSQL = "SELECT Add_ID,ADD_TP_ID,ADD_L1,ADD_L2,ADD_L3,ADD_L4,Add_L5,ADD_L6,ADD_PCode,ADD_Addressee,ADD_Cntry_ID,ADD_Description, " _
'            & " ADD_Phone,ADD_Fax,ADD_Email,ADD_BusPhone,ADD_PostageType,ADD_GetsCatalogue,ADD_ForMailing,ADD_Default, " _
'            & " ADD_BillTo,ADD_DelTo,ADD_OrderTo,ADD_Category,ADD_PCode FROM tADD WHERE Add_ID=" & udtProps.ID
'      rs.Open strSQL, oPC.COShort, adOpenDynamic, adLockOptimistic
'      If udtProps.IsNew Then rs.AddNew
'
'      With udtProps
'        rs!ADD_TP_ID = CustID
'        rs!ADD_L1 = FNS(udtProps.Line1)
'        rs!ADD_L2 = FNS(udtProps.Line2)
'        rs!ADD_L3 = FNS(udtProps.Line3)
'        rs!ADD_L4 = FNS(udtProps.Line4)
'        rs!Add_L5 = FNS(udtProps.Line5)
'        rs!ADD_L6 = FNS(udtProps.Line6)
'        rs!ADD_PCode = FNS(udtProps.pCODE)
'        rs!ADD_Addressee = FNS(udtProps.Addressee)
'        rs!ADD_Cntry_ID = FNN(udtProps.CountryID)
'        rs!ADD_Description = FNS(udtProps.Description)
'        rs!ADD_Phone = FNS(udtProps.Phone)
'        rs!ADD_Fax = FNS(udtProps.Fax)
'        rs!ADD_Email = FNS(udtProps.Email)
'        rs!ADD_BusPhone = FNS(udtProps.BusPhone)
'        rs!ADD_PostageType = FNN(udtProps.PostageType)
'        rs!ADD_GetsCatalogue = FNB(udtProps.GetsCatalogue)
'        rs!ADD_ForMailing = FNB(udtProps.ForMailing)
'        rs!ADD_Default = FNB(udtProps.Appro)
'        rs!ADD_BillTo = FNB(udtProps.BillTo)
'        rs!ADD_DelTo = FNB(udtProps.DelTo)
'        rs!ADD_OrderTo = FNB(udtProps.OrderTo)
'        rs!ADD_Category = FNN(udtProps.Category)
'        rs!ADD_PCode = FNS(udtProps.pCODE)
'        rs.Update
            If udtProps.IsNew Then udtProps.ID = cmd.Parameters("@ADDID")
            udtProps.IsNew = False
            udtProps.IsDirty = False
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
        Else
            DeleteAddress udtProps.ID
        End If
'
'        If .IsNew Then
'          rs.MoveLast
'          .ID = rs("Add_ID")
'        End If
'        udtProps.IsNew = False
'        udtProps.IsDirty = False
'      End With
'      LSet udtData = udtProps
'      lngCount = lngCount + 1
'      objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
'      rs.Close
'    Else
'      DeleteAddress udtProps.ID
'    End If
  Next
  objPBOut.WriteProperty "Count", lngCount
  
  Set objPB = Nothing
  Set rs = Nothing
  
  Save = objPBOut.Contents
  Set objPBOut = Nothing
'  If Not mobjContext Is Nothing Then mobjContext.SetComplete
  Exit Function
  
H:
    If Err = -2147217873 Then
        Save = ""
        Set rs = Nothing
        Exit Function
    Else
      Set objPB = Nothing
      Set objPBOut = Nothing
      Set rs = Nothing
    '  If Not mobjContext Is Nothing Then mobjContext.SetAbort
      Err.Raise Err.Number
    End If
  Resume
End Function

Private Sub DeleteAddress(PID As Long)
Dim cnClient As Connection
Dim strSQL As String
Dim oUtil As z_SQL
    strSQL = "DELETE FROM tADD WHERE ADD_ID=" & PID
    Set oUtil = New z_SQL
    oUtil.RunSQL strSQL
    Set oUtil = Nothing
End Sub

Public Sub DeleteObject(ByVal TPID As Long)
  
  oPC.COShort.execute "DELETE FROM tADD WHERE ADD_TP_ID=" & TPID
  Exit Sub

Errh:
  Err.Raise Err.Number
End Sub




