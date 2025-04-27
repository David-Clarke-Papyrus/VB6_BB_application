VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_COL_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByVal ID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As COLProps
Dim udtData As COLData
Dim oPBOUT As PropertyBag
Dim oPBOUT2 As PropertyBag
Dim oPBCOFF As PropertyBag
Dim lngCount As Long
Dim oProdCode As z_ProdCode
Dim objPersist As a_COFF_P
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT * FROM COLine_Fetch WHERE COL_TR_ID=" & ID & " ORDER BY P_TITLE ASC"
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Set oPBOUT = New PropertyBag
    lngCount = 0
    
    Do While Not rs.eof
        With udtProps
            .Sequence = FNN(rs.fields("COL_SEQ"))
            If .Sequence = 0 Then .Sequence = lngCount
            .COLineID = rs.fields("COL_ID")
            .TRID = rs.fields("COL_TR_ID")
            .ProdID = FNS(rs.fields("COL_P_ID"))
            .Qty = FNN(rs.fields("COL_Qty"))
            .QtyFirm = FNN(rs.fields("COL_QtyFirm"))
            .QtySS = FNN(rs.fields("COL_QtySS"))
            .QtyDispatched = FNN(rs.fields("COL_QtyDispatched"))
            .ActionTaken = FNN(rs.fields("COL_ActionTaken"))
            .LastAction = FNS(rs.fields("COL_LastAction"))
            .LastActionDate = FND(rs.fields("COL_LastActionDate"))
            .Fulfilled = FNS(rs.fields("COL_Fulfilled"))
            .Note = FNS(rs.fields("COL_Note"))
            .Price = FNN(rs.fields("COL_Price"))
            .FCID = FNN(rs.fields("COL_FCID"))
            .ForeignPrice = FNN(rs.fields("COL_FCPrice"))
            .FCFactor = FNDBL(rs.fields("COL_FCFactor"))
            .Deposit = FNN(rs.fields("COL_Deposit"))
            .DepositStatus = FNS(rs.fields("COL_DepositStatus"))
            .VATRate = FNS(rs.fields("P_VATRATE"))
            .Title = FNS(rs.fields("P_Title"))
            .Author = FNS(rs.fields("P_MainAuthor"))
            .Publisher = FNS(rs.fields("P_Publisher"))
            .ServiceItem = FNS(rs.fields("P_ProductType")) = "N"
            .Discount = FNDBL(rs.fields("COL_DiscountPercent"))
            .ETA = FND(rs.fields("COL_ETA"))
            .ETACode = FNS(rs.fields("COL_ETACode"))
            .Ref = FNS(rs.fields("COL_REF"))
            .POLID = FNN(rs.fields("COL_SOL_ID"))
            .WantDate = FND(rs.fields("COL_WantDate"))
            .DateReplaced = CDate(0) 'We do not fetch replaced lines here
            .Replacementfor = FNN(rs.fields("COL_ReplacementFor"))
            .CodeForExport = FNS(rs.fields("CODEF"))
            .CodeF = FNS(rs.fields("CODEFLocal"))
            .code = FNS(rs.fields("P_CODE"))
            .EAN = FNS(rs.fields("P_EAN"))
            .ExtraPID = FNS(rs.fields("COL_AdditionalProductPID"))
            .ExtraCharge = FNN(rs.fields("COL_AdditionalProductValue"))
            .ExtraCode = FNS(rs.fields("ExtraCode"))
            .ExtraChargeDescription = FNS(rs.fields("ExtraChargeDescription"))
            .SupplierID = FNS(rs.fields("SupplierID"))
          '  .ExtraVATRate = FND(rs.Fields("COL_AdditionalProductVatRate"))
            .DeliveryDocument = FNS(rs.fields("DeliveryDoc"))
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
        End With
        LSet udtData = udtProps
        lngCount = lngCount + 1
        oPBOUT.WriteProperty "Item" & CStr(lngCount), udtData.buffer
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    With oPBOUT
          .WriteProperty "Count", CStr(lngCount)
          Fetch = .Contents
    End With
    Set oPBOUT = Nothing
    Exit Function
errHandler:
    ErrPreserve
    If OpenResult = 0 Then oPC.DisconnectDBShort
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL_P.Fetch(ID)", , EA_DFTRBK, oPC.COShort
End Function



Public Function Save(ByVal buffer As String, ByVal ID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As COLProps
Dim udtData As COLData
Dim oPBIN As PropertyBag
Dim oPB2IN As PropertyBag
Dim oPBOUT As PropertyBag
Dim oPB2OUT As PropertyBag
Dim oPBCOFF As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
Dim objPersist As a_COFF_P
Dim OpenResult As Integer

    Set oPBIN = New PropertyBag
    Set oPBOUT = New PropertyBag
    arBuffer = buffer
    oPBIN.Contents = arBuffer
    Set rs = New ADODB.Recordset
    
    For lngIndex = 1 To oPBIN.ReadProperty("Count")
          udtData.buffer = oPBIN.ReadProperty("Item" & CStr(lngIndex))
          LSet udtProps = udtData
    '-------------------------------
      OpenResult = oPC.OpenDBSHort
    '-------------------------------
        If Not udtProps.IsDeleted Then
            strSQL = "SELECT * FROM tCOL WHERE COL_ID=" & udtProps.COLineID
            rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then
                rs.AddNew
                rs.fields("COL_FulFilled") = "OS"
            Else
              rs.fields("COL_Fulfilled") = FNS(udtProps.Fulfilled)
            End If
            With udtProps
                rs.fields("COL_Seq") = FNN(.Sequence)
                rs.fields("COL_TR_ID") = ID
                rs.fields("COL_P_ID") = FNS(.ProdID)
                rs.fields("COL_Qty") = FNN(.Qty)
                If oPC.AllowsSSInvoicing = True Then
                    rs.fields("COL_Qty") = FNN(.QtyFirm) + FNN(.QtySS)
                End If
                rs.fields("COL_QtyFirm") = FNN(.QtyFirm)
                rs.fields("COL_QtySS") = FNN(.QtySS)
                rs.fields("COL_QtyDispatched") = FNN(.QtyDispatched)
                rs.fields("COL_ActionTaken") = FNN(.ActionTaken)
                rs.fields("COL_Deposit") = FNN(.Deposit)
                rs.fields("COL_DepositStatus") = FNS(.DepositStatus)
                rs.fields("COL_DiscountPercent") = Round(FNDBL(.Discount), 2)
                rs.fields("COL_Price") = FNN(.Price)
                rs.fields("COL_FCPrice") = FNN(.ForeignPrice)
                rs.fields("COL_FCFactor") = FNDBL(.FCFactor)
                rs.fields("COL_FCID") = FNN(.FCID)
                rs.fields("COL_Note") = FNS(.Note)
                rs.fields("COL_ETA") = FND(.ETA)
                rs.fields("COL_ETACode") = FNS(.ETACode)
                rs.fields("COL_REF") = FNS(.Ref)
                rs.fields("COL_REPLACEMENTFOR") = FNN(.Replacementfor)
                rs.fields("COL_WANTDATE") = FND(.WantDate)
                If FNS(.ExtraPID) > "" Then rs.fields("COL_AdditionalProductPID") = FNS(.ExtraPID)
                rs.fields("COL_AdditionalProductValue") = FNN(.ExtraCharge)
                
                rs.Update
                If .IsNew Then
                  rs.Bookmark = rs.Bookmark
                  .COLineID = rs("COL_ID")
                End If
                .IsNew = False
                .IsDirty = False
            End With
            
            LSet udtData = udtProps
            
            lngCount = lngCount + 1
            oPBOUT.WriteProperty "Item" & CStr(lngCount), udtData.buffer   'oPB2OUT.Contents
            rs.Close
        Else
            DeleteCOLine udtProps.COLineID
        End If
        oPC.COShort.execute " UPDATE tCOL SET COL_FULFILLED = 'CAN' ,COL_DATEREPLACED = GETDATE() WHERE COL_ID = " & FNN(udtProps.Replacementfor) & " AND ISNULL(COL_DATEREPLACED,'1900-01-01') < '1950-01-01'"
    Next
    oPBOUT.WriteProperty "Count", lngCount
    
    Set oPBIN = Nothing
    Set rs = Nothing
    Save = oPBOUT.Contents
    Set oPBOUT = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    ErrPreserve
    Set oPBIN = Nothing
    Set oPBOUT = Nothing
    If OpenResult = 0 Then oPC.DisconnectDBShort
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL_P.Save(buffer,ID)", Array(buffer, ID), EA_DFTRBK, oPC.COShort
    Exit Function
    Resume
End Function

'Private Sub DeleteCOLine(COLID As Long)
'    On Error GoTo ErrHandler
'Dim OpenResult As Integer
'
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    oPC.COShort.execute "DELETE FROM tCOL WHERE COL_ID=" & COLID
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'    Exit Sub
'ErrHandler:
'    ErrPreserve
'    If OpenResult = 0 Then oPC.DisconnectDBShort
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_COL_P.DeleteCOLine(COLID)", COLID, EA_DFTRBK, oPC.COShort
'End Sub
''''''''Public Function Save(ByVal buffer As String, ByVal ID As Long, Optional pRef As String) As String
''''''''    On Error GoTo ErrHandler
''''''''Dim strSQL As String
''''''''Dim udtProps As COLProps
''''''''Dim udtData As COLData
''''''''Dim objPB As PropertyBag
''''''''Dim objPBOut As PropertyBag
''''''''Dim arBuffer() As Byte
''''''''Dim lngIndex As Long
''''''''Dim lngCount As Long
''''''''Dim cmd As ADODB.Command
''''''''Dim par As ADODB.Parameter
''''''''
''''''''  Set objPB = New PropertyBag
''''''''  Set objPBOut = New PropertyBag
''''''''  arBuffer = buffer
''''''''  objPB.Contents = arBuffer
''''''''
''''''''    For lngIndex = 1 To objPB.ReadProperty("Count")
''''''''      Set oPB2IN = New PropertyBag
''''''''      With oPB2IN
''''''''          arBuffer = oPBIN.ReadProperty("Item" & CStr(lngIndex))
''''''''          .Contents = arBuffer
''''''''          udtData.buffer = .ReadProperty("State")
''''''''          LSet udtProps = udtData
''''''''      End With
''''''''      udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
''''''''      LSet udtProps = udtData
''''''''      With udtProps
''''''''      If Not .IsDeleted Then
''''''''          Set cmd = New ADODB.Command
''''''''          cmd.ActiveConnection = oPC.COShort
''''''''          cmd.CommandType = adCmdStoredProc
''''''''          cmd.CommandText = "pers_COL"
''''''''          Set par = cmd.CreateParameter("@NU", adChar, adParamInput, 1)
''''''''          cmd.Parameters.Append par
''''''''          If udtProps.IsNew Then
''''''''              par.Value = "I"
''''''''          Else
''''''''              par.Value = "U"
''''''''          End If
''''''''              'rs.Fields("COL_TR_ID") = ID
''''''''          Set par = cmd.CreateParameter("@COLID", adInteger, adParamInput)
''''''''          cmd.Parameters.Append par
''''''''          par.Value = FNN(.COLineID)
''''''''              'rs.Fields("COL_P_ID") = FNS(.ProdID)
''''''''          Set par = cmd.CreateParameter("@PID", adGUID, adParamInput)
''''''''          cmd.Parameters.Append par
''''''''          par.Value = FNS(.ProdID)
''''''''              'rs.Fields("COL_Qty") = FNN(.Qty)
''''''''          Set par = cmd.CreateParameter("@QTY", adInteger, adParamInput)
''''''''          cmd.Parameters.Append par
''''''''          par.Value = FNN(.Qty)
''''''''              'rs.Fields("COL_QtyDispatched") = FNN(.QtyDispatched)
''''''''          Set par = cmd.CreateParameter("@QTYDISPATCHED", adInteger, adParamInput)
''''''''          cmd.Parameters.Append par
''''''''          par.Value = FNN(.QtyDispatched)
''''''''              'rs.Fields("COL_ActionTaken") = FNN(.ActionTaken)
''''''''          Set par = cmd.CreateParameter("@ACTIONTAKEN", adInteger, adParamInput)
''''''''          cmd.Parameters.Append par
''''''''          par.Value = FNN(.ActionTaken)
''''''''              'rs.Fields("COL_Deposit") = FNN(.Deposit)
''''''''          Set par = cmd.CreateParameter("@DEPOSIT", adInteger, adParamInput)
''''''''          cmd.Parameters.Append par
''''''''          par.Value = FNN(.Deposit)
''''''''              'rs.Fields("COL_DepositStatus") = FNS(.DepositStatus)
''''''''          Set par = cmd.CreateParameter("@DEPOSITSTATUS", adChar, adParamInput, 1)
''''''''          cmd.Parameters.Append par
''''''''          par.Value = FNS(.DepositStatus)
''''''''              'rs.Fields("COL_DiscountPercent") = Round(FNDBL(.Discount), 2)
''''''''          Set par = cmd.CreateParameter("@DISCOUNT", adDouble, adParamInput)
''''''''          cmd.Parameters.Append par
''''''''          par.Value = Round(FNDBL(.Discount), 2)
''''''''              'rs.Fields("COL_Price") = FNN(.Price)
''''''''          Set par = cmd.CreateParameter("@PRICE", adInteger, adParamInput)
''''''''          cmd.Parameters.Append par
''''''''          par.Value = FNN(.Price)
''''''''              'rs.Fields("COL_Note") = FNS(.Note)
''''''''          Set par = cmd.CreateParameter("@NOTE", adVarChar, adParamInput, 100)
''''''''          cmd.Parameters.Append par
''''''''          par.Value = FNS(.Note)
''''''''              'rs.Fields("COL_ETA") = FND(.ETA)
''''''''          Set par = cmd.CreateParameter("@ETA", adDate, adParamInput)
''''''''          cmd.Parameters.Append par
''''''''          par.Value = ReverseDate(FND(.ETA))
''''''''              'rs.Fields("COL_ETACode") = FNS(.ETACode)
''''''''          Set par = cmd.CreateParameter("@ETACODE", adVarChar, adParamInput, 1)
''''''''          cmd.Parameters.Append par
''''''''          par.Value = FNS(.ETACode)
''''''''              'rs.Fields("COL_REF") = FNS(.Ref)
''''''''          Set par = cmd.CreateParameter("@REF", adVarChar, adParamInput, 30)
''''''''          cmd.Parameters.Append par
''''''''          If FNS(udtProps.Ref) = "" Then
''''''''              par.Value = FNS(pRef)
''''''''          Else
''''''''              par.Value = FNS(.Ref)
''''''''          End If
''''''''              'rs.Fields("COL_WANTDATE") = FND(.WantDate)
''''''''          Set par = cmd.CreateParameter("@WANTDATE", adDate, adParamInput)
''''''''          cmd.Parameters.Append par
''''''''          par.Value = ReverseDate(FND(.WantDate))
''''''''
''''''''          Set par = cmd.CreateParameter("@FULFILLED", adVarChar, adParamInput, 4)
''''''''          cmd.Parameters.Append par
''''''''          If udtProps.IsNew Then
''''''''              par.Value = "OS"
''''''''          Else
''''''''              par.Value = FNS(udtProps.Fulfilled)
''''''''          End If
''''''''          cmd.execute
''''''''          If .IsNew Then .COLineID = cmd.Parameters("@ID")
''''''''              .IsNew = False
''''''''              .IsDirty = False
''''''''              LSet udtData = udtProps
''''''''              lngCount = lngCount + 1
''''''''              objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
''''''''          Else
''''''''              DeleteObject udtProps.POLID
''''''''          End If
''''''''          End With
''''''''      Next
''''''''      objPBOut.WriteProperty "Count", lngCount
''''''''
''''''''    Set objPB = Nothing
''''''''
''''''''    Save = objPBOut.Contents
''''''''    Set objPBOut = Nothing
''''''''
''''''''    Exit Function
''''''''ErrHandler:
''''''''    ErrPreserve
''''''''    Set objPB = Nothing
''''''''    Set objPBOut = Nothing
''''''''    If ErrMustStop Then Debug.Assert False: Resume
''''''''    ErrorIn "a_COL_P.Save(buffer,ID)", Array(buffer, ID), EA_DFTRBK, oPC.COShort
''''''''End Function

Private Sub DeleteCOLine(COLID As Long)
    On Error GoTo errHandler
Dim OpenResult As Integer
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "DELETE FROM tCOL WHERE COL_ID=" & COLID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    ErrPreserve
    If OpenResult = 0 Then oPC.DisconnectDBShort
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL_P.DeleteCOLine(COLID)", COLID, EA_DFTRBK, oPC.COShort
End Sub

Public Sub DeleteObject(ByVal COID As Long)
    On Error GoTo errHandler
Dim OpenResult As Integer
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
  oPC.COShort.execute "DELETE FROM tCOLine WHERE COL_TR_ID=" & COID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
  Exit Sub


    Exit Sub
errHandler:
    ErrPreserve
    If OpenResult = 0 Then oPC.DisconnectDBShort
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COL_P.DeleteObject(COID)", COID, EA_DFTRBK, oPC.COShort
End Sub


