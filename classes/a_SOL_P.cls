VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_POL_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByVal pTRID As Long, Optional pStatus As Integer) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As POLProps
Dim udtData As POLData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim oProdCode As z_ProdCode
'On Error GoTo H
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM POL_Fetch WHERE POL_TR_ID=" & pTRID & " ORDER BY POL_SEQ ASC"
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.eof
        With udtProps
            .Sequence = FNN(rs.fields("POL_SEQ"))
            .POLID = rs.fields("POL_ID")
            .TRID = rs.fields("POL_TR_ID")
            .PID = FNS(rs.fields("POL_P_ID"))
            .QtySS = FNN(rs.fields("POL_QtySS"))
            .QtyFirm = FNN(rs.fields("POL_QtyFirm"))
            .QtyReceivedSoFar = FNN(rs.fields("POL_QtyReceivedSoFar"))
            If pStatus = 2 Then
                .ProductTypeID = FNN(rs.fields("POL_ProductTypeID"))
            Else
                .ProductTypeID = FNN(rs.fields("P_ProductType_ID"))
            End If
            .Note = FNS(rs.fields("POL_Note"))
            .Discount = FNDBL(rs.fields("POL_Discount"))
            .Price = FNN(rs.fields("POL_Price"))
            .ForeignPrice = FNN(rs.fields("POL_ForeignPrice"))
            .Publisher = FNS(rs.fields("P_Publisher"))
            .PublicationDate = FNS(rs.fields("P_PubDate"))
            .Edition = FNS(rs.fields("P_Edition"))
    '5/4/2005 This is fetched from the POL line if the status is In Process,
    'otherwise it is assumed that the product record carries the correct value
 '           If pStatus = 2 And udtProps.IsNew Then
 '               .ProductTypeID = FNN(rs.Fields("POL_ProductTypeID"))
 '           Else
   '             .ProductTypeID = FNN(rs.Fields("P_ProductType_ID"))
 '           End If
            
            .DealID = FNN(rs.fields("POL_DL_ID"))
            .Ref = FNS(rs.fields("POL_REF"))
            .COLID = FNN(rs.fields("POL_COLID"))
            .Fulfilled = FNS(rs.fields("POL_FulFilled"))
            .ProductCode = FNS(rs.fields("P_Code"))
            .MainAuthor = FNS(rs.fields("P_MainAuthor"))
            .Section = FNS(rs.fields("POL_Section"))
            .Title = FNS(rs.fields("P_Title"))
            .ETA = FND(rs.fields("POL_ETA"))
            .LastAction = FNS(rs.fields("POL_LastAction"))
            .LastActionDate = FND(rs.fields("POL_LastActionDate"))
            .VATRate = FNDBL(rs.fields("POL_VATRATE"))
            .DateReplaced = FND(rs.fields("POL_DATEREPLACED"))
            .Replacementfor = FNN(rs.fields("POL_ReplacementFor"))
            .ProductCodeForExport = FNS(rs.fields("CodefEx"))
            .ProductCodeF = FNS(rs.fields("CodeF"))
            .EAN = FNS(rs.fields("P_EAN"))
            .COQty = FNN(rs.fields("COQty"))
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
          .WriteProperty "Count", CStr(lngCount)
          Fetch = .Contents
    End With
    Set objPB = Nothing
    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL_P.Fetch(pTRID)", pTRID
End Function



Public Function Save(ByVal buffer As String, ByVal pTRID As Long, Optional pRef As String) As String
    On Error GoTo errHandler
Dim strSQL As String
Dim udtProps As POLProps
Dim udtData As POLData
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
  
  For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData
        With udtProps
            If Not udtProps.IsDeleted Then
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = oPC.COShort
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "pers_POL"
                Set par = cmd.CreateParameter("@NU", adChar, adParamInput, 1)
                cmd.Parameters.Append par
                If udtProps.IsNew Then
                    par.Value = "I"
                Else
                    par.Value = "U"
                End If
                Set par = cmd.CreateParameter("@COLID", adInteger, adParamInput)
                cmd.Parameters.Append par
                par.Value = FNN(.COLID)
                
                Set par = cmd.CreateParameter("@DISCOUNT", adDouble, adParamInput)
                cmd.Parameters.Append par
                par.Value = Round(FNDBL(.Discount), 2)
                
                Set par = cmd.CreateParameter("@DLID", adInteger, adParamInput)
                cmd.Parameters.Append par
                par.Value = FNN(.DealID)
                
                Set par = cmd.CreateParameter("@ETA", adDate, adParamInput)
                par.Value = ReverseDate(FND(.ETA))
                cmd.Parameters.Append par
                
                Set par = cmd.CreateParameter("@FOREIGNPRICE", adInteger, adParamInput)
                par.Value = FNN(.ForeignPrice)
                cmd.Parameters.Append par
                
                Set par = cmd.CreateParameter("@FULFILLED", adVarChar, adParamInput, 4)
                cmd.Parameters.Append par
                If udtProps.IsNew Then
                    par.Value = "OS"
                Else
                    par.Value = FNS(udtProps.Fulfilled)
                End If
                
                Set par = cmd.CreateParameter("@LASTACTION", adVarChar, adParamInput, 1)
                cmd.Parameters.Append par
                par.Value = FNS(.LastActionCode)
                
                Set par = cmd.CreateParameter("@LASTACTIONDATE", adVarChar, adParamInput, 10)
                cmd.Parameters.Append par
                par.Value = ReverseDate(FND(.LastActionDate))
                Set par = cmd.CreateParameter("@NOTE", adVarChar, adParamInput, 100)
                cmd.Parameters.Append par
                par.Value = FNS(.Note)
                Set par = cmd.CreateParameter("@PID", adGUID, adParamInput)
                cmd.Parameters.Append par
                par.Value = FNS(.PID)
                Set par = cmd.CreateParameter("@POLID", adInteger, adParamInput)
                cmd.Parameters.Append par
                par.Value = FNN(.POLID)
                Set par = cmd.CreateParameter("@PRICE", adInteger, adParamInput)
                cmd.Parameters.Append par
                par.Value = FNN(.Price)
                Set par = cmd.CreateParameter("@PRODUCTTYPEID", adInteger, adParamInput)
                cmd.Parameters.Append par
                par.Value = FNN(.ProductTypeID)
                Set par = cmd.CreateParameter("@QTYFIRM", adInteger, adParamInput)
                cmd.Parameters.Append par
                par.Value = FNN(.QtyFirm)
                Set par = cmd.CreateParameter("@QTYRECEIVEDSOFAR", adInteger, adParamInput)
                cmd.Parameters.Append par
                par.Value = FNN(.QtyReceivedSoFar)
                Set par = cmd.CreateParameter("@QTYSS", adInteger, adParamInput)
                cmd.Parameters.Append par
                par.Value = FNN(.QtySS)
                Set par = cmd.CreateParameter("@REF", adVarChar, adParamInput, 30)
                cmd.Parameters.Append par
                If FNS(udtProps.Ref) = "" Then
                    par.Value = FNS(pRef)
                Else
                    par.Value = FNS(.Ref)
                End If
                Set par = cmd.CreateParameter("@SECTION", adVarChar, adParamInput, 51)
                cmd.Parameters.Append par
                par.Value = FNS(.Section)
                Set par = cmd.CreateParameter("@TRID", adInteger, adParamInput)
                cmd.Parameters.Append par
                par.Value = pTRID
                Set par = cmd.CreateParameter("@VATRATE", adDouble, adParamInput)
                cmd.Parameters.Append par
                par.Value = FNDBL(.VATRate)
        
                Set par = cmd.CreateParameter("@REPLACEMENTFOR", adInteger, adParamInput)
                cmd.Parameters.Append par
                par.Value = FNN(.Replacementfor)
                
                
                Set par = cmd.CreateParameter("@ID", adInteger, adParamOutput)
                cmd.Parameters.Append par
                
                Set par = cmd.CreateParameter("@SEQ", adInteger, adParamInput)
                cmd.Parameters.Append par
                par.Value = FNN(.Sequence)
                
                cmd.execute
               ' If .IsNew Then .POLID = cmd.Parameters("@SEQ")
                .IsNew = False
                .IsDirty = False
                LSet udtData = udtProps
                lngCount = lngCount + 1
                objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            Else
                DeletePOL udtProps.POLID
            End If
        End With
    Next
    objPBOut.WriteProperty "Count", lngCount
    
    Set objPB = Nothing
    
    Save = objPBOut.Contents
    Set objPBOut = Nothing
  
    Exit Function
errHandler:
    Set objPB = Nothing
    Set objPBOut = Nothing
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL_P.Save(buffer,pTRID,pRef)", Array(buffer, pTRID, pRef), , cmd
End Function

Private Sub DeletePOL(POLID As Long)
    On Error GoTo errHandler
    oPC.COShort.execute "DELETE FROM tPOL WHERE POL_ID=" & POLID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL_P.DeletePOL(POLID)", POLID
End Sub

Public Sub DeletePOLSet(ByVal TRID As Long)
    On Error GoTo errHandler

  oPC.COShort.execute "DELETE FROM tPOL WHERE POL_TR_ID=" & TRID
  Exit Sub

'H:
'  Err.Raise err.Number
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_POL_P.DeletePOLSet(TRID)", TRID
End Sub

'        strSQL = "SELECT * FROM tPOL WHERE POL_ID=" & udtProps.POLID
'        rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
'        If udtProps.IsNew Then
'            rs.AddNew
'            rs.Fields("POL_Fulfilled") = "OS"
'        Else
'            rs.Fields("POL_Fulfilled") = FNS(udtProps.Fulfilled)
'        End If

'        With udtProps
'            rs.Fields("POL_TR_ID") = pTRID
'            rs.Fields("POL_P_ID") = FNS(.PID)
'            rs.Fields("POL_QtySS") = FNN(.QtySS)
'            rs.Fields("POL_QtyFirm") = FNN(.QtyFirm)
'            rs.Fields("POL_QtyReceivedSoFar") = FNN(.QtyReceivedSoFar)
'            rs.Fields("POL_REF") = FNS(.Ref)
'            If FNS(.Ref) = "" Then
'                rs.Fields("POL_REF") = FNS(pRef)
'            Else
'                rs.Fields("POL_REF") = FNS(.Ref)
'            End If
'            rs.Fields("POL_DL_ID") = FNN(.DEALID)
'            rs.Fields("POL_Discount") = FNDBL(.Discount)
'            rs.Fields("POL_Price") = FNN(.Price)
'            rs.Fields("POL_COLID") = FNN(.COLID)
'            rs.Fields("POL_ETA") = FND(.ETA)
'            rs.Fields("POL_ForeignPrice") = FNN(.ForeignPrice)
'            rs.Fields("POL_VATRate") = FNDBL(.VATRate)
'            rs.Fields("POL_Note") = FNS(.Note)
'            rs.Fields("POL_Section") = FNS(.Section)
'            rs.Fields("POL_ProductTypeID") = FNN(.ProductTypeID)
'            rs.Fields("POL_LastAction") = FNS(.LastActionCode)
'            rs.Fields("POL_LastActionDate") = ReverseDate(FND(.LastActionDate))
'            rs.Update
'
'            If .IsNew Then
'              rs.Bookmark = rs.Bookmark
'              .POLID = rs.Fields("POL_ID")
'            End If
'
