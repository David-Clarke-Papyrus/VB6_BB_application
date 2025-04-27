VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_CNL_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByVal ID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As CNLProps
Dim udtData As CNLData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim oProdCode As z_ProdCode
'On Error GoTo H
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM Fetch_CNL WHERE CNL_TR_ID=" & ID & " ORDER BY P_TITLE ASC"
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rs.eof
        With udtProps
            .CNLineID = rs.fields("CNL_ID")
            .TRID = rs.fields("CNL_TR_ID")
            .PID = FNS(rs.fields("CNL_P_ID"))
            .PIID = FNN(rs.fields("CNL_PI_ID"))
            .Qty = FNN(rs.fields("CNL_Qty"))
            .QtyDAM = FNN(rs.fields("CNL_QtyDam"))
            .Note = FNS(rs.fields("CNL_Note"))
            .InvPrice = FNN(rs.fields("Price"))
            .InvLineRef = FNS(rs.fields("IL_REF"))
            .ServiceItem = FNS(rs.fields("P_ProductType")) = "N"
         '   .VATValue = FNN(rs.Fields("CNL_VATValue"))
            .VATRate = FNDBL(rs.fields("VATRate"))
            .InvForeignPrice = FNN(rs.fields("IL_ForeignPrice"))
            .ProductCode = FNS(rs.fields("P_Code")) & IIf(IsNull(rs.fields("PI_Serial")), "", "/" & FNS(rs.fields("PI_Serial")))
            .Title = FNS(rs.fields("P_Title"))
            .Author = FNS(rs.fields("P_MainAuthor"))
            .Publisher = FNS(rs.fields("P_Publisher"))
            .InvLineDate = FND(rs.fields("InvDate"))
            .INVLineCode = FNS(rs.fields("INVCODE"))
            .INVLineID = FNN(rs.fields("CNL_IL_ID"))
            .DiscountRate = FNDBL(rs.fields("DiscountRate"))
            .ProductCodeForExport = FNS(rs.fields("CODEFEX"))
            .ProductCodeF = FNS(rs.fields("CODEF")) & IIf(IsNull(rs.fields("PI_Serial")), "", "/" & FNS(rs.fields("PI_Serial")))
            .ProductCode = FNS(rs.fields("P_Code"))
            .EAN = FNS(rs.fields("P_EAN"))
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
  
'H:
'  rs.Close
'  Set rs = Nothing
'  Err.Raise err.Number
    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL_P.Fetch(ID)", ID
End Function



Public Function Save(ByVal buffer As String, ByVal ID As Long) As String
    On Error GoTo errHandler
  Dim rs As ADODB.Recordset
  Dim strSQL As String
  Dim udtProps As CNLProps
  Dim udtData As CNLData
  Dim objPB As PropertyBag
  Dim objPBOut As PropertyBag
  Dim arBuffer() As Byte
  Dim lngIndex As Long
  Dim lngCount As Long
  
  Set objPB = New PropertyBag
  Set objPBOut = New PropertyBag
  arBuffer = buffer
  objPB.Contents = arBuffer
  Set rs = New ADODB.Recordset
  
  For lngIndex = 1 To objPB.ReadProperty("Count")
    udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
    LSet udtProps = udtData
    
    If Not udtProps.IsDeleted Then
      strSQL = "SELECT * FROM tCNL WHERE CNL_ID=" & udtProps.CNLineID
      rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
      If udtProps.IsNew Then rs.AddNew
      
      With udtProps
        rs.fields("CNL_TR_ID") = ID
        rs.fields("CNL_P_ID") = FNS(.PID)
        rs.fields("CNL_INVCode") = FNS(.INVLineCode)
        rs.fields("CNL_PI_ID") = FNN(.PIID)
        rs.fields("CNL_IL_ID") = FNN(.INVLineID)
        rs.fields("CNL_Qty") = FNN(.Qty)
        rs.fields("CNL_QtyDam") = FNN(.QtyDAM)
        rs.fields("CNL_Price") = FNN(.InvPrice)
        rs.fields("CNL_VATRATE") = FNDBL(.VATRate)
        rs.fields("CNL_DISCOUNTRATE") = FNDBL(.DiscountRate)
      '  rs.Fields("CNL_SalesValue") = FNN(.SalesValue)
      '  rs.Fields("CNL_VATValue") = FNN(.VATValue)
        rs.fields("CNL_Note") = FNS(.Note)
        rs.Update
        
        If .IsNew Then
          rs.Bookmark = rs.Bookmark
          .CNLineID = rs("CNL_ID")
        End If
        .IsNew = False
        .IsDirty = False
      End With
      LSet udtData = udtProps
      lngCount = lngCount + 1
      objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
      rs.Close
    Else
      DeleteCNLine udtProps.CNLineID
    End If
  Next
  objPBOut.WriteProperty "Count", lngCount
  
  Set objPB = Nothing
  Set rs = Nothing
  
  Save = objPBOut.Contents
  Set objPBOut = Nothing
  Exit Function
  
'H:
'  Set objPB = Nothing
'  Set objPBOut = Nothing
'  Set rs = Nothing
'  Err.Raise err.Number
'  Exit Function
'  Resume
    Exit Function
errHandler:
    ErrPreserve
    Set objPB = Nothing
    Set objPBOut = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL_P.Save(buffer,ID)", Array(buffer, ID)
End Function

Private Sub DeleteCNLine(CNLID As Long)
    On Error GoTo errHandler
    oPC.COShort.execute "DELETE FROM tCNL WHERE CNL_ID=" & CNLID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL_P.DeleteCNLine(CNLID)", CNLID
End Sub

Public Sub DeleteObject(ByVal CNID As Long)
    On Error GoTo errHandler

  oPC.COShort.execute "DELETE FROM tCNL WHERE CNL_TR_ID=" & CNID
  Exit Sub

'H:
'  Err.Raise err.Number
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CNL_P.DeleteObject(CNID)", CNID
End Sub


