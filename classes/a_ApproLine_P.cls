VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_APPL_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByVal ID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim rsI As ADODB.Recordset
Dim rsR As ADODB.Recordset
Dim strSQL As String
Dim udtProps As APPLProps
Dim udtData As APPLData
Dim oPBOUT As PropertyBag
Dim lngCount As Long
Dim oProdCode As z_ProdCode
Dim strInvoices As String
Dim strReturns As String
Dim strPreviousRecipients As String
Dim strPos As String

    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM APPL_Fetch WHERE APPL_TR_ID=" & ID
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Set oPBOUT = New PropertyBag
    Do While Not rs.eof
        With udtProps
        
            .APPLID = rs.fields("APPL_ID")
            .TRID = rs.fields("APPL_TR_ID")
           ' .Code = FNS(rs.Fields("P_Code"))
            .Title = FNS(rs.fields("P_Title"))
            .Ref = FNS(rs.fields("APPL_Ref"))
            .Author = FNS(rs.fields("P_MainAuthor"))
            .Price = FNN(rs.fields("APPL_Price"))
            .Qty = FNN(rs.fields("APPL_Qty"))
            .QtyReturned = FNN(rs.fields("APPL_QtyReturned"))
            .Discount = FNDBL(rs.fields("APPL_DiscountRate"))
            .PID = FNS(rs.fields("APPL_P_ID"))
            .Title = FNS(rs.fields("P_Title"))
            .VATRate = FNDBL(rs.fields("APPL_VATRATE"))
            .Fulfilled = FNS(rs.fields("APPL_Fulfilled"))
            .Note = Left(FNS(rs.fields("APPL_Note")), 35)
            .COLID = FNN(rs.fields("APPL_COL_ID"))
            
            strInvoices = ""
            Set rsI = New ADODB.Recordset
            rsI.open "SELECT * FROM tILine JOIN tTR on IL_TR_ID = TR_ID WHERE IL_APPL_ID = " & .APPLID & " ORDER BY TR_DATE", oPC.COShort, adOpenForwardOnly, adLockReadOnly
            Do While Not rsI.eof
                If strInvoices > "" Then strInvoices = strInvoices & ", "
                strInvoices = rsI!TR_CODE & "(" & rsI!IL_QTY & ")"
                rsI.MoveNext
            Loop
            rsI.Close
            Set rsI = Nothing
            .Invoices = strInvoices
            
            strReturns = ""
            Set rsR = New ADODB.Recordset
            rsR.open "SELECT * FROM tAPPRL JOIN tTR on APPRL_TR_ID = TR_ID WHERE APPRL_APPL_ID = " & .APPLID & " ORDER BY TR_DATE", oPC.COShort, adOpenForwardOnly, adLockReadOnly
            Do While Not rsR.eof
                If strReturns > "" Then strReturns = strReturns & ", "
                strReturns = rsR!TR_CODE & "(" & rsR!APPRL_Qty & ")"
                rsR.MoveNext
            Loop
            rsR.Close
            Set rsR = Nothing
            .Returns = strReturns
            
          '  Set oProdcode = New z_ProdCode
          '  oProdcode.Load rs.Fields("P_Code")
            .CodeF = FNS(rs.fields("CODEF"))
            .CodeFForExport = FNS(rs.fields("CODEFEX"))
            .EAN = FNS(rs.fields("P_EAN"))
          '  Set oProdcode = Nothing
    
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
    With oPBOUT
          .WriteProperty "Count", CStr(lngCount)
          Fetch = .Contents
    End With
    Set oPBOUT = Nothing

    Exit Function
errHandler:
    ErrPreserve
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL_P.Fetch(ID)", ID, , rs, "strPos", Array(strPos)
End Function

Public Function Save(ByVal buffer As String, ByVal TRID As Long) As String
    On Error GoTo errHandler
  Dim rs As ADODB.Recordset
  Dim strSQL As String
  Dim udtProps As APPLProps
  Dim udtData As APPLData
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
      strSQL = "SELECT * FROM tAPPL WHERE APPL_ID=" & CStr(udtProps.APPLID)
      rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
      If udtProps.IsNew Then
        rs.AddNew
            rs.fields("APPL_Fulfilled") = "OS"
        Else
            rs.fields("APPL_Fulfilled") = FNS(udtProps.Fulfilled)
      End If
      With udtProps
        rs!APPL_TR_ID = TRID
        rs!APPL_P_ID = FNS(.PID)
        rs!APPL_Qty = FNN(.Qty)
        rs!APPL_QtyReturned = FNN(.QtyReturned)
        rs!APPL_DiscountRate = FNDBL(.Discount)
        rs!APPL_Price = FNN(.Price)
        rs!APPL_VATRate = FNDBL(.VATRate)
        rs!APPL_Note = FNS(.Note)
        rs!APPL_Ref = FNS(.Ref)
        rs!APPL_COL_ID = FNN(.COLID)
        rs.Update
        
        If .IsNew Then
          rs.Bookmark = rs.Bookmark
          .APPLID = rs("APPL_ID")
        End If
        .IsNew = False
        .IsDirty = False
      End With
      LSet udtData = udtProps
      lngCount = lngCount + 1
      objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
      rs.Close
    Else
      DeleteApproLine udtProps.APPLID
    End If
  Next
  objPBOut.WriteProperty "Count", lngCount
  
  Set objPB = Nothing
  Set rs = Nothing
  
  Save = objPBOut.Contents
  Set objPBOut = Nothing
    Exit Function
errHandler:
    ErrPreserve
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL_P.Save(buffer,TRID)", Array(buffer, TRID)
End Function

Public Sub DeleteObject(ByVal TRID As Long)
    On Error GoTo errHandler
Dim strSQL As String
  
    strSQL = "DELETE FROM tAPPL WHERE APPL_TR_ID=" & TRID
    oPC.COShort.execute strSQL


    Exit Sub
errHandler:
    ErrPreserve
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL_P.DeleteObject(TRID)", TRID
End Sub

Private Sub DeleteApproLine(APPLID As Long)
    On Error GoTo errHandler
Dim strSQL As String
    strSQL = "DELETE FROM tAPPL WHERE APPL_ID=" & APPLID
    oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    ErrPreserve
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APPL_P.DeleteApproLine(APPLID)", APPLID
End Sub
