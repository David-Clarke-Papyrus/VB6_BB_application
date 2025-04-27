VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_QUL_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByVal QuoteID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As QULProps
Dim udtData As QULData
Dim oPBOUT As PropertyBag
Dim lngCount As Long
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM vQUL_Fetch WHERE QUL_TR_ID=" & QuoteID & " ORDER BY QUL_SEQ ASC"
    rs.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Set oPBOUT = New PropertyBag
    lngCount = 0
    Do While Not rs.eof
        With udtProps
            .Sequence = FNN(rs.fields("QUL_SEQ"))
           ' If .Sequence = 0 Then .Sequence = lngCount
            .QULID = rs.fields("QUL_ID")
            .QUID = rs.fields("QUL_TR_ID")
            .PID = FNS(rs.fields("QUL_P_ID"))
            .Qty = FNN(rs.fields("QUL_Qty"))
            .Note = FNS(rs.fields("QUL_Note"))
            .DiscountPercent = FNDBL(rs.fields("QUL_DiscountPercent"))
            .Price = FNN(rs.fields("QUL_Price"))
            .FCID = FNN(rs.fields("QUL_FCID"))
            .ForeignPrice = FNN(rs.fields("QUL_FCPrice"))
            .FCFactor = FNDBL(rs.fields("QUL_FCFactor"))
            .VATRate = FNDBL(rs.fields("QUL_VATRATE"))
            .code = IIf(Len(FNS(rs.fields("P_Code"))) > 0, FNS(rs.fields("P_Code")), FNS(rs.fields("P_EAN")))
            .EAN = FNS(rs.fields("P_EAN"))
            .MainAuthor = FNS(rs.fields("P_MainAuthor"))
            .Title = FNS(rs.fields("P_Title"))
            .Publisher = FNS(rs.fields("P_Publisher"))
            .ServiceItem = FNS(rs.fields("P_ProductType")) = "N"
            .Ref = FNS(rs.fields("QUL_Ref"))
            .ID = FNS(rs.fields("Id"))
            .COLID = FNN(rs.fields("QUL_COL_ID"))
            If (.code > "") Then
                .CodeForExport = FNS(rs.fields("CODEFEX"))
                .CodeF = FNS(rs.fields("CODEF"))
            End If
            .ExtraPID = FNS(rs.fields("QUL_AdditionalProductPID"))
            .ExtraCharge = FNN(rs.fields("QUL_AdditionalProductValue"))
            .ExtraCode = FNS(rs.fields("ExtraCode"))
            .ExtraChargeDescription = FNS(rs.fields("ExtraChargeDescription"))
            .ExtraVATRate = FND(rs.fields("QUL_AdditionalProductVatRate"))
           .IsNew = False
            .IsDirty = False
            .IsDeleted = False
        End With
        LSet udtData = udtProps
        lngCount = lngCount + 1
        oPBOUT.WriteProperty "Item" & CStr(lngCount), udtData.buffer   ', oPBCOFF.Contents
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
    ErrorIn "a_QUL_P.Fetch(QuoteID)", QuoteID
End Function



Public Function Save(ByVal buffer As String, ByVal QuoteID As Long, Optional pUpdateStock As Boolean) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As QULProps
Dim udtData As QULData

Dim oPBIN As PropertyBag
Dim oPBOUT As PropertyBag

Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
  
  Set oPBIN = New PropertyBag
  Set oPBOUT = New PropertyBag
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseClient
  
  arBuffer = buffer
  oPBIN.Contents = arBuffer
  
  For lngIndex = 1 To oPBIN.ReadProperty("Count")
    udtData.buffer = oPBIN.ReadProperty("Item" & CStr(lngIndex))
    LSet udtProps = udtData
    If Not udtProps.IsDeleted Then
      strSQL = "SELECT * FROM tQUL WHERE QUL_ID=" & udtProps.QULID
      rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
      If udtProps.IsNew Then rs.AddNew
      
      With udtProps
        rs.fields("QUL_Seq") = FNN(.Sequence)
        rs.fields("QUL_TR_ID") = QuoteID
        rs.fields("QUL_P_ID") = FNS(.PID)
        rs.fields("QUL_Qty") = FNN(.Qty)
        rs.fields("QUL_DiscountPercent") = FNDBL(.DiscountPercent)
        rs.fields("QUL_Price") = FNN(.Price)
        rs.fields("QUL_FCPrice") = FNN(.ForeignPrice)
        rs.fields("QUL_FCFactor") = FNDBL(.FCFactor)
        rs.fields("QUL_FCID") = FNN(.FCID)
        rs.fields("QUL_VATRate") = FNN(.VATRate)
        rs.fields("QUL_Ref") = FNS(.Ref)
        rs.fields("QUL_Note") = FNS(.Note)
        rs.fields("QUL_COL_ID") = FNN(.COLID)
        If FNS(.ExtraPID) > "" Then rs.fields("QUL_AdditionalProductPID") = FNS(.ExtraPID)
        rs.fields("QUL_AdditionalProductValue") = FNN(.ExtraCharge)
        rs.fields("QUL_AdditionalProductVatRate") = FNN(.ExtraVATRate)
        rs.fields("Id") = FNS(.ID)
        rs.Update
        
        If .IsNew Then
          rs.Bookmark = rs.Bookmark
          .QULID = rs("QUL_ID")
        End If
        .IsNew = False
        .IsDirty = False
      End With
      
        LSet udtData = udtProps
        lngCount = lngCount + 1
        oPBOUT.WriteProperty "Item" & CStr(lngCount), udtData.buffer  'XXoPB2OUT.Contents
        rs.Close
    Else
      DeleteQuoteLine udtProps.QULID, udtProps.PID
    End If
  Next
  oPBOUT.WriteProperty "Count", lngCount
  
  Set oPBIN = Nothing
  Set rs = Nothing
  
  Save = oPBOUT.Contents
  Set oPBOUT = Nothing
  Exit Function
  
errHandler:
    ErrPreserve
    Set oPBIN = Nothing
    Set oPBOUT = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL_P.Save(buffer,QuoteID)", Array(buffer, QuoteID), , rs
End Function

Private Sub DeleteQuoteLine(ILID As Long, PID As String)
    On Error GoTo errHandler
    oPC.COShort.execute "DELETE FROM tQUL WHERE QUL_ID=" & ILID
    oPC.COShort.execute "EXECUTE UpdateQtyOHperPID2 '" & PID & "'"
 ' MsgBox "Line deleted"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL_P.DeleteInvoiceLine(ILID)", ILID
End Sub

Public Sub DeleteObject(ByVal QuoteID As Long)
    On Error GoTo errHandler

    oPC.COShort.execute "DELETE FROM tQUL WHERE QUL_TR_ID=" & QuoteID
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL_P.DeleteObject(QuoteID)", QuoteID
End Sub



