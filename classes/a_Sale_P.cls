VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Sale_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Option Explicit

Public Function Fetch(ByVal pCSLID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As SaleProps
Dim udtData As SaleData
Dim oPBOUT As PropertyBag
Dim oPBCOFF As PropertyBag
Dim lngCount As Long
Dim i As Long
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM vCSL WHERE CSL_ID=" & pCSLID
    rs.Open strSQL, oPC.DBConn, adOpenForwardOnly, adLockReadOnly
    Set oPBOUT = New PropertyBag
    i = 0
    Do While Not rs.EOF
        With udtProps
            .ID = FNN(rs.Fields("CSL_ID"))
            .EXCHANGE_GUID = FNS(rs.Fields("CSL_EXCH_ID"))
            .pID = FNS(rs.Fields("CSL_P_ID"))
            .Title = FNS(rs.Fields("P_Title"))
            .Author = FNS(rs.Fields("P_MainAuthor"))
            .Qty = FNN(rs.Fields("CSL_Qty"))
            .Price = FNN(rs.Fields("CSL_Price"))
            .DiscountRate = FNDBL(rs.Fields("CSL_DiscountRate"))
            .Discount = FNN(rs.Fields("CSL_Discount"))
            .VATRate = FNDBL(rs.Fields("CSL_VATRATE"))
            .VAT = FNN(rs.Fields("CSL_VAT"))
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
        End With
        LSet udtData = udtProps
        
        Set oPBCOFF = New PropertyBag
        
        oPBCOFF.WriteProperty "State", udtData.buffer
        lngCount = lngCount + 1
        oPBOUT.WriteProperty "Item" & CStr(lngCount), oPBCOFF.Contents
        Set oPBCOFF = Nothing
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
  
'ERRH:
'  rs.Close
'  Set rs = Nothing
'  Err.Raise Err.Number
'  Exit Function
'  Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale_P.Fetch(pCSLID)", pCSLID
End Function



Public Function Save(ByVal buffer As String, ByVal pEXCHID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As SaleProps
Dim udtData As SaleData

Dim oPBIN As PropertyBag
Dim oPB2IN As PropertyBag
Dim oPBOUT As PropertyBag
Dim oPB2OUT As PropertyBag
Dim oPBCOFF As PropertyBag

Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
  
  Set oPBIN = New PropertyBag
  Set oPBOUT = New PropertyBag
  Set rs = New ADODB.Recordset
  
  arBuffer = buffer
  oPBIN.Contents = arBuffer
  
  
  For lngIndex = 1 To oPBIN.ReadProperty("Count")
        Set oPB2IN = New PropertyBag
        With oPB2IN
            arBuffer = oPBIN.ReadProperty("Item" & CStr(lngIndex))
            .Contents = arBuffer
            udtData.buffer = .ReadProperty("State")
            LSet udtProps = udtData
        End With
    If Not udtProps.IsDeleted Then
      strSQL = "SELECT * FROM tCSL WHERE CSL_ID=" & udtProps.ID
      rs.Open strSQL, oPC.DBConn, adOpenKeyset, adLockOptimistic
      If udtProps.IsNew Then rs.AddNew
      
      With udtProps
      '  rs.Fields("CSL_ID") = FNN(.ID)
        rs.Fields("CSL_EXCH_ID") = FNS(pEXCHID)
        rs.Fields("CSL_P_ID") = FNS(.pID)
     '   rs.Fields("P_Title") = FNS(.Title)
     '   rs.Fields("IL_Qty") = FNN(.Qty)
     '   rs.Fields("P_MainAuthor") = FNS(.Author)
        rs.Fields("CSL_Qty") = FNN(.Qty)
        rs.Fields("CSL_Price") = FNN(.Price)
        rs.Fields("CSL_DiscountRate") = FNDBL(.DiscountRate)
        rs.Fields("CSL_Discount") = FNN(.Discount)
        rs.Fields("CSL_VATRate") = FNDBL(.VATRate)
        rs.Fields("CSL_VAT") = FNN(.VAT)
        
        rs.Update
        
        If .IsNew Then
          rs.Bookmark = rs.Bookmark
          .ID = rs("CSL_ID")
        End If
        .IsNew = False
        .IsDirty = False
      End With
      
        LSet udtData = udtProps
        Set oPB2OUT = New PropertyBag
        oPB2OUT.WriteProperty "State", udtData.buffer
    
        lngCount = lngCount + 1
        oPBOUT.WriteProperty "Item" & CStr(lngCount), oPB2OUT.Contents
        rs.Close
        Set oPB2IN = Nothing
        Set oPB2OUT = Nothing
        Set oPBCOFF = Nothing
    Else
      DeleteCSL udtProps.ID
    End If
  Next
  oPBOUT.WriteProperty "Count", lngCount
  
  Set oPBIN = Nothing
  Set rs = Nothing
  
  Save = oPBOUT.Contents
  Set oPBOUT = Nothing
  Exit Function
  
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale_P.Save(buffer,pEXCHID)", Array(buffer, pEXCHID)
End Function

Private Sub DeleteCSL(pCSLID As Long)
    On Error GoTo errHandler
    oPC.DBConn.Execute "DELETE FROM tCSL WHERE CSL_ID=" & pCSLID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale_P.DeleteCSL(pCSLID)", pCSLID
End Sub

Public Sub DeleteObject(ByVal pEXCHID As String)
    On Error GoTo errHandler

  oPC.DBConn.Execute "DELETE FROM tCSL WHERE CSL_EXCH_ID=" & pEXCHID
  Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale_P.DeleteObject(pEXCHID)", pEXCHID
End Sub


