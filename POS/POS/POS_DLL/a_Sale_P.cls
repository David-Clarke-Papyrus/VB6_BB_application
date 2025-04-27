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

Public Function Fetch(ByVal pEXCHID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As SaleProps
Dim udtData As SaleData
Dim oPBOUT As PropertyBag
Dim lngCount As Long
Dim i As Long
    
    oPC.OpenLocalDatabase
    
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM vCSL WHERE CSL_EXCH_ID='" & pEXCHID & "'"
    rs.Open strSQL, oPC.DBLocalConn, adOpenForwardOnly, adLockReadOnly
    Set oPBOUT = New PropertyBag
    i = 0
    Do While Not rs.EOF
        With udtProps
            .ID = FNN(rs.Fields("CSL_ID"))
            .CodeF = FNS(rs.Fields("CodeF"))
            .Code = FNS(rs.Fields("P_Code"))
            .PID = FNS(rs.Fields("CSL_P_ID"))
            .title = FNS(rs.Fields("P_Title"))
            .Author = FNS(rs.Fields("P_MainAuthor"))
            .Qty = FNN(rs.Fields("CSL_Qty"))
            .Price = FNN(rs.Fields("CSL_Price"))
            .PriceAlteration = FNN(rs.Fields("CSL_PriceAlteration"))
            .DiscountRate = FNDBL(rs.Fields("CSL_DiscountRate"))
            .Discount = FNN(rs.Fields("CSL_Discount"))
            .VATRate = FNDBL(rs.Fields("CSL_VATRATE"))
            .VAT = FNN(rs.Fields("CSL_VAT"))
            .DiscountDescription = FNS(rs.Fields("CSL_DiscountDescription"))
            .Counterfoil = FNS(rs.Fields("CSL_Counterfoil"))
            .COLID = FNN(rs.Fields("CSL_COLID"))
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
    
    oPC.CloseLocalDatabase
  
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale_P.Fetch(pEXCHID)", pEXCHID
End Function

Public Function FetchFromMainDB(ByVal pEXCHID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As SaleProps
Dim udtData As SaleData
Dim oPBOUT As PropertyBag
Dim lngCount As Long
Dim i As Long

    oPC.OpenLocalDatabase
    
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM vCSL_ForReprint WHERE CSL_EXCH_ID='" & pEXCHID & "'"
    rs.Open strSQL, oPC.DBMainConn, adOpenForwardOnly, adLockReadOnly
    Set oPBOUT = New PropertyBag
    i = 0
    Do While Not rs.EOF
        With udtProps
            .ID = FNN(rs.Fields("CSL_ID"))
            .CodeF = FNS(rs.Fields("CodeF"))
            .Code = FNS(rs.Fields("P_Code"))
            .PID = FNS(rs.Fields("CSL_P_ID"))
            .title = FNS(rs.Fields("P_Title"))
            .Author = FNS(rs.Fields("P_MainAuthor"))
            .Qty = FNN(rs.Fields("CSL_Qty"))
            .Price = FNN(rs.Fields("CSL_Price"))
            .PriceAlteration = FNN(rs.Fields("CSL_PriceAlteration"))
            .DiscountRate = FNDBL(rs.Fields("CSL_DiscountRate"))
            .Discount = FNN(rs.Fields("CSL_Discount"))
            .VATRate = FNDBL(rs.Fields("CSL_VATRATE"))
            .VAT = FNN(rs.Fields("CSL_VATRATE"))
            .DiscountDescription = FNS(rs.Fields("CSL_DiscountDescription"))
            .Counterfoil = FNS(rs.Fields("CSL_Counterfoil"))
            .COLID = FNN(rs.Fields("CSL_COLID"))
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
          FetchFromMainDB = .Contents
    End With
    Set oPBOUT = Nothing
    
    oPC.CloseLocalDatabase
  
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale_P.FetchFromMainDB(pEXCHID)", pEXCHID
End Function


Public Function Save(ByVal buffer As String, ByVal pEXCHID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As SaleProps
Dim udtData As SaleData

Dim oPBIN As PropertyBag
Dim oPBOUT As PropertyBag

Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
Dim strPos As String

    oPC.OpenLocalDatabase
    
  Set oPBIN = New PropertyBag
  Set oPBOUT = New PropertyBag
  Set rs = New ADODB.Recordset
  
  arBuffer = buffer
  oPBIN.Contents = arBuffer
  
  strPos = "Pos 1"
  For lngIndex = 1 To oPBIN.ReadProperty("Count")
    udtData.buffer = oPBIN.ReadProperty("Item" & CStr(lngIndex))
    LSet udtProps = udtData
  strPos = "Pos 2"
    If Not udtProps.IsDeleted Then
      strSQL = "SELECT * FROM tCSL WHERE CSL_ID=" & udtProps.ID
      rs.Open strSQL, oPC.DBLocalConn, adOpenKeyset, adLockOptimistic
      If udtProps.IsNew Then rs.AddNew
  strPos = "Pos 3"
      With udtProps
        rs.Fields("CSL_EXCH_ID") = FNS(pEXCHID)
        If .PID > "0" Then
            rs.Fields("CSL_P_ID") = FNS(.PID)
        End If
  strPos = "Pos 4"
        rs.Fields("CSL_Qty") = FNN(.Qty)
        rs.Fields("CSL_Price") = FNN(.Price)
        rs.Fields("CSL_PriceAlteration") = FNN(.PriceAlteration)
        rs.Fields("CSL_DiscountRate") = FNDBL(.DiscountRate)
        rs.Fields("CSL_Discount") = FNN(.Discount)
        rs.Fields("CSL_VATRate") = FNDBL(.VATRate)
        rs.Fields("CSL_VAT") = FNN(.VAT)
  strPos = "Pos 5"
        rs.Fields("CSL_Counterfoil") = FNS(.Counterfoil)
        rs.Fields("CSL_DiscountDescription") = FNS(.DiscountRule)
        rs.Fields("CSL_COLID") = FNN(.COLID)
  strPos = "Pos 6"
        rs.Update
  strPos = "Pos 7"
        .IsNew = False
        .IsDirty = False
      End With
      
      LSet udtData = udtProps
      lngCount = lngCount + 1
  strPos = "Pos 8"
      oPBOUT.WriteProperty "Item" & CStr(lngCount), udtData.buffer
        rs.Close
    Else
      DeleteCSL udtProps.ID
    End If
  Next
  oPBOUT.WriteProperty "Count", lngCount
  
  Set oPBIN = Nothing
  Set rs = Nothing
  
  Save = oPBOUT.Contents
  Set oPBOUT = Nothing
  
  oPC.CloseLocalDatabase
  
  Exit Function
  
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale_P.Save(buffer,pEXCHID)", Array(buffer, pEXCHID), , , "strPos", Array(strPos)
End Function

Private Sub DeleteCSL(pCSLID As Long)
    On Error GoTo errHandler
    oPC.OpenLocalDatabase
    
    oPC.DBLocalConn.Execute "DELETE FROM tCSL WHERE CSL_ID=" & pCSLID
    
    oPC.CloseLocalDatabase
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale_P.DeleteCSL(pCSLID)", pCSLID
End Sub

Public Sub DeleteObject(ByVal pEXCHID As String)
    On Error GoTo errHandler
    oPC.OpenLocalDatabase
    
    oPC.DBLocalConn.Execute "DELETE FROM tCSL WHERE CSL_EXCH_ID=" & pEXCHID
    
    oPC.CloseLocalDatabase

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Sale_P.DeleteObject(pEXCHID)", pEXCHID
End Sub


