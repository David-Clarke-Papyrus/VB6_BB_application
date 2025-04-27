VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_DeliveryLine_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(ByVal ID As Long) As String
    On Error GoTo errHandler
Dim rsDL As ADODB.Recordset
Dim strSQL As String
Dim udtProps As DELLProps
Dim udtData As DELLData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim oProdCode As z_ProdCode

    Set rsDL = New ADODB.Recordset
    strSQL = "SELECT TOP (100) PERCENT * FROM DELL_Fetch WHERE DELL_TR_ID =" & ID & " ORDER BY DELL_ID DESC"
    rsDL.open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
    Set objPB = New PropertyBag
    Do While Not rsDL.eof
        With udtProps
            .PID = FNS(rsDL.fields("P_ID"))
            .QtyFirm = FNN(rsDL.fields("DELL_QtyFirm"))
            .Note = FNS(rsDL.fields("DELL_Note"))
            .QtySS = FNN(rsDL.fields("DELL_QtySS"))
            .QtyTotal = FNN(rsDL.fields("DELL_QtyTotal"))
            .QtyShort = FNN(rsDL.fields("DELL_QtyShort"))
            .Discount = FNDBL(rsDL.fields("DELL_Discount"))
            .PriceSell = FNN(rsDL.fields("DELL_PriceSell"))
            .Price = FNN(rsDL.fields("DELL_Price"))
            .ForeignPrice = FNN(rsDL.fields("DELL_ForeignPrice"))
            .Title = FNS(rsDL.fields("P_Title"))
            .POLID = FNN(rsDL.fields("DELL_POL_ID"))
            .COLID = FNN(rsDL.fields("DELL_COL_ID"))
            .ClaimID = FNN(rsDL.fields("DELL_ClaimID"))
            .code = FNS(rsDL.fields("P_Code"))
            .Ref = FNS(rsDL.fields("COL_REF"))
            .DELLID = FNN(rsDL.fields("DELL_ID"))
            .PO_Code = FNS(rsDL.fields("TR_Code"))
            .POL_Discount = FNDBL(rsDL.fields("POL_Discount"))
            .POL_Price = FNN(rsDL.fields("POL_Price"))
            .POL_QtyFirm = FNN(rsDL.fields("POL_QTyFirm"))
            .POL_QtySS = FNN(rsDL.fields("POL_QTySS"))
            .PO_CURRID = FNN(rsDL.fields("PO_CURR_ID"))
            .PO_CURRRATE = FNN(rsDL.fields("PO_CURR_ID"))
            .POL_ForeignPrice = FNN(rsDL.fields("POL_ForeignPrice"))
            .CodeFForExport = FNS(rsDL.fields("CODEFEX"))
            .CodeF = FNS(rsDL.fields("CODEF"))
            .ReasonID = FNS(rsDL.fields("DELL_ReasonID"))
            .CorrectedDiscount = FNDBL(rsDL.fields("DELL_CorrectedDiscount"))
            .CorrectedPrice = FNDBL(rsDL.fields("DELL_CorrectedPrice"))
            .CorrectedForeignPrice = FNDBL(rsDL.fields("DELL_CorrectedForeignPrice"))
            .EAN = FNS(rsDL.fields("P_EAN"))
            .MBCode = FNS(rsDL.fields("DELL_MBCode"))
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
        End With
        LSet udtData = udtProps
        lngCount = lngCount + 1
        objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
        rsDL.MoveNext
    Loop
    rsDL.Close
    Set rsDL = Nothing
    With objPB
          .WriteProperty "Count", CStr(lngCount)
          Fetch = .Contents
    End With
    Set objPB = Nothing
    Exit Function
errHandler:
    RlsObjs rsDL
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine_P.Fetch(ID)", ID
End Function
Private Function Inspect(pIn) As Variant
    On Error GoTo errHandler
Dim i As Integer
Dim str As String
    For i = 0 To UBound(pIn)
        str = str & "  " & pIn(i)
    Next
    Inspect = str
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine_P.Inspect(pIn)", pIn
End Function
Private Sub Ins(pIn)
    On Error GoTo errHandler
Dim i As Integer
    For i = 0 To UBound(pIn)
        Debug.Print pIn(i)
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine_P.Ins(pIn)", pIn
End Sub
Public Function Save(ByVal buffer As String, ByVal ID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As DELLProps
Dim udtData As DELLData
Dim oPBIN As PropertyBag
Dim oPBOUT As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long

  Set oPBIN = New PropertyBag
  Set oPBOUT = New PropertyBag
  Set rs = New ADODB.Recordset
  
  arBuffer = buffer
  oPBIN.Contents = arBuffer
  
  For lngIndex = 1 To oPBIN.ReadProperty("Count")
    udtData.buffer = oPBIN.ReadProperty("Item" & CStr(lngIndex))
    LSet udtProps = udtData
    If Not udtProps.IsDeleted Then
        strSQL = "SELECT * FROM tDELL WHERE DELL_ID=" & udtProps.DELLID
        rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
        If udtProps.IsNew Then rs.AddNew
        With udtProps
          rs.fields("DELL_TR_ID") = ID
          rs.fields("DELL_POL_ID") = FNN(.POLID)
          rs.fields("DELL_QtySS") = FNN(.QtySS)
          rs.fields("DELL_QtyFirm") = FNN(.QtyFirm)
          rs.fields("DELL_QtyShort") = FNN(.QtyShort)
          rs.fields("DELL_QtyTotal") = FNN(.QtySS) + FNN(.QtyFirm)
          rs.fields("DELL_Price") = FNN(.Price)
          rs.fields("DELL_ForeignPrice") = FNN(.ForeignPrice)
          rs.fields("DELL_PriceSell") = FNN(.PriceSell)
          rs.fields("DELL_Discount") = FNDBL(Round(.Discount, 2))
          rs.fields("DELL_P_ID") = FNS(.PID)
          rs.fields("DELL_COL_ID") = FNN(.COLID)
          rs.fields("DELL_ClaimID") = FNN(.ClaimID)

          rs.fields("DELL_Note") = FNS(.Note)
          rs.fields("DELL_ReasonID") = FNS(.ReasonID)
         ' MsgBox udtProps.IsDirty
          rs.fields("DELL_MBCode") = FNS(.MBCode)
          rs.fields("DELL_ProductType_ID") = FNN(.ProductTypeID)
          rs.fields("DELL_CorrectedDiscount") = FNDBL(.CorrectedDiscount)
          rs.fields("DELL_CorrectedPrice") = FNDBL(.CorrectedPrice)
          rs.fields("DELL_CorrectedForeignPrice") = FNDBL(.CorrectedForeignPrice)
          rs.Update
          
          If .IsNew Then
            rs.Bookmark = rs.Bookmark
            .DELLID = rs("DELL_ID")
          End If
          .IsNew = False
          .IsDirty = False
        End With
        
        LSet udtData = udtProps
        lngCount = lngCount + 1
        oPBOUT.WriteProperty "Item" & CStr(lngCount), udtData.buffer
        rs.Close
    Else
      DeleteDELL udtProps.DELLID
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
    If Err = -2147217873 Then ' duplicate row
        Set rs = Nothing
        Save = "!: Duplicate row in delivery"
        Set oPBIN = Nothing
        Set oPBOUT = Nothing
        Exit Function
    End If
    
    Set oPBIN = Nothing
    Set oPBOUT = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine_P.Save(buffer,ID)", Array(buffer, ID), , , "PID,ID,line", Array(FNS(udtProps.PID), CStr(ID), Erl())
End Function

Public Sub DeleteDELL(DELLID As Long)
    On Error GoTo errHandler
    oPC.COShort.execute "DELETE FROM tDELL WHERE DELL_ID=" & DELLID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine_P.DeleteDELL(DELLID)", DELLID
End Sub

Public Sub DeleteObject(ByVal DELID As Long)
    On Error GoTo errHandler

  oPC.COShort.execute "DELETE FROM tDELL WHERE DELL_TR_ID=" & DELID
  Exit Sub

'H:
'  Err.Raise err.Number
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_DeliveryLine_P.DeleteObject(DELID)", DELID
End Sub



