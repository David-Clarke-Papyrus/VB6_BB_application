VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_RET_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Fetch(pTFLID As Long) As String
On Error GoTo ERRH
Dim udtProps As TFLProps
Dim udtData As TFLData
Dim oPBOUT As PropertyBag
Dim lngCount As Long
Dim oProdcode As z_ProdCode
Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    rs.Open "SELECT tTFRL.*,tProduct.* FROM tTFRL JOIN tPRODUCT On TFRL_P_ID = P_ID WHERE TFRL_TR_ID = " & pTFLID, oPC.CO, adOpenKeyset
    Set oPBOUT = New PropertyBag
    On Error GoTo ERRH
    Do While Not rs.EOF
        With rs
            udtProps.ID = .Fields("TFRL_ID") 'this seems to be the only way of referencing the Primary key
            udtProps.PID = FNS(.Fields("TFRL_P_ID"))
            udtProps.EAN = FNS(.Fields("P_EAN"))
            udtProps.Qty = FNN(.Fields("TFRL_Qty"))
            udtProps.Author = FNS(.Fields("P_MainAuthor"))
            udtProps.Title = FNS(.Fields("P_Title"))
            udtProps.Price = FNN(.Fields("TFRL_Price"))
            udtProps.TRID = .Fields("TFRL_TR_ID")
            udtProps.VATRate = FNDBL(.Fields("TFRL_VatRate"))
            udtProps.Discount = FNDBL(.Fields("TFRL_Discount"))
            Set oProdcode = New z_ProdCode
            oProdcode.Load rs.Fields("P_Code")
            udtProps.CodeFForExport = oProdcode.FormatCode(True)
            udtProps.CodeF = oProdcode.FormatCode(False)
            udtProps.Code = FNS(rs.Fields("P_Code"))
            Set oProdcode = Nothing
            udtProps.IsNew = False
            udtProps.IsDirty = False
            udtProps.IsDeleted = False
        End With
        LSet udtData = udtProps
        lngCount = lngCount + 1
        oPBOUT.WriteProperty "Item" & CStr(lngCount), udtData.buffer
        rs.MoveNext
    Loop
    Set rs = Nothing
    With oPBOUT
          .WriteProperty "Count", CStr(lngCount)
          Fetch = .Contents
    End With
    Set oPBOUT = Nothing
    Exit Function

EXIT_Handler:
    Exit Function
ERRH:
    oError.SetError Err, "a_CSL_P:Fetch:" & Error, Now(), "a_CSL", "Fetch", "unknown"
    GoTo EXIT_Handler
    Resume
End Function
Public Function Save(ByVal buffer As String, ByVal TRID As Long) As String
On Error GoTo ERRH
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As TFLProps
Dim udtData As TFLData
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
    
    On Error GoTo ERRH
    For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData
      
        If Not udtProps.IsDeleted Then
            strSQL = "SELECT * FROM tTFRL WHERE TFRL_ID=" & CStr(udtProps.ID)
            rs.Open strSQL, oPC.CO, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then
                rs.AddNew
            End If
          
            With udtProps
                    rs!TFRL_TR_ID = TRID
                    rs!TFRL_P_ID = FNS(.PID)
                    rs!TFRL_Qty = .Qty
                    rs!TFRL_Price = .Price
                    rs!TFRL_Discount = .Discount
                    rs!TFRL_VATRATE = .VATRate
                    rs.Update
                  
                    If .IsNew Then
                        rs.Bookmark = rs.Bookmark
                        .ID = rs("TFRL_ID")
                    End If
                    .IsNew = False
                    .IsDirty = False
            End With
      LSet udtData = udtProps
      lngCount = lngCount + 1
      objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
      rs.Close
    Else
      DeleteTFLine udtProps.ID
    End If
  Next
  objPBOut.WriteProperty "Count", lngCount
  
  Set objPB = Nothing
  Set rs = Nothing
  
  Save = objPBOut.Contents
  Set objPBOut = Nothing
  Exit Function
  
EXIT_Handler:
    Exit Function
ERRH:
    oError.SetError Err, "a_CSL_P:Save:" & Error, Now(), "", "", ""
    Exit Function
    Resume
End Function
Public Sub DeleteObject(ByVal TRID As Long)
Dim strSQL As String
  
    strSQL = "DELETE FROM tTFL WHERE TFL_TR_ID=" & TRID
    oPC.CO.execute strSQL
    Exit Sub

ERRH:
  Err.Raise Err.Number
End Sub

Public Sub DeleteTFLine(TFLID As Long)
Dim strSQL As String
    strSQL = "DELETE FROM tTFL WHERE TFL_ID=" & TFLID
    oPC.CO.execute strSQL
End Sub



