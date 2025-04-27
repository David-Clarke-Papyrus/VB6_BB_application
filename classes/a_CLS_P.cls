VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_CSL_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(pCSLID As Long) As String
    On Error GoTo errHandler
Dim udtProps As CSLProps
Dim udtData As CSLData
Dim oPBOUT As PropertyBag
Dim lngCount As Long
'Dim oProdCode As z_ProdCode
Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    rs.open "SELECT * FROM vCashSales  WHERE CSL_TR_ID = " & pCSLID, oPC.COShort, adOpenKeyset
    Set oPBOUT = New PropertyBag
    Do While Not rs.eof
        With rs
            udtProps.ID = .fields("CSL_ID") 'this seems to be the only way of referencing the Primary key
            udtProps.PID = FNS(.fields("CSL_P_ID"))
            udtProps.EAN = FNS(.fields("P_EAN"))
            udtProps.Qty = FNN(.fields("CSL_Qty"))
            udtProps.Author = FNS(.fields("P_MainAuthor"))
            udtProps.Title = FNS(.fields("P_Title"))
            udtProps.Price = FNN(.fields("CSL_PRICE"))
            udtProps.DateTime = FND(.fields("CSL_TimeOfSale"))
            udtProps.Discount = FNN(.fields("CSL_Discount"))
            udtProps.VATRate = FNN(.fields("CSL_VATRate"))
            udtProps.TRID = .fields("CSL_TR_ID")
            udtProps.CodeFForExport = FormatProductCode(FNS(rs.fields("P_Code")), True)
            udtProps.CodeF = FormatProductCode(FNS(rs.fields("P_Code")), False)
            udtProps.code = FNS(rs.fields("P_Code"))
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

EXIT_Handler:
    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL_P.Fetch(pCSLID)", pCSLID
End Function
Public Function Save(ByVal buffer As String, ByVal TRID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As CSLProps
Dim udtData As CSLData
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
            strSQL = "SELECT * FROM tCSL WHERE CSL_ID=" & CStr(udtProps.ID)
            rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then
                rs.AddNew
            End If
          
            With udtProps
                    rs!CSL_TR_ID = TRID
                    If FNS(.PID) > "" Then
                        rs!CSL_P_ID = FNS(.PID)
                    Else
                        rs!CSL_Code = FNS(.CodeF)
                    End If
                    rs!CSL_Qty = .Qty
                    rs!CSL_Price = .Price
                    rs!CSL_TimeOfSale = .DateTime
                    rs.Update
                  
                    If .IsNew Then
                        rs.Bookmark = rs.Bookmark
                        .ID = rs("CSL_ID")
                    End If
                    .IsNew = False
                    .IsDirty = False
            End With
      LSet udtData = udtProps
      lngCount = lngCount + 1
      objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
      rs.Close
    Else
      DeleteCSLine udtProps.ID
    End If
  Next
  objPBOut.WriteProperty "Count", lngCount
  
  Set objPB = Nothing
  Set rs = Nothing
  
  Save = objPBOut.Contents
  Set objPBOut = Nothing
  Exit Function
  
EXIT_Handler:
'H:
'    oor.Setor , "a_CSL_P:Save:" & or, Now(), "", "", ""
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL_P.Save(buffer,TRID)", Array(buffer, TRID)
End Function
Public Sub DeleteObject(ByVal TRID As Long)
    On Error GoTo errHandler
Dim strSQL As String
  
    strSQL = "DELETE FROM tCSL WHERE CSL_TR_ID=" & TRID
    oPC.COShort.execute strSQL
    Exit Sub

'H:
'  Err.Raise err.Number
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL_P.DeleteObject(TRID)", TRID
End Sub

Public Sub DeleteCSLine(CSLID As Long)
    On Error GoTo errHandler
Dim strSQL As String
    strSQL = "DELETE FROM tCSL WHERE CSL_ID=" & CSLID
    oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CSL_P.DeleteCSLine(CSLID)", CSLID
End Sub

