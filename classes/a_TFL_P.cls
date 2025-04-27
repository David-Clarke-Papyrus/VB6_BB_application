VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_TFL_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(pTFLID As Long) As String
    On Error GoTo errHandler
Dim udtProps As TFLProps
Dim udtData As TFLData
Dim oPBOUT As PropertyBag
Dim lngCount As Long
Dim rs As ADODB.Recordset
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open "SELECT tTFRL.*,tProduct.*,dbo.CODEF(P_CODE,P_EAN,0) as CODEF,dbo.CODEF(P_CODE,P_EAN,1) as CODEFEX FROM tTFRL JOIN tPRODUCT On TFRL_P_ID = P_ID WHERE TFRL_TR_ID = " & pTFLID, oPC.COShort, adOpenKeyset
    Set oPBOUT = New PropertyBag
    Do While Not rs.eof
        With rs
            udtProps.ID = .fields("TFRL_ID") 'this seems to be the only way of referencing the Primary key
            udtProps.PID = FNS(.fields("TFRL_P_ID"))
            udtProps.EAN = FNS(.fields("P_EAN"))
            udtProps.Qty = FNN(.fields("TFRL_Qty"))
            udtProps.Author = FNS(.fields("P_MainAuthor"))
            udtProps.Title = FNS(.fields("P_Title"))
            udtProps.Note = FNS(.fields("TFRL_Note"))
            udtProps.Price = FNN(.fields("TFRL_Price"))
            udtProps.Binding = FNS(.fields("P_BindingCode"))
            udtProps.SeriesTitle = FNS(.fields("P_SeriesTitle"))
            If FNN(udtProps.Cost) = 0 Then
                udtProps.Cost = FNN(udtProps.Price) * (100 - FNDBL(oPC.GetProperty("DiscountToCalculateDefaultCostForTransfers"))) / 100
            End If
            udtProps.TRID = .fields("TFRL_TR_ID")
            udtProps.VATRate = FNDBL(.fields("TFRL_VatRate"))
            udtProps.Discount = FNDBL(.fields("TFRL_Discount"))
            udtProps.CodeFForExport = FNS(.fields("CODEFEX"))
            udtProps.CodeF = FNS(.fields("CODEF"))
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
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL_P.Fetch(pTFLID)", pTFLID
End Function
Public Function Save(ByVal buffer As String, ByVal TRID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As TFLProps
Dim udtData As TFLData
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
Dim OpenResult As Integer
    
    Set objPB = New PropertyBag
    Set objPBOut = New PropertyBag
    arBuffer = buffer
    objPB.Contents = arBuffer
    Set rs = New ADODB.Recordset
    '--------------
    OpenResult = oPC.OpenDBSHort
    '--------------
    For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData
      
        If Not udtProps.IsDeleted Then
            strSQL = "SELECT * FROM tTFRL WHERE TFRL_ID=" & CStr(udtProps.ID)
            rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
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
                    rs!TFRL_Note = FNS(.Note)
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
    '--------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
    '--------------
  objPBOut.WriteProperty "Count", lngCount
  
  Set objPB = Nothing
  Set rs = Nothing
  
  Save = objPBOut.Contents
  Set objPBOut = Nothing
  Exit Function
  
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL_P.Save(buffer,TRID)", Array(buffer, TRID)
End Function
Public Sub DeleteObject(ByVal TRID As Long)
    On Error GoTo errHandler
Dim strSQL As String
  
    strSQL = "DELETE FROM tTFRL WHERE TFRL_TR_ID=" & TRID
    oPC.COShort.execute strSQL
    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL_P.DeleteObject(TRID)", TRID
End Sub

Public Sub DeleteTFLine(TFLID As Long)
    On Error GoTo errHandler
Dim strSQL As String
    strSQL = "DELETE FROM tTFRL WHERE TFRL_ID=" & TFLID
    oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_TFL_P.DeleteTFLine(TFLID)", TFLID
End Sub


