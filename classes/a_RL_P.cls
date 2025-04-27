VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_RL_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(pRLID As Long) As String
    On Error GoTo errHandler
Dim udtProps As RLProps
Dim udtData As RLData
Dim oPBOUT As PropertyBag
Dim lngCount As Long
Dim oProdCode As z_ProdCode
Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    rs.open "SELECT dbo.tDEL.DEL_SupplierInvoiceDate,tRL.*,dbo.PubCode(P_Code) as Pubcode,tProduct.*,dbo.FlattenSections(dbo.tProduct.P_ID) as FlattenedSections, dbo.CODEF(P_CODE,P_EAN,0) AS CODEF,dbo.CODEF(P_CODE,P_EAN,1) AS CODEFEX FROM tRL JOIN tPRODUCT On RL_P_ID = P_ID LEFT OUTER JOIN tDELL ON RL_DELL_ID = DELL_ID LEFT OUTER JOIN tTR ON DELL_TR_ID = TR_ID LEFT OUTER JOIN tDEL ON TR_ID = DEL_ID WHERE RL_TR_ID = " & pRLID & " ORDER BY P_TITLE", oPC.COShort, adOpenKeyset
    Set oPBOUT = New PropertyBag
    Do While Not rs.eof
        With rs
            udtProps.RLID = .fields("RL_ID") 'this seems to be the only way of referencing the Primary key
            udtProps.PID = FNS(.fields("RL_P_ID"))
            udtProps.TRID = .fields("RL_TR_ID")
            udtProps.QtyRequested = FNN(.fields("RL_QtyRequested"))
            udtProps.QtyApproved = FNN(.fields("RL_QtyApproved"))
            udtProps.QtyReturned = FNN(.fields("RL_QtyREturned"))
            udtProps.QtyRejected = FNN(.fields("RL_QtyRejected"))
            udtProps.MainAuthor = FNS(.fields("P_MainAuthor"))
            udtProps.Title = FNS(.fields("P_Title"))
            udtProps.Price = FNN(.fields("RL_Price"))
            udtProps.ForeignPrice = FNN(.fields("RL_ForeignPrice"))
            udtProps.Note = FNS(.fields("RL_MEMO"))
            udtProps.Discount = Round(FNDBL(.fields("RL_Discount")), 2)
            udtProps.Status = FNS(.fields("RL_Status"))
            udtProps.Pubcode = FNS(.fields("Pubcode"))
            udtProps.DELLID = FNN(.fields("RL_DELL_ID"))
            udtProps.Section = FNS(.fields("FlattenedSections"))
          '  Set oProdCode = New z_ProdCode
          '  oProdCode.Load rs.Fields("P_Code")
            udtProps.CodeF = FNS(rs.fields("CODEF"))
            udtProps.EAN = FNS(rs.fields("P_EAN"))
            udtProps.SINVRef = FNS(rs.fields("RL_SIREF"))  ' & "   " & FND(rs.Fields("RL_SINVDATE"))
            udtProps.SINVDate = FND(rs!DEL_SUpplierInvoiceDate)
            
            udtProps.code = FNS(rs.fields("P_Code"))
          '  Set oProdCode = Nothing
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
'H:
'    oor.Setor , "a_RL_P:Fetch:" & or, Now(), "a_CSL", "Fetch", "unknown"
'    GoTo EXIT_Handler
'    Resume
    Exit Function
errHandler:
    ErrPreserve
    Set oPBOUT = Nothing

    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL_P.Fetch(pRLID)", pRLID
End Function
Public Function Save(ByVal buffer As String, ByVal TRID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As RLProps
Dim udtData As RLData
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
    
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    Set objPB = New PropertyBag
    Set objPBOut = New PropertyBag
    arBuffer = buffer
    objPB.Contents = arBuffer
    Set rs = New ADODB.Recordset
    
    For lngIndex = 1 To objPB.ReadProperty("Count")
        udtData.buffer = objPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData
      
        If Not udtProps.IsDeleted Then
            strSQL = "SELECT * FROM tRL WHERE RL_ID=" & udtProps.RLID
            rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then
                rs.AddNew
            End If
          
            With udtProps
                    rs!RL_TR_ID = TRID
                    rs!RL_P_ID = FNS(.PID)
                    rs!RL_QtyRequested = FNN(.QtyRequested)
                    rs!RL_QtyApproved = FNN(.QtyApproved)
                    rs!RL_QtyReturned = FNN(.QtyReturned)
                    rs!RL_Price = FNN(.Price)
                    rs!RL_ForeignPrice = FNN(.ForeignPrice)
                    rs!RL_Discount = Round(FNDBL(.Discount), 2)
                    rs!RL_Memo = FNS(.Note)
                    rs!RL_SiREF = FNS(.SINVRef)
                    rs!RL_DELL_ID = FNN(.DELLID)
                    rs!RL_Status = FNS(.Status)
                    rs.Update
                  
                    If .IsNew Then
                        rs.Bookmark = rs.Bookmark
                        .RLID = rs("RL_ID")
                    End If
                    .IsNew = False
                    .IsDirty = False
            End With
      LSet udtData = udtProps
      lngCount = lngCount + 1
      objPBOut.WriteProperty "Item" & CStr(lngCount), udtData.buffer
      rs.Close
    Else
      DeleteRLine udtProps.RLID
    End If
  Next
  objPBOut.WriteProperty "Count", lngCount
  
  Set objPB = Nothing
  Set rs = Nothing
  
  Save = objPBOut.Contents
  Set objPBOut = Nothing
  
EXIT_Handler:
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    ErrPreserve
    Set objPBOut = Nothing
    Set objPB = Nothing
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL_P.Save(buffer,TRID)", Array(buffer, TRID)
End Function
Public Sub DeleteObject(ByVal TRID As Long)
    On Error GoTo errHandler
Dim strSQL As String
  
    strSQL = "DELETE FROM tRL WHERE RL_TR_ID=" & TRID
    oPC.COShort.execute strSQL
    Exit Sub

errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL_P.DeleteObject(TRID)", TRID
End Sub

Public Sub DeleteRLine(RLID As Long)
    On Error GoTo errHandler
Dim strSQL As String
    strSQL = "DELETE FROM tRL WHERE RL_ID=" & RLID
    oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_RL_P.DeleteRLine(RLID)", RLID
End Sub



