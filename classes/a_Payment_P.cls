VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Payment_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Option Explicit

Public Function Fetch(ByVal pEXCHID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As PaymentProps
Dim udtData As PaymentData
Dim oPBOUT As PropertyBag
Dim lngCount As Long
    
    oPC.OpenLocalDatabase
    
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tPayment WHERE PAY_EXCH_ID='" & pEXCHID & "'"
    rs.Open strSQL, oPC.DBLocalConn, adOpenForwardOnly, adLockReadOnly
    Set oPBOUT = New PropertyBag
    Do While Not rs.EOF
        With udtProps
            .PaymentID = FNN(rs.Fields("PAY_ID"))
            .EXCHANGE_GUID = FNS(rs.Fields("PAY_EXCH_ID"))
            .Amt = FNN(rs.Fields("PAY_Amt"))
            .AmtTendered = FNN(rs.Fields("PAY_Amt_Tendered"))
            .Type = FNS(rs.Fields("PAY_PaymentType"))
            .COLID = FNN(rs.Fields("PAY_COLID"))
            .Reference = FNS(rs.Fields("PAY_Ref"))
            .Note = FNS(rs.Fields("PAY_Note"))
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
    ErrorIn "a_Payment_P.Fetch(pEXCHID)", pEXCHID
End Function

Public Function FetchFromMainDB(ByVal pEXCHID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As PaymentProps
Dim udtData As PaymentData
Dim oPBOUT As PropertyBag
Dim lngCount As Long

    oPC.OpenLocalDatabase
    
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM vPayment_ForReprint WHERE PAY_EXCH_ID='" & pEXCHID & "'"
    rs.Open strSQL, oPC.DBMainConn, adOpenForwardOnly, adLockReadOnly
    Set oPBOUT = New PropertyBag
    Do While Not rs.EOF
        With udtProps
            .PaymentID = FNN(rs.Fields("PAY_ID"))
            .EXCHANGE_GUID = FNS(rs.Fields("PAY_EXCH_ID"))
            .Amt = FNN(rs.Fields("PAY_Amt"))
            .AmtTendered = FNN(rs.Fields("PAY_Amt_Tendered"))
            .Type = FNS(rs.Fields("PAY_PaymentType"))
            .COLID = FNN(rs.Fields("PAY_COLID"))
            .Reference = FNS(rs.Fields("PAY_Ref"))
            .Note = FNS(rs.Fields("PAY_Note"))
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
    ErrorIn "a_Payment_P.FetchFromMainDB(pEXCHID)", pEXCHID
End Function



Public Function Save(ByVal buffer As String, ByVal pEXCHID As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As PaymentProps
Dim udtData As PaymentData
Dim oPBIN As PropertyBag
Dim oPBOUT As PropertyBag

Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
  
  oPC.OpenLocalDatabase
  
  Set oPBIN = New PropertyBag
  Set oPBOUT = New PropertyBag
  Set rs = New ADODB.Recordset
  
  arBuffer = buffer
  oPBIN.Contents = arBuffer
  
  
  For lngIndex = 1 To oPBIN.ReadProperty("Count")
        With oPBIN
            udtData.buffer = oPBIN.ReadProperty("Item" & CStr(lngIndex))
            LSet udtProps = udtData
        End With
        If Not udtProps.IsDeleted Then
            strSQL = "SELECT * FROM tPayment WHERE Pay_ID=" & udtProps.PaymentID
            rs.Open strSQL, oPC.DBLocalConn, adOpenKeyset, adLockOptimistic
            If udtProps.IsNew Then
                rs.AddNew
            End If
            With udtProps
                rs.Fields("PAY_EXCH_ID") = pEXCHID
                rs.Fields("PAY_COLID") = FNN(.COLID)
                rs.Fields("PAY_Amt") = FNN(.Amt)
                rs.Fields("PAY_Amt_Tendered") = FNN(.AmtTendered)
                rs.Fields("PAY_PaymentType") = FNS(.Type)
                rs.Fields("PAY_REF") = FNS(.Reference)
                rs.Fields("PAY_NOTE") = FNS(.Note)
                rs.Update
                If .IsNew Then
                  rs.Bookmark = rs.Bookmark
                  .PaymentID = rs("PAY_ID")
                End If
                .IsNew = False
                .IsDirty = False
            End With
          
            LSet udtData = udtProps
            lngCount = lngCount + 1
            oPBOUT.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.Close
        Else
          DeletePayment udtProps.PaymentID
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
    ErrorIn "a_Payment_P.Save(buffer,pEXCHID)", Array(buffer, pEXCHID)
End Function

Private Sub DeletePayment(pPAYID As Long)
    On Error GoTo errHandler
    oPC.OpenLocalDatabase
    
    oPC.DBLocalConn.Execute "DELETE FROM tPayment WHERE PAY_ID=" & pPAYID
    
    oPC.CloseLocalDatabase
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Payment_P.DeleteCSL(pPAYID)", pPAYID
End Sub

Public Sub DeleteObject(ByVal pEXCHID As String)
    On Error GoTo errHandler
    oPC.OpenLocalDatabase
    
    oPC.DBLocalConn.Execute "DELETE FROM tPayment WHERE PAY_EXCH_ID=" & pEXCHID
    
    oPC.CloseLocalDatabase
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Payment_P.DeleteObject(pEXCHID)", pEXCHID
End Sub


