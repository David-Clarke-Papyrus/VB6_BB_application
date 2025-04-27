VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Invoice_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function Fetch(TransactionID As Long) As String
Dim rs As ADODB.Recordset
  
Dim strSQL As String
Dim udtProps As InvoiceProps
Dim udtData As InvoiceData
Dim objPersist As a_InvoiceLine_P
Dim objPB As PropertyBag
Dim oUtil As z_Batch

    Set rs = New ADODB.Recordset
    Set oUtil = New z_Batch
    lngResult = oUtil.RunGetRecordset("SELECT * FROM Invoice_Fetch WHERE I_ID = " & TransactionID, Text, Array(), "", rs)
    
    On Error GoTo ERRH
    
    If Not rs.EOF Then
        With udtProps
            .InvoiceID = rs!TR_ID
            .TPID = FixNullsNumber(rs!TR_TP_ID)
            .InvDocAddressID = FixNullsNumber(rs.Fields("AddressDOCID"))
            .InvGoodsAddressID = FixNullsNumber(rs.Fields("AddressGOODSID"))
            .CurrencyID_Foreign = FixNullsNumber(rs!I_CURR_ID)
            .CurrencyFactor = FixNullsDouble(rs!I_CURRRate)
            .VATRate = FixNullsDouble(rs!I_VATRate)
            .DiscountRate = FixNullsDouble(rs!I_DiscountRate)
            .TotalDiscount = FixNullsNumber(rs!I_DiscountAmount)
            .TotalVAT = FixNullsNumber(rs!I_VATAmount)
            .TotalPayable = FixNullsNumber(rs!I_PayableAmount)
            .CompanyID = FixNullsNumber(rs!TR_COMP_ID)
            .ShowVAT = FixNullsBoolean(rs!I_ShowVAT)
            .CurrencyFormat = FixNullsString(rs!CURR_Format)
            .BusPhone = FixNullsString(rs!ADD_BusPhone)
            .TPAccNum = FixNullsString(rs!TP_ACno)
            .TPFax = FixNullsString(rs!ADD_Fax)
            .TPPhone = FixNullsString(rs!TP_Phone)
            .TPName = FixNullsString(rs!TP_Name)
            .TPMemo = FixNullsString(rs!TR_Note)
            .Transcode = FixNullsString(rs!TR_Code)
            If rs!TR_Date > DateSerial(1995, 1, 1) Then .TransDate = DateSerial(Year(rs!TR_Date), Month(rs!TR_Date), Day(rs!TR_Date))
            .CaptureDate = FixNullsDate(rs!TR_CaptureDate)
            .Vatable = FixNullsBoolean(rs!I_Vatable)
            .ShowVAT = FixNullsBoolean(rs!I_ShowVAT)
            .Status = rs!TR_Status
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            Set oUtil = Nothing
            LSet udtData = udtProps
              Set objPersist = New a_InvoiceLine_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.Buffer
              .WriteProperty "ILS", objPersist.Fetch(udtProps.InvoiceID)
              Fetch = .Contents
            End With
            Set objPB = Nothing
            Set objPersist = Nothing
        End With
    Else
        ' force an error
        rs.MoveNext
    End If
    
  Exit Function
  
ERRH:
  rs.Close
  Set rs = Nothing
    Err.Raise Err.Number
 ' Resume
End Function

Public Function Save(ByVal Buffer As String) As String
Dim rs As ADODB.Recordset
'Dim cmd As New ADODB.Command
Dim strSQL As String
Dim objCode As New z_Code
Dim udtProps As InvoiceProps
Dim udtData As InvoiceData
Dim objPersist As a_InvoiceLine_P
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte

    Set objPB = New PropertyBag
    arBuffer = Buffer
    With objPB
      .Contents = arBuffer
      udtData.Buffer = .ReadProperty("State")
    End With
    LSet udtProps = udtData

    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tTR WHERE TR_ID =" & udtProps.InvoiceID
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockOptimistic
    On Error GoTo ERRH
    With udtProps
        If .IsNew Then rs.AddNew
        rs!TR_Type = 3
        rs!TR_TP_ID = FixNullsNumber(.TPID)
        rs!TR_Date = FixNullsDate(.TransDate)
        rs!TR_CaptureDate = FixNullsDate(.CaptureDate)
        rs!TR_Code = FixNullsString(.Transcode)
        rs!TR_Note = FixNullsString(.TPMemo)
        rs!TR_AddressDOCID = FixNullsNumber(.InvDocAddressID)
        rs!TR_AddressGoodsID = FixNullsNumber(.InvGoodsAddressID)
        rs!TR_Status = .Status
        rs!TR_COMP_ID = .CompanyID
        rs.Update
        If .IsNew Then
           rs.MoveLast
           .InvoiceID = rs!TR_ID
        End If
        rs.Close
    End With

    strSQL = "SELECT * FROM tInvoice WHERE I_ID =" & udtProps.InvoiceID
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockOptimistic
    
    With udtProps
        If udtProps.IsNew Then rs.AddNew
        rs!I_ID = .InvoiceID
        rs!I_CURR_ID = .CurrencyID_Foreign
      '  rs!I_Postage = .Postage
      '  rs!I_Insurance = .Insurance
        rs!I_CURRRate = .CurrencyFactor  '   Conversion rate
        rs!I_DiscountRate = .DiscountRate  '   Invoice discount (general) rate
        rs!I_VATRate = .VATRate  '
        rs!I_Vatable = .Vatable
        rs!I_DiscountAmount = .TotalDiscount
        rs!I_VATAmount = .TotalVAT
        rs!I_PayableAmount = .TotalPayable
        rs!I_ShowVAT = .ShowVAT
        rs.Update
        .IsNew = False
        .IsDirty = False
    End With
    rs.Close
    Set rs = Nothing
    
    Set objPBOut = New PropertyBag
    LSet udtData = udtProps
    Set objPersist = New a_InvoiceLine_P
      With objPBOut
        .WriteProperty "State", udtData.Buffer
        .WriteProperty "ILS", _
        objPersist.Save(objPB.ReadProperty("ILS"), udtProps.InvoiceID)
      End With
      Set objPB = Nothing
      Set objPersist = Nothing
    
      Save = objPBOut.Contents
      Set objPBOut = Nothing
      
      Exit Function

ERRH:
    rs.Close
    Set rs = Nothing
    Err.Raise Err.Number
End Function


Private Sub DeleteInvoice(ID As Long)
    oPC.CO.execute "DELETE FROM tPI WHERE PI_ID=" & ID

End Sub

Public Sub DeleteObject(ByVal ID As Long)
  Dim objPersist As a_InvoiceLine_P
  
    oPC.CO.BeginTrans
    On Error GoTo ERRH
    oPC.CO.execute "DELETE FROM tInvoice WHERE ID=" & ID
  
    Set objPersist = New a_InvoiceLine_P
    objPersist.DeleteObject ID
    Set objPersist = Nothing
    oPC.CO.CommitTrans
    Exit Sub
  
ERRH:
    oPC.CO.RollbackTrans
    Err.Raise Err.Number
End Sub

