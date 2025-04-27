VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSaleLineCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim rs As New ADODB.Recordset

Dim sServerPath As String
Dim sCompID As String

Dim iExchange As Integer
Dim iSaleLine As Integer
Dim iPayment As Integer

Private Sub Class_Initialize()
    CreateRecordset
End Sub

Private Sub Class_Terminate()
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Private Function CreateRecordset() As Recordset
Dim sSQL As String
    
    sSQL = " SHAPE APPEND new adGUID as Exchange_GUID, " & _
           "New adDate as Ex_Date, " & _
           "New adInteger as Amount, " & _
           "New adInteger as Operator_ID, " & _
           "New adInteger as CostValue, " & _
           "New adVarChar(20) as Till_ID, " & _
           "New adSingle as General_Discount, " & _
           "((SHAPE APPEND New adGUID as ExchangeID, " & _
           "New adInteger as ProductID, " & _
           "New adInteger as SellPrice, " & _
           "New adInteger as Nominal_SellPrice, " & _
           "New adSingle as Discount_Line, " & _
           "New adInteger as Deposit_Credited, " & _
           "New adInteger as Cost_Value, " & _
           "New adGUID as Product_GUID, " & _
           "New adVarChar(13) as Product_Code, " & _
           "New adVarChar(50) as Product_Description, " & _
           "New adVarChar(1) as Exception_Code) " & _
           "RELATE Exchange_GUID to ExchangeID) as CashSaleLine, "
           
 sSQL = sSQL & "((SHAPE APPEND New adGUID as ExchangeID, " & _
           "New adInteger as AMT_Tendered, " & _
           "New adDate as CCExpiryDate, " & _
           "New adVarChar(25) as CCNumber, " & _
           "New adVarChar(1) as PaymentType) " & _
           "RELATE Exchange_GUID to ExchangeID) as Payment"
           
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = "provider=msdatashape;data provider=none;"
    rs.Open sSQL, , adOpenStatic, adLockOptimistic
    
'    rs.AddNew
'
'    rs!Ex_Date = Now
'    rs!Amount = 35467
'    rs!Till_ID = "ABCD456"
'
'    Set rsSL = rs("CashSaleLine").Value
'    rsSL.AddNew
'    rsSL!ProductID = 234543
'
'    Set rsPay = rs("Payment").Value
'    rsPay.AddNew
'    rsPay!AMT_Tendered = 3456
'    rsPay!CCNumber = "What an Idiot!"
'
'
'    rs.Update
'    rsSL.Update
'
'    rsPay.Update
'    rsPay.MoveFirst
'    rsSL.MoveFirst
'    rs.MoveFirst
'
'
'    rsPay.Close
'    rsSL.Close
'    rs.Close
'
'    Set rs = Nothing
'    Set rsSL = Nothing
'    Set rsPay = Nothing
End Function

Public Sub AddExchange(Amount As Double, OperatorID As Long, _
        GenDiscount As Single, CostValue As Double)
Dim oGUID As New CGuid
Dim msg As String

    If rs Is Nothing Then
        msg = "Recordset not activated!"
        GoTo EH
    End If
    On Error GoTo EH
    With rs
        .AddNew
        !Exchange_GUID = oGUID.GetGuid
        !EX_Date = Now
        !Operator_ID = OperatorID
        !Amount = CLng(Amount * 100)
        !General_Discount = GenDiscount
        !CostValue = CLng(CostValue * 100)
        .Update
        iExchange = iExchange + 1
    End With
    Set oGUID = Nothing
    Exit Sub
EH:
    If msg = "" Then msg = "Can't add ExchangeData to Recordset"
    Err.Raise 456, , msg
End Sub

Public Sub AddSaleLine(ProductID As Long, SellPrice As Double, NomSellPrice As Double, _
        LineDisc As Single, DepositCredited As Double, CostValue As Double, _
        ProdCode As String, ProdDescription As String, ExceptionCode As String)
        
Dim rsSL As ADODB.Recordset
Dim msg As String

    If rs Is Nothing Then
        msg = "Recordset not activated!"
        GoTo EH
    End If
    If IsNull(rs!Exchange_GUID) Then
        msg = "Exchange Data not loaded!"
        GoTo EH
    End If
    On Error GoTo EH
    
    Set rsSL = rs("CashSaleLine").Value
    With rsSL
        !ExchangeID = rs!Exchange_GUID
        !ProductID = ProductID
        !SellPrice = CLng(SellPrice * 100)
        !Nominal_SellPrice = CLng(NomSellPrice * 100)
        !Discount_Line = LineDisc
        !Deposit_Credited = CLng(DepositCredited * 100)
        !Cost_Value = CLng(CostValue)
        If ProductID = 0 Then
            Dim oGUID As New CGuid
            !Product_Code = ProdCode
            !Product_GUID = oGUID.GetGuid
            !Product_Description = ProdDescription
            !Exception_Code = "N"
            Set oGUID = Nothing
        End If
        .Update
        iSaleLine = iSaleLine + 1
    End With
    Set rsSL = Nothing
    Exit Sub
EH:
    If msg = "" Then msg = "Can't add SaleLineData to Recordset"
    Err.Raise 456, , msg
End Sub

Public Sub AddPayment(AmtTendered As Double, CCExpDate As String, _
        CCNumber As String, PaymentType As String)
        
Dim rsP As ADODB.Recordset
Dim msg As String

    If rs Is Nothing Then
        msg = "Recordset not activated!"
        GoTo EH
    End If
    If IsNull(rs!Exchange_GUID) Then
        msg = "Exchange Data not loaded!"
        GoTo EH
    End If
    On Error GoTo EH
    
    Set rsP = rs("Payment").Value
    With rsP
        !ExchangeID = rs!Exchange_GUID
        !Amt_Tendered = CLng(AmtTendered * 100)
        If IsDate(CCExpDate) Then !CCExpiryDate = CDate(CCExpDate)
        !CCNumber = CCNumber
        !PaymentType = PaymentType
        .Update
        iPayment = iPayment + 1
    End With
    Set rsP = Nothing
    Exit Sub
EH:
    If msg = "" Then msg = "Can't add SaleLineData to Recordset"
    Err.Raise 456, , msg
End Sub

Public Sub SendExchange()
Dim msg As String
Dim QI As MSMQQueueInfo
Dim qRespQ As MSMQQueue
Dim qMsgResp As MSMQMessage

Dim xDisp As MSMQTransactionDispenser
Dim xAct As MSMQTransaction


    If iExchange = 0 Then
        msg = "Exchange data missing!"
    ElseIf iSaleLine = 0 Then
        msg = "SaleLine data missing!"
    ElseIf iPayment = 0 Then
        msg = "Payment data missing!"
    End If
    If msg <> "" Then GoTo EH
    On Error GoTo EH
    
    Set QI = New MSMQQueueInfo
    QI.PathName = ServerPath
    
    Set qRespQ = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
    Set qMsgResp = New MSMQMessage
    
    qMsgResp.Body = rs
    qMsgResp.Label = "Sale Record from POS: " & sCompID
    Set xDisp = New MSMQTransactionDispenser
    Set xAct = xDisp.BeginTransaction
    
    qMsgResp.Send qRespQ, xAct
    xAct.Commit
    qRespQ.Close
        
    Set QI = Nothing
    Set qRespQ = Nothing
    Set qMsgResp = Nothing
    Set xDisp = Nothing
    Set xAct = Nothing
    Exit Sub
EH:
    msg = "Can't send Sale data!" & vbLf & msg
    Err.Raise 456, , msg
End Sub

Public Property Get ServerPath() As String
    ServerPath = sServerPath
End Property
