VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMSMQExchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type tPayment
    Amount As Double
    CCExpDate As String
    CCNumber As String
    CCType As String
    Type As String
    TotPay As Double
    Change As Double
End Type

Dim rs As ADODB.Recordset

Const sInPath = ".\PRIVATE$\PapyPOS_In"
Dim sOutPath As String
Dim sCompID As String

Dim iExchange As Integer
Dim iSaleLine As Integer
Dim iPayment As Integer
Dim bRSOK As Boolean
Dim bServPathOK As Boolean


Private Sub Class_Initialize()
    Set rs = LoadRecordset
    bRSOK = Not rs Is Nothing
    bServPathOK = ReadQPathFromReg
    
    If Not bServPathOK Then
        bServPathOK = ReadQPathFromFile
    End If
    CreateInPath
    
End Sub

Private Sub Class_Terminate()
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    Set rs = Nothing
End Sub

Public Property Get RecSetLoaded() As Boolean
    RecSetLoaded = bRSOK
End Property
Public Property Get ServerPathOK() As Boolean
    ServerPathOK = bServPathOK And sOutPath <> ""
End Property


'Private Function CreateRecordset() As Recordset
'Dim sSQL As String
'
'    sSQL = " SHAPE {tblExchange} APPEND new adGUID as Exchange_GUID, " & _
'           "New adDate as Ex_Date, " & _
'           "New adInteger as Amount, " & _
'           "New adInteger as Operator_ID, " & _
'           "New adInteger as CostValue, " & _
'           "New adVarChar(20) as Till_ID, " & _
'           "New adSingle as General_Discount, " & _
'           "((SHAPE {tblCashSaleLine} APPEND New adGUID as ExchangeID, " & _
'           "New adInteger as ProductID, " & _
'           "New adInteger as SellPrice, " & _
'           "New adInteger as Nominal_SellPrice, " & _
'           "New adSingle as Discount_Line, " & _
'           "New adInteger as Deposit_Credited, " & _
'           "New adInteger as Cost_Value, " & _
'           "New adGUID as Product_GUID, " & _
'           "New adVarChar(13) as Product_Code, " & _
'           "New adVarChar(50) as Product_Description, " & _
'           "New adVarChar(1) as Exception_Code) " & _
'           "RELATE Exchange_GUID to ExchangeID) as CashSaleLine, "
'
' sSQL = sSQL & "((SHAPE {tblPayment} APPEND New adGUID as ExchangeID, " & _
'           "New adInteger as AMT_Tendered, " & _
'           "New adDate as CCExpiryDate, " & _
'           "New adVarChar(25) as CCNumber, " & _
'           "New adVarChar(1) as PaymentType) " & _
'           "RELATE Exchange_GUID to ExchangeID) as Payment"
'
'    If Not rs Is Nothing Then
'        If rs.State = adStateOpen Then rs.Close
'        Set rs = Nothing
'    End If
'    Set rs = New ADODB.Recordset
'    rs.ActiveConnection = "provider=msdatashape;data provider=none;"
'    rs.Open sSQL, , adOpenStatic, adLockOptimistic
    
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
'End Function

Private Function LoadRecordset() As ADODB.Recordset
Dim fs As New FileSystemObject
Dim rs As ADODB.Recordset
Dim msg As String

    If Not fs.FileExists(App.Path & "\POS.rs") Then
        msg = "Can't find Recordset file!"
        GoTo EH
    End If
    On Error GoTo EH
    Set rs = New ADODB.Recordset
    rs.Open (App.Path & "\POS.rs")
    Set LoadRecordset = rs
    Exit Function
EH:
    If msg = "" Then msg = Err.Description
    MsgBox msg
    
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
        !Till_ID = "XYZ"        '<<-----------------needs proper value!!!
        .Update
        iExchange = iExchange + 1
    End With
    Set oGUID = Nothing
    Exit Sub
EH:
    If msg = "" Then msg = "Can't add ExchangeData to Recordset"
    Err.Raise 456, , msg
End Sub

Public Sub AddSaleLine(ProductID As Long, QTY As Integer, SellPrice As Double, NomSellPrice As Double, _
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
    
    Set rsSL = rs("rsSaleLine").Value
    With rsSL
        .AddNew
'        !ExchangeID = rs!Exchange_GUID
        !ProductID = ProductID
        !SellPrice = CLng(SellPrice * 100)
        !QTY = QTY
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
        ElseIf ExceptionCode = "P" Then
            !Product_Code = ProdCode
            !Product_Description = ProdDescription
            !Exception_Code = "P"
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
        CCNumber As String, PaymentType As String, TotAmount As Double, Change As Double)
        
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
    
    Set rsP = rs("rsPayment").Value
    With rsP
        .AddNew
'        !ExchangeID = rs!Exchange_GUID
        !Amt_Tendered = CLng(AmtTendered * 100)
        If IsDate(CCExpDate) Then !CCExpiryDate = CDate(CCExpDate)
        !CCNumber = CCNumber
        !PaymentType = PaymentType
        !Tot_Received = CLng(TotAmount * 100)
        !Change = CLng(Change * 100)
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
    QI.PathName = sOutPath
    
    Set qRespQ = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
    Set qMsgResp = New MSMQMessage
    
    qMsgResp.Body = rs
    qMsgResp.Label = "SalesData" '"Sale Record from POS: " & sCompID
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

'Public Property Get ServerPath() As String
'    ServerPath = sInPath
'End Property

Private Sub SaveQPathToReg()
    SaveSetting "POSClient", "QPath", "OutPath", sOutPath
'    SaveSetting "POSClient", "QPath", "InPath", sInPath
End Sub

Private Function ReadQPathFromReg() As Boolean
Dim sTmp As String

    sTmp = GetSetting("POSClient", "QPath", "OutPath", "")
    If sTmp <> "" Then
        If TestMSMQPath(sTmp) Then sOutPath = sTmp
    End If
'    sInPath = GetSetting("POSClient", "QPath", "InPath", "")
    ReadQPathFromReg = sOutPath <> "" 'And sInPath <> ""
End Function

Private Function ReadQPathFromFile() As Boolean
Dim fs As New FileSystemObject
Dim f As TextStream
Dim sTmp As String

    On Error GoTo EH
       
    If Not fs.FileExists(App.Path & "\QPath.dat") Then GoTo EH
    Set f = fs.OpenTextFile(App.Path & "\QPath.dat", ForReading)
    sTmp = f.ReadLine
    If Not TestMSMQPath(sTmp) Then GoTo EH
    
    
    sOutPath = sTmp
'    sInPath = f.ReadLine
    
    ReadQPathFromFile = sOutPath <> "" 'And sInPath <> ""
MEX:
    f.Close
    Set fs = Nothing
    Set f = Nothing
    Exit Function
EH:
    GoTo MEX
End Function
Private Sub SaveQPathToFile()
Dim fs As New FileSystemObject
Dim f As TextStream

    On Error GoTo EH
       
    Set f = fs.CreateTextFile(App.Path & "\QPath.dat", True)
    f.WriteLine (sOutPath)
'    f.WriteLine (sInPath)
    f.Close
MEX:
    Set fs = Nothing
    Set f = Nothing
    Exit Sub
EH:
    GoTo MEX
End Sub

Public Sub SavePath(OutPath As String, SaveType As String)
    sOutPath = OutPath
'    sInPath = InPath
    Select Case SaveType
        Case "both"
            SaveQPathToReg
            SaveQPathToFile
        Case "reg"
            SaveQPathToReg
        Case "file"
            SaveQPathToFile
    End Select
End Sub

Public Function TestMSMQPath(sPathName As String) As Boolean
Dim QI As New MSMQQueueInfo
Dim msgQ As MSMQQueue
    On Error GoTo EH
    QI.PathName = sPathName
    QI.Refresh
    
    TestMSMQPath = True
MEX:
    Set QI = Nothing
    Exit Function
EH:
    Err.Clear
    GoTo MEX
End Function

Public Function RequestRSFromServer() As Boolean
Dim msg As String
Dim bStop As Boolean
Dim QI As MSMQQueueInfo
Dim respQI As MSMQQueueInfo
Dim qRespQ As MSMQQueue
Dim qMsg As MSMQMessage

Dim xDisp As MSMQTransactionDispenser
Dim xAct As MSMQTransaction
   
Dim siStart As Single
Dim siElapse As Single
  
    On Error GoTo EH
    
    'Try to load RS from previous requests
    If GetRSFromServer(0) Then
        RequestRSFromServer = True
        Exit Function
    End If
    
    'set up request for RS
    Set QI = New MSMQQueueInfo
    Set respQI = New MSMQQueueInfo
    
    QI.PathName = sOutPath
    respQI.PathName = sInPath
    respQI.Label = "RS"
    
    Set qRespQ = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
    Set qMsg = New MSMQMessage
    Set qMsg.ResponseQueueInfo = respQI
    qMsg.Label = "Send RS"
    Set xDisp = New MSMQTransactionDispenser
    Set xAct = xDisp.BeginTransaction
    
    qMsg.Send qRespQ, xAct
    xAct.Commit
    qRespQ.Close
    
    'Try to load response
    RequestRSFromServer = GetRSFromServer(15)
    
    Set QI = Nothing
    Set qRespQ = Nothing
    Set qMsg = Nothing
    Set xDisp = Nothing
    Set xAct = Nothing
    Exit Function
EH:
    msg = "Retreiving RS Data failed!"
    Err.Raise 456, , msg

End Function

Private Function GetRSFromServer(Timeout As Single) As Boolean
Dim bStop As Boolean
Dim QI As New MSMQQueueInfo
Dim MQ As MSMQQueue
Dim qMsg As MSMQMessage
Dim siStart As Single
Dim siElapse As Single

    QI.PathName = sInPath
    Set MQ = QI.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
    
    'set timeout length in seconds
    siElapse = Timeout
    siStart = Timer
    Do While Not bStop
        DoEvents
        Set qMsg = MQ.PeekCurrent(ReceiveTimeOut:=100)
        'loop through all received messages in case
        'there are more then one
        Do While Not qMsg Is Nothing
            If qMsg.Label = "RS" Then
                'only process if RS is still missing
                If Not bRSOK Then
                    Dim rs As ADODB.Recordset
                    Set rs = qMsg.Body
                    'Save RS to disc
                    rs.Save App.Path & "\POS.rs", adPersistADTG
                    bRSOK = True
                End If
                'remove message from queue
                Set qMsg = MQ.ReceiveCurrent(ReceiveTimeOut:=100)
                GetRSFromServer = True
                bStop = True
            End If
            'try to load next message
            On Error Resume Next
            Set qMsg = MQ.PeekNext(ReceiveTimeOut:=100)
            bStop = qMsg Is Nothing
        Loop
        If Not bStop Then
            bStop = (Timer > siStart + siElapse)
        End If
    Loop
    MQ.Close
    Set QI = Nothing
    Set MQ = Nothing
    Set qMsg = Nothing
End Function

Private Sub CreateInPath()
Dim qInfo As New MSMQQueueInfo
    On Error Resume Next
    With qInfo
        .PathName = sInPath
        .Label = "Papyrus POS IN queue"
        .Create True
        If Err Then
            If Err <> MQ_ERROR_QUEUE_EXISTS Then
                bServPathOK = False
            End If

            Err.Clear
        End If
    End With
 Set qInfo = Nothing
End Sub
