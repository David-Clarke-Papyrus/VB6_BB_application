VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HandleRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event DoneReq(IsSale As Boolean)

Private xClientNames() As String
Private oDP As z_DatabasePersist

Dim sInPath As String
Dim bDbConOK As Boolean



Private Sub Class_Initialize()
    If Not ReadQPathFromReg Then
        ReadQPathFromFile
    End If
    Set oDP = New z_DatabasePersist
    bDbConOK = oDP.dbConnect("", "") = 0
    ReDim xClientNames(0)
    LoadClientNames
End Sub

Private Sub Class_Terminate()
    If oDP.IsConnected Then
        oDP.dbCloseConnect
    End If
    Set oDP = Nothing
'    ReqQ.Close
'    Set ReqQ = Nothing
End Sub

Public Property Get DBName() As String
    DBName = oDP.DBName
End Property

Public Function CheckNewClient() As Boolean
'Dim QI As MSMQQueueInfo
'Dim qReqQ As MSMQQueue
'Dim qMsg As MSMQMessage
'
'Dim QI2 As MSMQQueueInfo
'Dim qRespQ As MSMQQueue
'Dim qMsgResp As MSMQMessage
'
''Dim lID As Long
''Dim sISBN As String
'Dim xVal
'Dim xDisp As MSMQTransactionDispenser
'Dim xAct As MSMQTransaction
'
'    On Error GoTo EH
'
'    Set QI = New MSMQQueueInfo
'    QI.PathName = sInPath '"urs\private$\PapyPos"
'    Set qReqQ = QI.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
'    Set qMsg = qReqQ.PeekCurrent(ReceiveTimeout:=100)
'    Do While Not qMsg Is Nothing
'        If qMsg.Label = "Registration request" Then
'            AddClient qMsg.Body
'            SaveClientNames
'
'
'            Set QI2 = qMsg.ResponseQueueInfo
'            Set qRespQ = QI2.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
'            Set qMsgResp = New MSMQMessage
'
'            qMsgResp.Body = ""
'            qMsgResp.Label = "Client Registered OK!"
'            qMsgResp.CorrelationId = qMsg.ID
'
'            Set xDisp = New MSMQTransactionDispenser
'            Set xAct = xDisp.BeginTransaction
'
'            qMsgResp.Send qRespQ, xAct
'            xAct.Commit
'            qRespQ.Close
'            'remove message
'            Set qMsg = qReqQ.ReceiveCurrent(ReceiveTimeout:=100)
'            RaiseEvent DoneReq(False)
'        End If
'        On Error Resume Next
'        Set qMsg = qReqQ.PeekNext(ReceiveTimeout:=100)
'    Loop
'MEX:
'
'    Exit Function
'EH:
'    Err.Raise 400, "cls_HandleRequest CheckNewClient", Error

End Function


Private Sub AddClient(sName As String)
Dim sTmp As String
    'CHECK IF CLIENT ALLREADY EXISTS
    sTmp = Join(xClientNames, ", ")
    If InStr(1, sTmp, sName) > 0 Then Exit Sub
    On Error Resume Next
    If xClientNames(UBound(xClientNames)) <> "" Then
        ReDim Preserve xClientNames(UBound(xClientNames) + 1)
    End If
    xClientNames(UBound(xClientNames)) = sName
End Sub

Private Function SaveClientNames() As Integer
Dim fs As New FileSystemObject
Dim f As TextStream
Dim sTmp As String
Dim i As Integer

    If xClientNames(0) = "" Then
        If fs.FileExists(App.Path & CLIENT_FILE) Then
            fs.DeleteFile App.Path & CLIENT_FILE, True
        End If
        Set fs = Nothing
        Exit Function
    End If
    
    Set f = fs.CreateTextFile(App.Path & CLIENT_FILE, True)
    For i = 0 To UBound(xClientNames)
        f.WriteLine xClientNames(i)
    Next i
    f.Close
    Set fs = Nothing
    Set f = Nothing
    SaveClientNames = i
End Function

Public Function LoadClientNames() As Integer
Dim fs As New FileSystemObject
Dim f As TextStream
Dim sTmp As String
Dim i As Integer
Dim msg As String
        
    On Error GoTo EH
    If Not fs.FileExists(App.Path & CLIENT_FILE) Then
        msg = "Can't find file: " & CLIENT_FILE
        GoTo EH
    End If
    
    Set f = fs.OpenTextFile(App.Path & CLIENT_FILE, ForReading, False)
    Do While Not f.AtEndOfStream
        sTmp = ""
        sTmp = f.ReadLine
        AddClient sTmp
        i = i + 1
    Loop
    LoadClientNames = i
    f.Close
MEX:
    
    Set fs = Nothing
    Set f = Nothing
Exit Function
EH:
    GoTo MEX
End Function




Public Sub GetMessage()
'Dim QI As MSMQQueueInfo
'Dim qReqQ As MSMQQueue
'Dim qMsg As MSMQMessage
'
'Dim QI2 As MSMQQueueInfo
'Dim qRespQ As MSMQQueue
'Dim qMsgResp As MSMQMessage
'
'Dim lID As Long
'Dim sISBN As String
'Dim xVal
'Dim xDisp As MSMQTransactionDispenser
'Dim xAct As MSMQTransaction
'Dim QID As Variant
'
'    On Error GoTo EH
'
'    Set QI = New MSMQQueueInfo
'    QI.PathName = sInPath '"urs\private$\PapyPos"
'    Set qReqQ = QI.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
'    'Set qMsg = qReqQ.Receive(ReceiveTimeOut:=100)
'    Set qMsg = qReqQ.Peek(WantBody:=True, ReceiveTimeout:=100)
'    If Not qMsg Is Nothing Then
'        QID = qMsg.ID
''        Set QI2 = qMsg.ResponseQueueInfo
''        Set qRespQ = QI2.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
''        Set qMsgResp = New MSMQMessage
'
'        If qMsg.Label = "Registration request" Then
'
'
'            AddClient qMsg.Body
'            SaveClientNames
'            qMsgResp.Body = ""
'            qMsgResp.Label = "Client Registered OK!"
'        ElseIf qMsg.Label = "SalesData" Then
'            Dim rs As ADODB.Recordset
'            Set rs = qMsg.Body
'            If Not SaveSaleData(rs) Then
'                'keep message in queue
'
'            End If
'        Else
'
'            xVal = Split(qMsg.Body, "=")
'
'            If xVal(0) = "IDNum" Then
'                lID = Val(xVal(1))
'            Else
'                sISBN = xVal(1)
'            End If
'            If lID = 0 And Trim$(sISBN) = "" Then Exit Sub
'
'
'            qMsgResp.Body = LoadDataset(lID, sISBN)
'            qMsgResp.Label = qMsg.Label
'
'        End If
'
'        qMsgResp.CorrelationId = qMsg.ID
'
'        Set xDisp = New MSMQTransactionDispenser
'        Set xAct = xDisp.BeginTransaction
'
'        qMsgResp.Send qRespQ, xAct
'        xAct.Commit
'        qRespQ.Close
'        RaiseEvent DoneReq(False)
'
'        'This wil remove message from queue after succsessful saving
'
'        Set qMsg = qReqQ.Receive(WantBody:=True, ReceiveTimeout:=100)
'
'    End If
'    Exit Sub
'EH:
'    Err.Raise 400, "cls_HandleRequest GetMessage", Error
End Sub

Public Sub GetSales()
'Dim QI As MSMQQueueInfo
'Dim qReqQ As MSMQQueue
'Dim qMsg As MSMQMessage
'Dim sRetPath As String
'
''Dim QID As Variant
'
'    On Error GoTo EH
'
'    Set QI = New MSMQQueueInfo
'    QI.PathName = sInPath '"urs\private$\PapyPos_resp"
'    Set qReqQ = QI.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
'    Set qMsg = qReqQ.PeekCurrent(WantBody:=True, ReceiveTimeout:=100)
'    Do While Not qMsg Is Nothing
''        QID = qMsg.
'        If qMsg.Label = "SalesData" Then
'            Dim rs As ADODB.Recordset
'            Set rs = qMsg.Body
'            If SaveSaleData(rs) Then
'                'This wil remove message from queue after succsessful saving
'                Set qMsg = qReqQ.Receive(WantBody:=True, ReceiveTimeout:=100)
'                RaiseEvent DoneReq(True)
'            Else
'                GoTo EH
'            End If
'        ElseIf qMsg.Label = "Send RS" Then
'            'Send Recordset
'            'check if not allready send to  this client
'            If qMsg.ResponseQueueInfo.PathName <> sRetPath Then
'                sRetPath = qMsg.ResponseQueueInfo.PathName
'                AddClient sRetPath
'                SendRecordset qMsg
'                RaiseEvent DoneReq(False)
'            End If
'            'This wil remove message from queue after processing
'            Set qMsg = qReqQ.ReceiveCurrent(WantBody:=True, ReceiveTimeout:=100)
'
'        End If
'        Set qMsg = qReqQ.Peek(WantBody:=True, ReceiveTimeout:=100)
'
'    Loop 'While Not qMsg Is Nothing
'
'
'    Exit Sub
'EH:
'    Err.Raise 400, "cls_HandleRequest GetMessage", Error
End Sub

'Private Sub SendRecordset(qMsg As MSMQMessage)
'Dim QI As MSMQQueueInfo
''Dim respQI As New MSMQQueueInfo
'Dim respMSG As New MSMQMessage
'Dim QResp As MSMQQueue
'Dim xDisp As MSMQTransactionDispenser
'Dim xAct As MSMQTransaction
'
'Dim rs As ADODB.Recordset
'
'    respMSG.Label = "RS"
'    'load recordset
'    Set rs = CreateDisconnRecSet(False)
'
'    respMSG.Body = rs
'
'    respMSG.CorrelationId = qMsg.ID
'    Set QI = qMsg.ResponseQueueInfo
'    Set QResp = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
'    Set xDisp = New MSMQTransactionDispenser
'    Set xAct = xDisp.BeginTransaction
'    respMSG.Send QResp, xAct
'    xAct.Commit
'    QResp.Close
'
'    Set QI = Nothing
'    Set QResp = Nothing
'    Set xDisp = Nothing
'    Set xAct = Nothing
'    Set respMSG = Nothing
'End Sub

Private Function LoadDataset(lID As Long, sISBN As String) As Recordset
Dim rs As ADODB.Recordset
Dim sSQL As String
'Dim oDP As z_DatabasePersist
Dim iR As Integer
Dim r As Long
    
    On Error GoTo EH
    sSQL = "SELECT Product_ID, P_ISBN,P_Publisher, P_SeriesTitle, " & _
                 "P_MainAuthor, P_Title, P_SAPrice " & _
           "FROM Product "
    If lID > 0 Then
        sSQL = sSQL & "WHERE Product_ID=" & lID
    ElseIf sISBN <> "" Then
        sSQL = sSQL & "WHERE P_ISBN='" & sISBN & "'"
    Else
        Exit Function
    End If
'    Set oDP = New z_DatabasePersist

    If Not oDP.IsConnected Then
        iR = oDP.dbConnect("URS", "sru")
        If iR <> 0 Then Err.Raise 400, "", "Can't create DB connection!"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    rs.Open sSQL, oDP.dbConn
    Set rs.ActiveConnection = Nothing
    
    Set LoadDataset = rs
MEX:
'    oDP.dbCloseConnect
    Set rs = Nothing
    Exit Function
EH:
    Err.Raise 400, "cls_HandleRequest LoadDataset", Error
    GoTo MEX
End Function

Public Function PollDB() As Boolean
'Dim oDP As New z_DatabasePersist
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim iR As Integer

    On Error GoTo EH
    If Not oDP.IsConnected Then
        iR = oDP.dbConnect("", "")
        If iR <> 0 Then Err.Raise 400, "", "Can't create DB connection!"
    End If
    
    oDP.dbConn.BeginTrans
    
    sSQL = "SELECT * FROM tmpProduct"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    rs.Open sSQL, oDP.dbConn
    'DISCONNECT RECORDSET
    rs.ActiveConnection = Nothing
    
    If Not rs.EOF Then
        If SendUpdate(rs) Then
            sSQL = "DELETE tmpProduct"
            oDP.dbConn.Execute sSQL
            oDP.dbConn.CommitTrans
            PollDB = True
            RaiseEvent DoneReq(False)
        Else
            oDP.dbConn.RollbackTrans
        End If
    End If
MEX:
'   oDP.dbCloseConnect
'   Set oDP = Nothing
    If rs.State = adStateOpen Then rs.Close
   Set rs = Nothing
    
Exit Function
EH:
    oDP.dbConn.RollbackTrans
    GoTo MEX
End Function

Private Function SendUpdate(oRS As Recordset) As Boolean
Dim i As Integer
    If xClientNames(0) = "" Then
        'NO CLIENTS REGISTERED, ABORT
        Exit Function
    End If
    
    For i = 0 To UBound(xClientNames)
        ClientUpdate xClientNames(i), oRS
        
    Next i
    SendUpdate = True
End Function

Private Sub ClientUpdate(sPathName As String, oRS As Recordset)
'Dim QI As MSMQQueueInfo
'Dim qRespQ As MSMQQueue
'Dim qMsgResp As MSMQMessage
'
'Dim xDisp As MSMQTransactionDispenser
'Dim xAct As MSMQTransaction
'
'    Set QI = New MSMQQueueInfo
'    QI.PathName = sPathName
'    Set qRespQ = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
'    Set qMsgResp = New MSMQMessage
'
'    qMsgResp.Body = oRS
'    qMsgResp.Label = "New DB Update"
''    qMsgResp.CorrelationId = qMsg.ID
'
'    Set xDisp = New MSMQTransactionDispenser
'    Set xAct = xDisp.BeginTransaction
'
'    qMsgResp.Send qRespQ, xAct
'    xAct.Commit
'    qRespQ.Close
End Sub

Public Sub LoadClientList(oList As ListBox)
Dim i As Integer
    
    oList.Clear
    If xClientNames(0) = "" Then Exit Sub
    For i = 0 To UBound(xClientNames)
        oList.AddItem xClientNames(i)
    Next i
    
End Sub

Public Sub SaveClientList(oList As ListBox)
Dim i As Integer
    ReDim xClientNames(0)
    For i = 0 To oList.ListCount - 1
        If Not xClientNames(UBound(xClientNames)) = "" Then
            ReDim Preserve xClientNames(UBound(xClientNames) + 1)
        End If
        
        xClientNames(UBound(xClientNames)) = oList.List(i)
    Next i
    SaveClientNames
End Sub

Private Function SaveSaleData(rs As ADODB.Recordset) As Boolean
'Dim oDP As New z_DatabasePersist
Dim rsSL As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim msg As String
    If Not oDP.IsConnected Then
        If oDP.dbConnect("", "") <> 0 Then
            msg = "Can't connect to database!"
            GoTo EH
        End If
    End If
    On Error GoTo EH
    oDP.dbConn.BeginTrans
    rs.ActiveConnection = oDP.dbConn
    rs.MoveFirst
'    rs!Till_ID = "ABC"
    Debug.Print rs!Exchange_GUID
    rs.UpdateBatch
    Set rsSL = rs("rsSaleLine").Value
    rsSL.MoveFirst
    Do While Not rsSL.EOF
        rsSL.UpdateBatch
        rsSL.MoveNext
    Loop
    Set rsP = rs("rsPayment").Value
    rsP.MoveFirst
    Do While Not rsP.EOF
        rsP.UpdateBatch
        rsP.MoveNext
    Loop
    oDP.dbConn.CommitTrans
    SaveSaleData = True
MEX:
    rs.Close
    Set rs = Nothing
    Set rsSL = Nothing
    Set rsP = Nothing
'    oDP.dbCloseConnect
    Exit Function
EH:
    If Err = -2147217900 Then 'Data allready saved
        oDP.dbConn.RollbackTrans
        SaveSaleData = True
        Err.Clear
        GoTo MEX
    End If
    If msg <> "" Then
        MsgBox "Can't save Sales Data!" & vbLf & msg
    Else
        oDP.dbConn.RollbackTrans
        MsgBox "Can't save Sales Data!" & vbLf & Err.Description
    End If

End Function

Public Function CreateDisconnRecSet(SaveToDisc As Boolean) As ADODB.Recordset
Dim oShapeDB As New z_DatabasePersist
Dim rs As New ADODB.Recordset
Dim sSQL As String
Dim msg As String

    If oShapeDB.dbConnect("", "", "MSDataShape") <> 0 Then
        msg = "Failed to create database connection!"
        GoTo EH
    End If
    sSQL = "SHAPE {SELECT * FROM tExchange WHERE tExchange.Exchange_GUID = Null} " & _
           "APPEND ({SELECT * FROM tCSL WHERE tCSL.ExchangeID = Null} " & _
           "RELATE Exchange_GUID TO ExchangeID) AS rsSaleLine, " & _
           "({SELECT * FROM tPayment WHERE tPayment.ExchangeID = Null} " & _
           "RELATE Exchange_GUID TO ExchangeID) AS rsPayment"

'    sSQL = "SELECT * FROM tblCashSaleLine WHERE ExchangeID = Null"
    On Error GoTo EH
    rs.ActiveConnection = oShapeDB.dbConn
    rs.CursorLocation = adUseClient
    rs.Open sSQL, , adOpenStatic, adLockBatchOptimistic
    Set rs.ActiveConnection = Nothing
    Set CreateDisconnRecSet = rs
    
    If SaveToDisc Then SaveRSToFile rs
    
MEX:
    oShapeDB.dbCloseConnect
    Set oShapeDB = Nothing
    Set rs = Nothing
    Exit Function
EH:
    If msg = "" Then msg = Err.Description
    MsgBox msg
    Err.Clear
    GoTo MEX
    

End Function

Private Sub SaveRSToFile(rs As ADODB.Recordset)
Dim fs As New FileSystemObject
Dim sPath As String
    
    sPath = App.Path & "\Test.rs"
    If fs.FileExists(sPath) Then
        fs.DeleteFile sPath, True
    End If
    rs.Save sPath, adPersistADTG
    
End Sub
Public Property Get DBConnectionOK() As Boolean
    DBConnectionOK = bDbConOK
End Property

Public Property Get ServerPathOK() As Boolean
    ServerPathOK = Len(sInPath) > 0
End Property
Public Sub SavePath(InPath As String, SaveType As String)
'    sOutPath = OutPath
    sInPath = InPath
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
Private Sub SaveQPathToReg()
'    SaveSetting "POSServer", "QPath", "OutPath", sOutPath
    SaveSetting "POSServer", "QPath", "InPath", sInPath
End Sub

Private Function ReadQPathFromReg() As Boolean
'Dim sTmp As String
''    sOutPath = GetSetting("POSServer", "QPath", "OutPath", "")
'    sTmp = GetSetting("POSServer", "QPath", "InPath", "")
'    If sTmp <> "" Then
'        If TestMSMQPath(sTmp) Then sInPath = sTmp
'    End If
'    ReadQPathFromReg = sInPath <> "" 'And sOutPath <> ""
End Function

Private Function ReadQPathFromFile() As Boolean
'Dim fs As New FileSystemObject
'Dim f As TextStream
'Dim sTmp As String
'    On Error GoTo EH
'
'    If Not fs.FileExists(App.Path & "\QPath.dat") Then GoTo EH
'    Set f = fs.OpenTextFile(App.Path & "\QPath.dat", ForReading)
''    sOutPath = f.ReadLine
'    sTmp = f.ReadLine
'    If Not TestMSMQPath(sTmp) Then GoTo MEX
'
'
'    sInPath = sTmp
'    ReadQPathFromFile = sInPath <> "" 'And sOutPath <> ""
'MEX:
'    f.Close
'    Set fs = Nothing
'    Set f = Nothing
'    Exit Function
'EH:
'    GoTo MEX
End Function
Private Sub SaveQPathToFile()
Dim fs As New FileSystemObject
Dim f As TextStream

    On Error GoTo EH
       
    Set f = fs.CreateTextFile(App.Path & "\QPath.dat", True)
'    f.WriteLine (sOutPath)
    f.WriteLine (sInPath)
    f.Close
MEX:
    Set fs = Nothing
    Set f = Nothing
    Exit Sub
EH:
    GoTo MEX
End Sub

'Public Function TestMSMQPath(sPathName As String) As Boolean
'Dim QI As New MSMQQueueInfo
'Dim msgQ As MSMQQueue
'    If sPathName = "" Then GoTo MEX
'    On Error GoTo EH
'    QI.PathName = sPathName
'    QI.Refresh
'
'    TestMSMQPath = True
'MEX:
'    Set QI = Nothing
'    Exit Function
'EH:
'    Err.Clear
'    GoTo MEX
'End Function
'
