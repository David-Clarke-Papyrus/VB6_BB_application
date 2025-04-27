VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExchange"
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
    Type As String ' M = Cash, C = Credid Card, P = Check, V = Voucher
    TotPay As Double
    Change As Double
End Type

Private Type tSLine
    Qty As Integer
    Code As String
    Author As String
    Title As String
    UnitPr As Double
    Disc As Single
    LineTot As Double
End Type

Event PollingStopped(Msg As String)

Dim rs As ADODB.Recordset
Dim sLines() As tSLine

Dim WithEvents oPoll As clsPoll
Attribute oPoll.VB_VarHelpID = -1

Dim sRoot As String
Dim sInBox As String
Dim sOutBox As String
Dim sServerPath As String
Dim sTillCode As String
Dim sSharedInboxName As String

Dim sBill As String
Dim sCustomer As String

Dim iExchange As Integer
Dim iSaleLine As Integer
Dim iPayment As Integer
Dim bRSOK As Boolean
Dim sGUID As String
Dim sPaymentType As String
Dim oGD As z_GetData


Private Sub Class_Initialize()
    On Error GoTo errHandler
    
    ConfigureFromDB
    CheckFolders
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.Class_Initialize"
End Sub

Public Function TryToStartPolling() As Boolean
    On Error GoTo errHandler
    If RegisterWithServer(False) = True Then
        Set oPoll = New clsPoll
        With oPoll
            .InBox = sRoot & sInBox
            .OutBox = sRoot & sOutBox
            .ServerPath = sServerPath
            If .SetDB(oPC.DBName, oPC.DBPath) Then
                .StartPolling POLL_INTERVAL
            End If
        End With
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.TryToStartPolling"
End Function
Private Sub Class_Terminate()
    On Error GoTo errHandler
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    Set rs = Nothing
    If Not oPoll Is Nothing Then
        With oPoll
            If .IsPolling Then .StopPolling ""
            Set oPoll = Nothing
        End With
    End If
'EH:
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.Class_Terminate"
End Sub

Private Sub CheckFolders()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim oFShare As New clsFolderShare
Dim sSharedName As String
    
    sRoot = oPC.SharedFolderRoot & "\"
    
    sInBox = "POSCLI_IN"
    sOutBox = "POSCLI_OUT"
    sSharedName = "POSCLI_IN_S"
    If Not oFShare.CreateFolder(sRoot & sInBox, True, sSharedName) Then
        MsgBox "Failed to set InBox as shared!"
        Exit Sub
    End If
    sSharedInboxName = "\\" & oFShare.GetCompName & "\" & sSharedName
    If Not fs.FolderExists(sRoot & sOutBox) Then
        fs.CreateFolder sRoot & sOutBox
    End If
    
    sServerPath = "\\" & oPC.servername & "\POSSVR_IN_S"
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.CheckFolders"
End Sub

Public Function RegisterWithServer(Optional bCheckInBox As Boolean) As Boolean
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim F As TextStream
Dim isStart As Single
Dim isDelay As Single
Dim bExit As Boolean
Dim Msg As String
Dim i As Integer

    If bCheckInBox Then GoTo CheckInbox
    RegisterWithServer = True
    If Not fs.FolderExists(sServerPath) Then
              RegisterWithServer = False
        GoTo MEX
    End If
    Set F = fs.CreateTextFile(sRoot & sOutBox & "\" & sTillCode & "_Reg.DAT")
    If Not F Is Nothing Then
        F.WriteLine TillCode
        F.WriteLine sSharedInboxName
        F.Close
        'move file to server
        If fs.FileExists(ServerPath & "\" & sTillCode & "_Reg.DAT") Then
            fs.DeleteFile ServerPath & "\" & sTillCode & "_Reg.DAT", True
        End If
        fs.MoveFile sRoot & sOutBox & "\" & sTillCode & "_Reg.DAT", ServerPath & "\" & sTillCode & "_Reg.DAT"
        If Err.Number <> 58 And Err.Number <> 0 Then
            Err.Clear
            Msg = "Can't send file to Server!" & vbLf & _
              "Make sure this computer has write access on the server computer!"
        GoTo errHandler
        End If
    End If
CheckInbox:
    isDelay = 10
    isStart = Timer
    Set F = Nothing
    bExit = False
    Do While Not bExit
        If fs.FileExists(sRoot & sInBox & "\RegOK.Dat") Then
            For i = 1 To 10
                On Error Resume Next
                Set F = fs.OpenTextFile(sRoot & sInBox & "\RegOK.Dat")
            Next i
            bExit = True
        End If
        If Not bExit Then
            bExit = Timer > isStart + isDelay
        End If
    Loop
    If F Is Nothing Then
        Msg = "Server response timed out!"
        GoTo MEX
    End If
    sServerPath = F.ReadLine
    F.Close
    If Not SaveSetup(sTillCode, sServerPath) Then
        Msg = "Failed to save TillCode and ServerPath to DB!"
        GoTo MEX
    Else
        fs.DeleteFile (sRoot & sInBox & "\RegOK.dat")

    End If
    RegisterWithServer = True
MEX:
    Set fs = Nothing
    Set F = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.RegisterWithServer(TillCode,bCheckInBox)", Array(TillCode, _
          bCheckInBox)
    Exit Function
    Resume
End Function

Public Property Get ServerPathOK() As Boolean
    On Error GoTo errHandler
    ServerPathOK = sServerPath <> ""
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.ServerPathOK"
End Property

Public Function LoadRecordset() As Boolean
On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim trs As ADODB.Recordset
Dim Msg As String
Dim iCnt As Long

    Set trs = New ADODB.Recordset
TOLoop:
    If fs.FileExists(sRoot & "POS.rs") Then
        trs.Open (sRoot & "POS.rs")
    Else
        RequestRSFromServer
        iCnt = iCnt + 1
        WAIT 1000
        DoEvents
        GoTo TOLoop
    End If
BOLoop:
    Check fs.FileExists(sRoot & "POS.rs"), EXC_MISSING_RS, "Cannot get response from server to request for recordset"
    
    Set rs = trs
    bRSOK = Not rs Is Nothing
    LoadRecordset = bRSOK
    ReDim sLines(0)
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.LoadRecordset"
End Function

Public Sub AddExchange(Amount As Double, OperatorID As Long, SupervisorID As Long, _
        GenDiscount As Single, CostValue As Double, Optional Customer As String, Optional pChange As Long)
    On Error GoTo errHandler
Dim oGUID As New CGuid
Dim Msg As String

'    If rs Is Nothing Then
'        Msg = "Recordset not activated!"
'        GoTo errHandler
'    ElseIf rs.State <> adStateOpen Then
'        rs.Open sRoot & "POS.rs"
'    End If
    With rs
        .AddNew
        sGUID = oGUID.GetGuid
        !EXCH_Exchange_GUID = sGUID
        !Exch_Date = Now
        !EXCH_Amount = CLng(Amount * 100)
        !EXCH_Till_Code = sTillCode
        !EXCH_Operator_ID = OperatorID
        !EXCH_General_Disc = GenDiscount
        !EXCH_CostValue = CLng(CostValue * 100)
        !EXCH_SupervisorID = SupervisorID
        !EXCH_ChangeGiven = pChange
        !EXCH_OpSessionID = oPC.GD.OpSessionID
        .Update
        iExchange = iExchange + 1
    End With
    
    If Customer <> "" Then sCustomer = Customer
    Set oGUID = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.AddExchange(Amount,OperatorID,SupervisorID,GenDiscount,CostValue,Customer)", _
         Array(Amount, OperatorID, SupervisorID, GenDiscount, CostValue, Customer)
End Sub

Public Sub AddSaleLine(ProductID As String, Qty As Integer, UnitPrice As Double, SellPrice As Double, _
        LineDisc As Single, ProdCode As String, Author As String, ProdDescription As String, ExceptionCode As String)
    On Error GoTo errHandler
        
Dim rsSL As ADODB.Recordset
Dim Msg As String
       
    If rs Is Nothing Then
        Msg = "Recordset not activated!"
        GoTo errHandler
    End If
    If IsNull(rs!EXCH_Exchange_GUID) Then
        Msg = "Exchange Data not loaded!"
        GoTo errHandler
    End If
    If sLines(UBound(sLines)).Code <> "" Then
        ReDim Preserve sLines(UBound(sLines) + 1)
    End If
    With sLines(UBound(sLines))
        .Code = ProdCode
        .Qty = Qty
        .Author = Author
        .Title = ProdDescription
        .UnitPr = UnitPrice
        .LineTot = UnitPrice * Qty
        .Disc = LineDisc
    End With
    Set rsSL = rs("rsSaleLine").Value
    With rsSL
        .AddNew
        !CSL_P_ID = ProductID
        !CSL_Qty = Qty
        !CSL_Price = CLng(UnitPrice * 100)
        !CSL_Discount = LineDisc
        !CSL_TimeOfSale = Now
  '      !CSL_Deposit_Credited = CLng(DepositCredited * 100)
  '      !CSL_Product_Code = ProdCode
        !CSL_Exception_Code = ExceptionCode
        
        .Update
        iSaleLine = iSaleLine + 1
    End With
    Set rsSL = Nothing
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.AddSaleLine(ProductID,Qty,UnitPrice,SellPrice,LineDisc,ProdCode,Author," & _
        "ProdDescription,ExceptionCode)", Array(ProductID, Qty, UnitPrice, SellPrice, LineDisc, ProdCode, _
         Author, ProdDescription, ExceptionCode)
End Sub

Public Sub AddPayment(AmountDue As Double, CCExpDate As String, _
        CCNumber As String, PaymentType As String, AmountPaid As Double)
    On Error GoTo errHandler
        
Dim rsP As ADODB.Recordset
Dim Msg As String

    If rs Is Nothing Then
        Msg = "Recordset not activated!"
        GoTo errHandler
    End If
    If IsNull(rs!EXCH_Exchange_GUID) Then
        Msg = "Exchange Data not loaded!"
        GoTo errHandler
    End If
    
    Set rsP = rs("rsPayment").Value
    With rsP
        .AddNew
        !PAY_Amt_Tendered = CLng(AmountDue * 100)
        If IsDate(CCExpDate) Then !PAY_CCExpiryDate = Right(CCExpDate, 4)
        !PAY_CCLastFour = Right(CCNumber, 4)
        !PAY_PaymentType = PaymentType
        !PAY_Tot_Received = CLng(AmountPaid * 100)
        .Update
        iPayment = iPayment + 1
    End With
    Set rsP = Nothing
'EH:
'    If Msg = "" Then Msg = "Can't add SaleLineData to Recordset"
'    Err.Raise 456, , Msg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.AddPayment(AmountDue,CCExpDate,CCNumber,PaymentType,AmountPaid)", _
         Array(AmountDue, CCExpDate, CCNumber, PaymentType, AmountPaid)
End Sub

Public Sub SaveSale()
    On Error GoTo errHandler
Dim SaleRS As New ADODB.Recordset
Dim PayRS As ADODB.Recordset
Dim sSQL As String
Dim Msg As String
    
    sSQL = "SELECT * FROM tSale WHERE Sale_ID = 0"
    
    Set PayRS = rs("rsPayment").Value
    With SaleRS
        .Open sSQL, oPC.DBConn, adOpenDynamic, adLockOptimistic
        .AddNew
        !S_GUID = sGUID
        !S_OperatorID = rs!EXCH_Operator_ID
        !S_Date = rs!Exch_Date
        !S_TotalDue = rs!EXCH_Amount
        PayRS.MoveFirst
        Do While Not PayRS.EOF
            Select Case PayRS!PAY_PaymentType
                Case "M" 'Cash
                    !S_CashAmount = PayRS!PAY_Amt_Tendered
                Case "C" 'Credit Card
                    !S_CardAmount = PayRS!PAY_Amt_Tendered
                Case "P" 'Check
                    !S_CheckAmount = PayRS!PAY_Amt_Tendered
                Case "V" 'Voucher
                    !S_VoucherAmount = PayRS!PAY_Amt_Tendered
            End Select
            PayRS.MoveNext
        Loop
        .Update
        .Close
    End With
MEX:
    Set SaleRS = Nothing
    Set PayRS = Nothing
'EH:
'    Msg = "Failed to save Sale data!" & vbLf & Err.Description
'    Err.Raise 456, , Msg
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.SaveSale"
End Sub

Public Sub SendExchange(GD As z_GetData)
    On Error GoTo errHandler
Dim Msg As String
Dim sFileName As String
    If rs Is Nothing Then
        Msg = "Recordset not initialized!"
        GoTo errHandler
    End If
    rs.MoveFirst
    If rs.EOF Then
        Msg = "Recordset is empty!"
    Else
        sFileName = sTillCode & "-" & Format(GD.GetNextFileNum(), "00000")
        sFileName = "\" & sFileName & ".pos"
        rs.Save sRoot & sOutBox & sFileName, adPersistADTG
    End If
    CreateBill
    'clear arrays
    ReDim sLines(0)
    'load empty RS
    MsgBox "WHat goes on here?"
    LoadRecordset
    sCustomer = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.SendExchange(GD)", GD
End Sub

Private Sub CreateBill()
    On Error GoTo errHandler
Dim sTemp As String
Dim sT2 As String
Dim sLine As String
Dim sT As String
Dim iQty As Integer
Dim iLine As Integer
'Dim rsLine As ADODB.Recordset
Dim rsPay As ADODB.Recordset
Const iW = 20
Dim NL As String
Dim i As Integer
Dim bDisc As Integer
Dim dTotal As Double
    
    NL = vbNewLine
    
    sBill = ""
    sT = Chr(vbKeyTab)
    
    'check if we got discount
    For i = 0 To UBound(sLines)
        If sLines(i).Disc > 0 Then
            bDisc = True
            Exit For
        End If
    Next i
    
    rs.MoveFirst
'    Set rsLine = rs("rsSaleLine").Value
'    rsLine.MoveFirst
    Set rsPay = rs("rsPayment").Value
    rsPay.MoveFirst
    
    sTemp = Centre("BOOKSHOP NAME", iW)
    
    sBill = sTemp & NL & NL
    sTemp = Centre("Tel: (021) 789 8787", iW)
    sBill = sBill & sTemp & NL
    sTemp = Centre("TAX INVOICE", iW)
    sBill = sBill & sTemp & NL
    sBill = sBill & Centre("VAT Reg. No.", iW) & NL
    sBill = sBill & Centre("0098700987", iW) & NL & NL
    
    sBill = sBill & Format(rs!Exch_Date, "dd mmm yyyy") & NL
    sBill = sBill & "Time:     " & Format(rs!Exch_Date, "hh:nn") & NL
    sBill = sBill & "Till:     " & rs!EXCH_Till_Code & NL
    sBill = sBill & "Cashier:  " & rs!EXCH_Operator_ID & NL & NL
    If sCustomer <> "" Then
        sBill = sBill & "Customer:" & NL
        sBill = sBill & sCustomer & NL & NL
    End If
    
    sTemp = "ISBN" & NL & "Title" & NL
    sBill = sBill & sTemp
    If bDisc Then
        sTemp = "Qty " & "Unit Disc" & sT & "Total" & NL
    Else
        sTemp = "Qty " & "Unit" & sT & "Total" & NL
    End If
    
    sLine = String(iW, "-") & NL
    sBill = sBill & sTemp & sLine
   
    For i = 0 To UBound(sLines)
        With sLines(i)
            sBill = sBill & .Code & NL
            sBill = sBill & Left(.Title, iW) & NL
            If bDisc Then
                sBill = sBill & .Qty & "  " & _
                    Format(.UnitPr, "R0.00") & " %" & .Disc & sT & _
                    Format(.LineTot, "R0.00") & NL
            Else
                sBill = sBill & .Qty & "  " & _
                    Format(.UnitPr, "R0.00") & sT & _
                    Format(.LineTot, "R0.00") & NL
            End If
            sBill = sBill & sLine
            iQty = iQty + .Qty
            dTotal = dTotal + .LineTot
        End With
    Next i
    sBill = sBill & sLine
    sBill = sBill & "Total Items: " & iQty & NL
    
    sTemp = "Amount Due:"
    sT2 = Format(dTotal, "R0.00")
    sBill = sBill & sTemp & Space(iW - Len(sTemp & sT2)) & sT2 & NL
    
    sTemp = sPaymentType
    sT2 = Format(rsPay!PAY_Tot_Received / 100, "R0.00")
    sBill = sBill & sTemp & Space(iW - Len(sTemp & sT2)) & sT2 & NL
    
    sTemp = "Change"
  '  sT2 = Format(rsPay!PAY_Change / 100, "R 0.00")
    sBill = sBill & sTemp & Space(iW - Len(sTemp & sT2)) & sT2 & NL
    
    sBill = sBill & String(iW, "=") & NL & NL
    sBill = sBill & Centre("Thank you for shopping", iW) & NL
    sBill = sBill & Centre("with us!", iW)
    
    Set rsPay = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.CreateBill"
End Sub

Private Function Centre(sText As String, iWidth As Integer) As String
    On Error GoTo errHandler
    If Len(sText) < iWidth Then
        Centre = Space((iWidth - Len(sText)) \ 2) & sText
    Else
        Centre = sText
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.Centre(sText,iWidth)", Array(sText, iWidth)
End Function

Public Property Get PrintInvoice() As String
    On Error GoTo errHandler
    PrintInvoice = sBill
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.PrintInvoice"
End Property


Public Property Get ServerPath() As String
    On Error GoTo errHandler
    ServerPath = sServerPath
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.ServerPath"
End Property
Public Property Get TillCode() As String
    On Error GoTo errHandler
    TillCode = sTillCode
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.TillCode"
End Property
Public Property Get guid() As String
    On Error GoTo errHandler
    guid = sGUID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.guid"
End Property
Public Property Let PaymentType(val As String)
    sPaymentType = val
End Property
Private Sub SaveServerPathToReg()
    On Error GoTo errHandler
    SaveSetting "POSClient", "QPath", "ServPath", sServerPath
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.SaveServerPathToReg"
End Sub

Private Function ReadServerPathFromReg() As Boolean
    On Error GoTo errHandler
Dim fs As New FileSystemObject

    sServerPath = GetSetting("POSClient", "QPath", "ServPath", "")
    If sServerPath <> "" Then
        If Not fs.FolderExists(sServerPath) Then
            ReadServerPathFromReg = False
        Else
            ReadServerPathFromReg = True
        End If
    Else
        ReadServerPathFromReg = False
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.ReadServerPathFromReg"
End Function

Private Function ReadServerPathFromFile() As Boolean
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim F As TextStream

       
    If Not fs.FileExists(oPC.SharedFolderRoot & "\SPath.dat") Then GoTo errHandler
    Set F = fs.OpenTextFile(oPC.SharedFolderRoot & "\SPath.dat", ForReading)
    sServerPath = F.ReadLine
    
    
    ReadServerPathFromFile = sServerPath <> ""
MEX:
    If Not F Is Nothing Then F.Close
    Set fs = Nothing
    Set F = Nothing
'EH:
'    GoTo MEX
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.ReadServerPathFromFile"
End Function
Private Sub SaveServerPathToFile()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim F As TextStream

       
    Set F = fs.CreateTextFile(oPC.SharedFolderRoot & "\SPath.dat", True)
    F.WriteLine (sServerPath)
    F.Close
MEX:
    Set fs = Nothing
    Set F = Nothing
'EH:
'    GoTo MEX
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.SaveServerPathToFile"
End Sub

Public Function SaveSetup(TillCode As String, ServerPath As String) As Boolean
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim sSQL As String
    
    If Right(ServerPath, 1) <> "\" Then
        sServerPath = ServerPath & "\"
    Else
        sServerPath = ServerPath
    End If
    sTillCode = Trim$(TillCode)
    sSQL = "SELECT * FROM tAppSettings WHERE tAppSettings.ID = 1"
    rs.Open sSQL, oPC.DBConn, adOpenDynamic, adLockOptimistic
    If rs.EOF Then rs.AddNew
    rs!ID = 1
    rs!TillCode = sTillCode
 '   rs!ServerPath = sServerPath
    rs.Update
    rs.Close
    Set rs = Nothing
    
    SaveSetup = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.SaveSetup(TillCode,ServerPath)", Array(TillCode, ServerPath)
End Function

Public Function ConfigureFromDB() As Boolean
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim sSQL As String
    
    
    sSQL = "SELECT * FROM tAppSettings WHERE tAppSettings.ID = 1"
    rs.Open sSQL, oPC.DBConn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        If Not IsNull(rs!TillCode) Then sTillCode = rs!TillCode
    End If
    rs.Close
    ConfigureFromDB = sTillCode <> ""
    Set rs = Nothing
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.ConfigureFromDB"
End Function

Public Function TestServerPath(sPathName As String) As Boolean
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fol As Folder

    'this has to be changed for net enabled version   << --------NOT NET ENABLED!
   ' TestServerPath = fs.FolderExists(sPathName)
    TestServerPath = fs.GetDrive(sPathName) > ""
        
    
    '----------------------------------------------------------------------------
    Set fs = Nothing
    Set fol = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.TestServerPath(sPathName)", sPathName
End Function

Public Function RequestRSFromServer() As Boolean
    On Error GoTo errHandler
Dim Msg As String
Dim fs As FileSystemObject
Dim F As TextStream
Dim sFileName As String
    
    If sTillCode = "" Then
        Msg = "Computer ID missing!"
        GoTo errHandler
    End If
'    'Try to load RS from previous requests
'    If GetRSFromServer(0) Then
'        RequestRSFromServer = True
'        Exit Function
'    End If
    
    'set up request for RS
    Set fs = New FileSystemObject
    sFileName = sTillCode & ".req"
    Set F = fs.CreateTextFile(sRoot & sOutBox & "\" & sFileName, True)
    If F Is Nothing Then
        Msg = "Can't create request file in server folder!"
        GoTo errHandler
    End If
    F.WriteLine "Send RS Data"
    F.WriteLine sTillCode
    F.WriteLine sSharedInboxName
    F.Close
    'send to server
    If fs.FolderExists(sServerPath) Then
        fs.CopyFile sRoot & sOutBox & "\" & sFileName, sServerPath & sFileName, True
        If Err Then
            Err.Clear
            Msg = "Can't send rs request file to folder!"
            GoTo errHandler
        End If
        fs.DeleteFile sRoot & sOutBox & "\" & sFileName
    Else
    
    End If
    'Try to load response
    RequestRSFromServer = GetRSFromServer(15)
    
    Set fs = Nothing
    Set F = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.RequestRSFromServer"
End Function

Private Function GetRSFromServer(Timeout As Single) As Boolean
    On Error GoTo errHandler
Dim bStop As Boolean
Dim siStart As Single
Dim siElapse As Single
Dim sNewFile As String
Dim tmpRS As ADODB.Recordset
    
    'set timeout length in seconds
    siElapse = Timeout
    siStart = Timer
    Do While Not bStop
        DoEvents
        sNewFile = Dir(sRoot & sInBox & "\*.rs")
        If sNewFile <> "" Then
            Set tmpRS = New ADODB.Recordset
            tmpRS.Open (sRoot & sInBox & "\" & sNewFile)
            Set rs = tmpRS
            'Save RS to disc
            rs.Save sRoot & "POS.rs", adPersistADTG
            bRSOK = True
            bStop = True
            'Delete RS file from In_Folder
            Dim fs As New FileSystemObject
            tmpRS.Close
            fs.DeleteFile sRoot & sInBox & "\" & sNewFile
            Set fs = Nothing
        End If
        'exit loop if is not found within timeout
        If Not bStop Then
            bStop = (Timer > siStart + siElapse)
        End If
    Loop
    GetRSFromServer = bRSOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.GetRSFromServer(Timeout)", Timeout
End Function


Private Sub oPoll_PollingStopped(Msg As String)
    On Error GoTo errHandler
    RaiseEvent PollingStopped(Msg)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.oPoll_PollingStopped(Msg)", Msg
End Sub
Public Sub StartPolling()
    On Error GoTo errHandler
    oPoll.StartPolling POLL_INTERVAL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.StartPolling"
End Sub
Public Sub StopPolling()
    oPoll.StopPolling ""
End Sub
