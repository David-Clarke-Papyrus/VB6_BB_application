VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event PollingStopped(Msg As String)
Event ConnectionStatus(i As Integer)
Event HasServer()
Event LostServer()
Event Counter(lngCount As Long)
Event UpdatingLocalDatabase(bON As Boolean, lngCount As Long)
Dim WithEvents oTimer As XTimer
Attribute oTimer.VB_VarHelpID = -1
Dim sInBox As String
Dim sOutBox As String
Dim sLocalServer As String
Dim sPBKSServerInbox As String
Dim sPBKSServerOutbox As String
Dim bPolling As Boolean
Dim bDBOK As Boolean


Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set oTimer = New XTimer
    Exit Sub
errHandler:
    LogSaveToFile "clsPOLL:Class_Initialize  " & Err.Description
    RaiseEvent ConnectionStatus(3)
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    
    TransmitAnySales
    UpdateClientFromServerFiles
    
    If bPolling Then
        oTimer.Enabled = False
    End If
    
    Set oTimer = Nothing
    
    oPC.dbCloseConnect
    Set oPC = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsPoll.Class_Terminate"
End Sub
Private Sub oTimer_Tick()
    On Error GoTo errHandler
    oTimer.Enabled = False
    
    TransmitAnySales
    UpdateClientFromServerFiles
    
    oTimer.Enabled = True
    Exit Sub
errHandler:
    LogSaveToFile "clsPOLL:oTimer_Tick  " & Err.Description
    RaiseEvent ConnectionStatus(3)
End Sub
Public Function OpenDB() As Boolean
    On Error GoTo errHandler
    Set oPC = New z_POSConnection
    bDBOK = oPC.dbConnect(Me.LocalServer) = 0
    OpenDB = bDBOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsPoll.SetDB()"
End Function
Public Function CloseDB()
    oPC.dbCloseConnect
End Function
Public Sub StartPolling(Interval As Long)
    On Error GoTo errHandler
    oTimer.Interval = Interval
    oTimer.Enabled = True
    bPolling = True
    Exit Sub
errHandler:

    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsPoll.StartPolling(Interval)", Interval
    
End Sub

Public Sub StopPolling()
    On Error GoTo errHandler
    oTimer.Enabled = False
    bPolling = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsPoll.StopPolling(Msg)"  ', Msg
End Sub

Public Property Let InBox(Val As String)
    sInBox = Val
End Property
Public Property Get InBox() As String
    InBox = sInBox
End Property
Public Property Let OutBox(Val As String)
    sOutBox = Val
End Property
Public Property Get OutBox() As String
    OutBox = sOutBox
End Property
Public Property Let LocalServer(Val As String)
    sLocalServer = Val
End Property
Public Property Get LocalServer() As String
    LocalServer = sLocalServer
End Property
Public Property Let PBKSServerInbox(Val As String)
    sPBKSServerInbox = Val
End Property
Public Property Get PBKSServerInbox() As String
    PBKSServerInbox = sPBKSServerInbox
End Property
Public Property Let PBKSServerOutbox(Val As String)
    sPBKSServerOutbox = Val
End Property
Public Property Get PBKSServerOutbox() As String
    PBKSServerOutbox = sPBKSServerOutbox
End Property
Public Property Get IsPolling() As Boolean
    IsPolling = bPolling
End Property
Public Property Get IsDBOK() As Boolean
    IsDBOK = bDBOK
End Property

Friend Sub TransmitAnySales()
    On Error GoTo errHandler
Dim xFiles() As String
Dim sFileName As String
Dim fs As New FileSystemObject
Dim i As Integer
Dim tmpErr As Long


    ReDim xFiles(0)
    sFileName = Dir(sOutBox & "\*.POS")
    Do While sFileName <> ""
        If xFiles(UBound(xFiles)) <> "" Then
            ReDim Preserve xFiles(UBound(xFiles) + 1)
        End If
        xFiles(UBound(xFiles)) = sFileName
        sFileName = Dir
    Loop
    For i = 0 To UBound(xFiles)
        If UCase(Right(xFiles(i), 4)) = ".POS" Then
        
            On Error Resume Next  'Dont crash here if network or server is down!
            fs.CopyFile sOutBox & "\" & xFiles(i), Me.PBKSServerInbox & "\" & xFiles(i)
            tmpErr = Err.Number
            If tmpErr = 0 Then  'Copied OK
                fs.DeleteFile sOutBox & "\" & xFiles(i), True
                On Error GoTo errHandler
                                     '''''''''''''    Err.Raise 12000, "Transmit sales", "test"
            Else
                Err.Clear
                On Error GoTo errHandler
                RaiseEvent ConnectionStatus(1)
            End If
        End If
    Next i
    
    Set fs = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsPoll.CheckSales"
End Sub


Friend Sub UpdateClientFromServerFiles()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim fol, F, fc
Dim rsFiles As ADODB.Recordset
Dim strFN As String
Dim strPath As String

'Need to protect against connection going down here
'--------------------------------------------------
    'Identify files to fetch
On Error GoTo SpecialError
    Set fol = fs.GetFolder(PBKSServerOutbox)
    Set fc = fol.Files
    For Each F In fc
        
        If UCase(Right(F.Name, 4)) = ".UPD" _
        Or UCase(Right(F.Name, 4)) = ".USD" _
        Or UCase(Right(F.Name, 4)) = ".UCD" _
        Or UCase(Right(F.Name, 4)) = ".UOD" _
        Or UCase(Right(F.Name, 4)) = ".UMD" Then
            fs.CopyFile F.Path, sInBox & "\", True
            If fs.FileExists(sInBox & "\" & F.Name) Then
                fs.DeleteFile PBKSServerOutbox & "\" & F.Name
            End If
        End If
    Next
'--------------------------------------------------
On Error GoTo errHandler
    'Update using local copies
    Set fol = fs.GetFolder(sInBox)
    Set fc = fol.Files
    
 'At this point we must sort files by date created so that we apply updates in the order that they were generated on the server
 'We use a recordset to hold the names and dates, sort it and read it
    Set rsFiles = New ADODB.Recordset
    rsFiles.Fields.Append "Filename", adVarChar, 50
    rsFiles.Fields.Append "Filepath", adVarChar, 150
    rsFiles.Fields.Append "Created", adDate
    rsFiles.Open
    For Each F In fc
        rsFiles.AddNew
        rsFiles.Fields("Filename") = F.Name
        rsFiles.Fields("Created") = F.DateCreated
        rsFiles.Fields("Filepath") = F.Path
        rsFiles.Update
    Next
    rsFiles.Sort = "Created ASC"
    If Not (rsFiles.EOF And rsFiles.BOF) Then
        rsFiles.MoveFirst
        Do While Not rsFiles.EOF
           strFN = rsFiles.Fields("Filename")
           strPath = rsFiles.Fields("Filepath")
    '    For Each F In fc
            If UCase(Right(strFN, 4)) = ".UPD" Then
                'Load product updates
                rs.Open strPath
                RaiseEvent UpdatingLocalDatabase(True, rs.RecordCount)
                If SaveProductUpdate(rs) Then
                    fs.DeleteFile strPath
                End If
                RaiseEvent UpdatingLocalDatabase(False, 0)
            ElseIf UCase(Right(strFN, 4)) = ".USD" Then
                'Load Staff Member updates
                rs.Open strPath
                RaiseEvent UpdatingLocalDatabase(True, rs.RecordCount)
                If SaveStaffUpdate(rs) Then
                    fs.DeleteFile strPath
                End If
                RaiseEvent UpdatingLocalDatabase(False, 0)
            ElseIf UCase(Right(strFN, 4)) = ".UCD" Then
                'Load Customer updates
                rs.Open strPath
                RaiseEvent UpdatingLocalDatabase(True, rs.RecordCount)
                If SaveCustomerUpdate(rs) Then
                    fs.DeleteFile strPath
                End If
                RaiseEvent UpdatingLocalDatabase(False, 0)
            ElseIf UCase(Right(strFN, 4)) = ".UOD" Then
                'Load Customer order updates
                rs.Open strPath
                RaiseEvent UpdatingLocalDatabase(True, rs.RecordCount)
                If SaveCustomerOrderUpdate(rs) Then
                    fs.DeleteFile strPath
                End If
                RaiseEvent UpdatingLocalDatabase(False, 0)
            ElseIf UCase(Right(strFN, 4)) = ".UMD" Then
                'Load marketing updates
                rs.Open strPath
                RaiseEvent UpdatingLocalDatabase(True, rs.RecordCount)
                If SaveMarketingUpdate(rs) Then
                    fs.DeleteFile strPath
                End If
                RaiseEvent UpdatingLocalDatabase(False, 0)
            End If
            rsFiles.MoveNext
        Loop
    'next
    End If
    Set rs = Nothing
    Set fc = Nothing
    Set fol = Nothing
    Set fs = Nothing
    
    rsFiles.Close
    Set rsFiles = Nothing
    
    Exit Sub
SpecialError:
    RaiseEvent LostServer
    Err.Clear
    Exit Sub
errHandler:
    ErrPreserve
''''''    If Not fs.FolderExists(PBKSServerOutbox) Then
''''''
''''''    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsPoll.CheckUpdates"
End Sub

Private Function SaveProductUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim lngCnt As Long
    
    rs.MoveFirst
    lngCnt = 1
    Do While Not rs.EOF
        lngCnt = lngCnt + 1
        If lngCnt Mod 100 = 0 Then
            RaiseEvent Counter(lngCnt)
        End If
        Set cmd = New ADODB.Command
        cmd.CommandText = "dbo.sp_InsertProductUpdateToFD "
        cmd.CommandType = adCmdStoredProc
        
        Set par = cmd.CreateParameter("@PID", adGUID, , , rs!PRU_P_ID)
        cmd.Parameters.Append par
        Set par = Nothing
        Set par = cmd.CreateParameter("@CODE", adVarChar, adParamInput, 50, FNS(rs!PRU_Code))
        cmd.Parameters.Append par
        Set par = Nothing
        Set par = cmd.CreateParameter("@EAN", adVarChar, adParamInput, 50, FNS(rs!PRU_EAN))
        cmd.Parameters.Append par
        Set par = Nothing
        Set par = cmd.CreateParameter("@PUBLISHER", adVarChar, adParamInput, 50, FNS(rs!PRU_Publisher))
        cmd.Parameters.Append par
        Set par = Nothing
        Set par = cmd.CreateParameter("@SERIESTITLE", adVarChar, adParamInput, 225, FNS(rs!PRU_SeriesTitle))
        cmd.Parameters.Append par
        Set par = Nothing
        Set par = cmd.CreateParameter("@AUTHOR", adVarChar, adParamInput, 225, FNS(rs!PRU_MainAuthor))
        cmd.Parameters.Append par
        Set par = Nothing
        Set par = cmd.CreateParameter("@TITLE", adVarChar, adParamInput, 225, FNS(rs!PRU_Title))
        cmd.Parameters.Append par
        Set par = Nothing
        Set par = cmd.CreateParameter("@SP", adInteger, adParamInput, , FNN(rs!PRU_SP))
        cmd.Parameters.Append par
        Set par = Nothing
        Set par = cmd.CreateParameter("@VATRATE", adNumeric, adParamInput, 10, FNDBL(rs!PRU_VATRATE))
        par.Precision = 8
        par.NumericScale = 2
        cmd.Parameters.Append par
        Set par = Nothing
        Set par = cmd.CreateParameter("@PTID", adInteger, adParamInput, , FNN(rs!PRU_PTID))
        cmd.Parameters.Append par
        Set par = Nothing
        Set par = cmd.CreateParameter("@SECID", adInteger, adParamInput, , FNN(rs!PRU_SECID))
        cmd.Parameters.Append par
        Set par = Nothing
        
        cmd.ActiveConnection = oPC.DBConn
        cmd.Execute
        
        Set cmd = Nothing
        rs.MoveNext
        DoEvents
    Loop


    SaveProductUpdate = True
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsPoll.SaveProductUpdate(rs)", rs
End Function

Private Function SaveStaffUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim NewRS As New ADODB.Recordset
Dim sSQL As String
Dim sName As String
Dim i As Integer

    rs.MoveFirst
    Do While Not rs.EOF
        sSQL = "SELECT * FROM tStaffMembers WHERE tStaffMembers.SM_ID =" & rs!SMU_ID
        NewRS.LockType = adLockOptimistic
        NewRS.CursorType = adOpenDynamic
        Set NewRS.ActiveConnection = oPC.DBConn
        NewRS.Open sSQL  ', adOpenDynamic, adLockPessimistic
        
        If NewRS.EOF Then
            NewRS.AddNew
        End If
        If Not IsNull(rs!SMU_ID) Then NewRS!SM_ID = rs!SMU_ID
        If Not IsNull(rs!SMU_NAME) Then NewRS!SM_Name = Trim$(rs!SMU_NAME)
        If Not IsNull(rs!SMU_Role) Then NewRS!SM_Role = Trim$(rs!SMU_Role)
        If Not IsNull(rs!SMU_Telephone) Then NewRS!SM_Telephone = Trim$(rs!SMU_Telephone)
        If Not IsNull(rs!SMU_Mobile) Then NewRS!SM_Mobile = Trim$(rs!SMU_Mobile)
        If Not IsNull(rs!SMU_Password) Then NewRS!SM_Password = Trim$(rs!SMU_Password)
        If Not IsNull(rs!SMU_Level) Then NewRS!SM_Level = rs!SMU_Level
        If Not IsNull(rs!SMU_Shortname) Then NewRS!SM_SHortname = rs!SMU_Shortname
        sName = Trim$(rs!SMU_NAME)
        i = 1
DoUpdate:
        NewRS.Update
        If Err = -2147217887 Then
          NewRS!SM_Code = Left(NewRS!SM_Code, 3) & CStr(i)
          i = i + 1
          Err.Clear
          GoTo DoUpdate
        ElseIf Err <> 0 Then
          GoTo errHandler
        End If
        NewRS.Close
        DoEvents
        rs.MoveNext
    Loop
 '   oPC.DBConn.CommitTrans
    SaveStaffUpdate = True
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set NewRS = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsPoll.SaveStaffUpdate(rs)", rs
End Function

Private Function SaveCustomerUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim NewRS As New ADODB.Recordset
Dim sSQL As String
Dim sName As String
Dim i As Integer
    rs.MoveFirst
    Do While Not rs.EOF
        sSQL = "SELECT * FROM tCustomer WHERE tCustomer.Customer_ID =" & rs!CU_ID
        NewRS.LockType = adLockOptimistic
        NewRS.CursorType = adOpenDynamic
        Set NewRS.ActiveConnection = oPC.DBConn
        NewRS.Open sSQL
        
        If UCase(FNS(rs!CU_TYPE)) = "X" Then
            NewRS.Delete
        Else
            If NewRS.EOF Then
                NewRS.AddNew
            End If
            NewRS!Customer_ID = FNN(rs!CU_ID)
            NewRS!C_Name = FNS((rs!CU_Name))
            NewRS!C_DefaultDiscount = FNN(rs!CU_DefaultDiscount)
            NewRS!C_Acno = FNS(rs!CU_Acno)
            NewRS!C_Initials = Left(FNS(rs!CU_Initials), 8)
            NewRS!C_Title = FNS(rs!CU_Title)
            NewRS!C_Phone = FNS(rs!CU_Phone)
            NewRS!C_VATABLE = FNN(rs!CU_VATABLE)
            NewRS!C_BALANCE = FNN(rs!CU_BALANCE)
            NewRS!C_Type = FNS(rs!CU_TYPE)
        End If
        NewRS.Update
        NewRS.Close
        DoEvents
        rs.MoveNext
    Loop
    SaveCustomerUpdate = True
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set NewRS = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsPoll.SaveCustomerUpdate(rs)", rs
End Function

Private Function SaveCustomerOrderUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim NewRS As New ADODB.Recordset
Dim sSQL As String
Dim sName As String
Dim i As Integer
    rs.MoveFirst
    Do While Not rs.EOF
        sSQL = "SELECT * FROM tCOL WHERE COL_COLID = " & rs!COU_COLID
        NewRS.LockType = adLockOptimistic
        NewRS.CursorType = adOpenDynamic
        Set NewRS.ActiveConnection = oPC.DBConn
        NewRS.Open sSQL  ', adOpenDynamic, adLockPessimistic
        
        If NewRS.EOF Then
            NewRS.AddNew
        End If
        NewRS!COL_COLID = FNN(rs!COU_COLID)
        NewRS!COL_TPID = FNN(rs!COU_TPID)
        NewRS!COL_TRID = FNN(rs!COU_TRID)
        NewRS!COL_Date = FND(rs!COU_Date)

        NewRS!COL_CODE = FNS(rs!COU_CODE)
        NewRS!COL_PID = FNS(rs!COU_PID)
        NewRS!COL_Qty = FNN(rs!COU_QTY)
        NewRS!COL_QTYDISPATCHED = FND(rs!COU_QTYDISPATCHED)

        NewRS!COL_PRICE = FNN(rs!COU_PRICE)
        NewRS!COL_DISCOUNTRATE = FNN(rs!COU_DISCOUNTRATE)
        NewRS!COL_DEPOSIT = FNN(rs!COU_DEPOSIT)
        NewRS!COL_DEPOSITSTATUS = FNS(rs!COU_DEPOSITSTATUS)

        NewRS!COL_DELETE = FNS(rs!COU_DOCSTATUS)
DoUpdate:
        NewRS.Update
        NewRS.Close
        DoEvents
        rs.MoveNext
    Loop
    
    SaveCustomerOrderUpdate = True
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set NewRS = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsPoll.SaveCustomerOrderUpdate(rs)", rs
End Function


Private Function SaveMarketingUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim NewRS As New ADODB.Recordset
Dim sSQL As String
Dim sName As String
Dim i As Integer
    rs.MoveFirst
    Do While Not rs.EOF
        sSQL = "SELECT * FROM tMarketing WHERE M_PT_ID = " & rs!MC_PT_ID & " AND M_SECTION_ID = " & rs!MC_SECTION_ID
        NewRS.LockType = adLockOptimistic
        NewRS.CursorType = adOpenDynamic
        Set NewRS.ActiveConnection = oPC.DBConn
        NewRS.Open sSQL  ', adOpenDynamic, adLockPessimistic
        If rs!MC_TYPE = "DEL" Then   'must delete record
            If Not NewRS.EOF Then
                NewRS.Delete
                NewRS.Update
            End If
        Else
            If NewRS.EOF Then
                NewRS.AddNew
            End If
            NewRS!M_PT_ID = FNN(rs!MC_PT_ID)
            NewRS!M_SECTION_ID = FNN(rs!MC_SECTION_ID)
            NewRS!M_LOYALTYDISCOUNT = FNN(rs!MC_LOYALTYDISCOUNT)
            NewRS!M_DISCOUNT = FND(rs!MC_DISCOUNT)
    
            NewRS!M_DESCRIPTION = FNS(rs!MC_DESCRIPTION)
            NewRS!M_NODISCOUNTALLOWABLE = FNS(rs!MC_NODISCOUNTALLOWABLE)
            NewRS!M_IDENTIFYCUSTOMER = FNN(rs!MC_IDENTIFYCUSTOMER)
            NewRS!M_ACTIVE = FNN(rs!MC_ACTIVE)
DoUpdate:
            NewRS.Update
            NewRS.Close
        End If
        DoEvents
        rs.MoveNext
    Loop
    SaveMarketingUpdate = True
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set NewRS = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsPoll.SaveMarketingUpdate(rs)", rs
End Function

