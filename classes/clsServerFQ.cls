VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServerFQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type tClientList
    Name As String
    pathname As String
    OutBox As String
End Type


Dim ClientList() As tClientList

Public Event FileProcessed(FileType As Integer) '0 = Other, 1 = Sale, 2 = Update

Dim oDB As z_DatabasePersist
Dim bDBOK As Boolean
Dim bFoldersOK As Boolean


Dim sOutBoxRoot As String
Dim sRoot As String
Dim sInBox As String
Dim sInBoxSharedName As String

Private Sub Class_Initialize()
    Set oDB = New z_DatabasePersist
    bDBOK = oDB.dbConnect("", "") = 0
    ReDim ClientList(0)
    CheckFolders
End Sub

Private Sub Class_Terminate()
Dim r As Long
    If bDBOK Then
        r = oDB.dbCloseConnect()
    End If
    bDBOK = False
    Set oDB = Nothing
End Sub

Private Function CheckFolders()
Dim fs As New FileSystemObject
Dim oFShare As New clsFolderShare
Dim sSharedName As String

    sRoot = "D:\PBKS\POS_SVR\"
    sInBox = sRoot & "InBox"
    sOutBoxRoot = sRoot & "ClientOut"
    
    sSharedName = "POSServer_InBox"
    
    If Not fs.FolderExists(sInBox) Then
        If Not oFShare.CreateFolder(sInBox, True, sSharedName) Then Exit Function
    Else
        'make sure folder is set as shared
'        If Not oFShare.CheckShared(oFShare.GetCompName, sInBox) Then
            oFShare.SetShared oFShare.GetCompName, sSharedName, sInBox
'        End If
    End If
    If Not fs.FileExists(sOutBoxRoot) Then
        If Not oFShare.CreateFolder(sOutBoxRoot, False) Then Exit Function
    End If
    sInBoxSharedName = "\\" & oFShare.GetCompName & "\" & sSharedName
    bFoldersOK = True

End Function
Public Property Get SharedInboxName() As String
    SharedInboxName = sInBoxSharedName
End Property

Public Property Get FoldersOK() As Boolean
    FoldersOK = bFoldersOK
End Property
Public Property Get DBConnectionOK() As Boolean
    DBConnectionOK = bDBOK
End Property
Public Property Get ClientListCount() As Integer
    If ClientList(UBound(ClientList)).Name <> "" Then
        ClientListCount = UBound(ClientList)
    Else
        ClientListCount = -1
    End If
End Property

Public Function ClientName(Index As Integer) As String
    If Index <= UBound(ClientList) Then
        ClientName = ClientList(Index).Name
    End If
End Function
Public Function ClientPathName(Index As Integer) As String
    If Index <= UBound(ClientList) Then
        ClientPathName = ClientList(Index).pathname
    End If
End Function

Public Function LoadClientList() As Integer
Dim rs As New ADODB.Recordset
Dim sSQL As String
Dim i As Integer
    
    sSQL = "SELECT * FROM tPOSClient"
    ReDim ClientList(0)
    On Error GoTo EH
    rs.Open sSQL, oDB.dbConn, adOpenDynamic, adLockReadOnly
    Do While Not rs.EOF
        If ClientList(UBound(ClientList)).Name <> "" Then
            ReDim Preserve ClientList(UBound(ClientList) + 1)
        End If
        With ClientList(UBound(ClientList))
            If Not IsNull(rs!TillCode) Then .Name = rs!TillCode
            If Not IsNull(rs!pathname) Then .pathname = rs!pathname
            If Not IsNull(rs!OutBox) Then .OutBox = rs!OutBox
        End With
        i = i + 1
        rs.MoveNext
    Loop
    LoadClientList = i
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Exit Function
EH:
    LoadClientList = -99
    GoTo MEX
    Resume
End Function

Public Function PollDB() As Boolean
'Dim oDP As New z_DatabasePersist
Dim rs As ADODB.Recordset
Dim sSQL As String

    On Error GoTo EH
    
    
    oDB.dbConn.BeginTrans
    
    sSQL = "SELECT * FROM tProdUpdates"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockBatchOptimistic
    rs.Open sSQL, oDB.dbConn
    'DISCONNECT RECORDSET
    rs.ActiveConnection = Nothing
    
    If Not rs.EOF Then
        If PrepareUpdate(rs) Then
            sSQL = "DELETE tProdUpdates"
            oDB.dbConn.Execute sSQL
            
            PollDB = True
            RaiseEvent FileProcessed(2)
        Else
            GoTo EH
        End If
    End If
    oDB.dbConn.CommitTrans
MEX:
'   oDP.dbCloseConnect
'   Set oDP = Nothing
    If rs.State = adStateOpen Then rs.Close
   Set rs = Nothing
    
Exit Function
EH:
    oDB.dbConn.RollbackTrans
    GoTo MEX
End Function

Private Function PrepareUpdate(rs As ADODB.Recordset) As Boolean
Dim i As Integer
Dim oRS As ADODB.Recordset
Dim fs As FileSystemObject
Dim sFile As String

    Err.Clear
    On Error Resume Next
    Set fs = New FileSystemObject
    Set oRS = rs
    oRS.MoveFirst
    sFile = oRS!Update_GUID & ".upd"
    oRS.Save sOutBoxRoot & "\" & sFile, adPersistADTG
    For i = 0 To UBound(ClientList)
        With ClientList(i)
            If Not fs.FolderExists(sOutBoxRoot & "\" & .OutBox) Then
                fs.CreateFolder sOutBoxRoot & "\" & .OutBox
            End If
            fs.CopyFile sOutBoxRoot & "\" & sFile, sOutBoxRoot & "\" & .OutBox & "\"
        End With
    Next i
    fs.DeleteFile sOutBoxRoot & "\" & sFile
    PrepareUpdate = Err = 0
    Err.Clear

End Function
Public Sub PollForNewClient()
Dim fs As New FileSystemObject
Dim f As TextStream
Dim sFile As String
Dim sTillCode As String
Dim sPath As String
Dim xFiles() As String
Dim i As Integer, k As Integer

    On Error Resume Next
    ReDim xFiles(0)
    sFile = Dir(sInBox & "\*.*") '"\*_Reg.dat")
    Do While sFile <> ""
        If xFiles(UBound(xFiles)) <> "" Then
            ReDim Preserve xFiles(UBound(xFiles) + 1)
        End If
        xFiles(UBound(xFiles)) = sInBox & "\" & sFile
        sFile = Dir
    Loop
    For k = 0 To UBound(xFiles)
        If UCase(Right(xFiles(k), 7)) = "REG.DAT" Then
            Set f = fs.OpenTextFile(xFiles(k))
            sTillCode = f.ReadLine
            sPath = f.ReadLine
            f.Close
            
            i = AddClient(sTillCode, sPath)
            If i >= 0 Then
                fs.DeleteFile (xFiles(k))
                'Create and send confirmation file
                Set f = fs.CreateTextFile(sOutBoxRoot & "\RegOK.dat")
                f.WriteLine sInBoxSharedName
                f.Close
                fs.MoveFile sOutBoxRoot & "\RegOK.dat", sPath & "\RegOK.dat"
                If Err Then Err.Clear
            End If
            
        ElseIf UCase(Right(xFiles(k), 4)) = ".REQ" Then
            Dim ClientPath As String
            Dim TillCode  As String
            Dim sTmp As String
            
            Set f = fs.OpenTextFile(xFiles(k))
            sTmp = f.ReadLine
            TillCode = f.ReadLine
            ClientPath = f.ReadLine
            f.Close
            If sTmp = "Send RS Data" Then
                If SendRSFile(TillCode, ClientPath) Then
                    fs.DeleteFile xFiles(k), True
                    RaiseEvent FileProcessed(0)
                End If
            End If
        End If
    Next k
        
    Set fs = Nothing
    Set f = Nothing
End Sub

Private Function ClientExists(sTillCode As String, sPath As String) As String
Dim i As Integer
    For i = 0 To UBound(ClientList)
        If UCase(sPath) = UCase(ClientList(i).pathname) Then
            ClientExists = "PA"
            Exit For
        ElseIf UCase(sTillCode) = UCase(ClientList(i).Name) Then
            ClientExists = "TC"
            Exit For
        End If
    Next i
End Function

Friend Function AddClient(Name As String, pathname As String) As Integer
Dim rs As New ADODB.Recordset
Dim sSQL As String
Dim oFS As New clsFolderShare
Dim sRes As String
Dim sOldPath As String

    sRes = ClientExists(Name, pathname)
    If sRes = "TC" Then
        sSQL = "SELECT * FROM tPOSClient WHERE tPOSClient.TillCode ='" & UCase(Name) & "'"
    ElseIf sRes = "PA" Then
        sSQL = "SELECT * FROM tPOSClient WHERE tPOSClient.PathName ='" & UCase(pathname) & "'"
    Else
        sSQL = "SELECT * FROM tPOSClient WHERE tPOSClient.TillCode =''"
    End If
'    ReDim ClientList(0)
    On Error GoTo EH
    oDB.dbConn.BeginTrans
    rs.Open sSQL, oDB.dbConn, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        rs.AddNew
    End If
    rs!TillCode = UCase(Name)
    If Not IsNull(rs!OutBox) Then sOldPath = rs!OutBox
    rs!pathname = UCase(pathname)
    rs!OutBox = "\" & rs!TillCode & "_OutBox"
    rs!CreateDate = Now
    rs.Update
    
    'delete old path if nessesary
    If sOldPath <> "" And sOldPath <> rs!OutBox Then DeletePath (sOutBoxRoot & sOldPath)
    
    'create new OutBox for new client
    If Not oFS.CreateFolder(sOutBoxRoot & rs!OutBox, False) Then GoTo EH
    oDB.dbConn.CommitTrans
    
    
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    LoadClientList
    Exit Function
EH:
    
    
    oDB.dbConn.RollbackTrans
    AddClient = -99
    GoTo MEX
End Function

Private Sub DeletePath(sPath As String)
Dim fs As New FileSystemObject
Dim sFile As String

    'check if folder is empty
    sFile = Dir(sPath & "\*.*")
    If sFile <> "" Then
'        if msgbox("The folder:" & vblf & sPath & vblf & "is not empty!" & vblf & _
'                  "Delete it anyway?", vbyesno, "
    Else
        fs.DeleteFolder sPath
    End If
    
    
End Sub

Public Sub DeleteClient(ClientName As String)
Dim i As Integer
Dim sSQL As String
Dim msg As String
    
    On Error GoTo EH
    For i = 0 To UBound(ClientList)
        If ClientList(i).Name = ClientName Then
            sSQL = "DELETE FROM tPOSClient WHERE tPOSClient.TillCode ='" & ClientName & "'"
            oDB.dbConn.Execute sSQL
            'reload client list
            i = LoadClientList
            If i = 0 Then
                msg = "Client list is empty!"
            ElseIf i = -99 Then
                msg = "Failed to load client list!"
            End If
        End If
    Next i
    Exit Sub
EH:
    If msg = "" Then
        msg = "Failed to delete client: " & ClientName & vbLf & _
              "from list."
    End If
    MsgBox msg, vbOKOnly + vbCritical, "Delete Client"
End Sub
Public Property Get DBName() As String
    DBName = oDB.DBName
End Property

Public Sub PollSales()
Dim rs As New ADODB.Recordset
Dim sFile As String
Dim xFiles() As String
Dim fs As New FileSystemObject
Dim i As Integer, k As Integer

    On Error Resume Next
    ReDim xFiles(0)
    sFile = Dir(sInBox & "\*.*") '"\*_Reg.dat")
    Do While sFile <> ""
        If xFiles(UBound(xFiles)) <> "" Then
            ReDim Preserve xFiles(UBound(xFiles) + 1)
        End If
        xFiles(UBound(xFiles)) = sInBox & "\" & sFile
        sFile = Dir
    Loop
    For k = 0 To UBound(xFiles)
        If UCase(Right(xFiles(k), 4)) = ".POS" Then
            rs.Open xFiles(k)
            If SaveSaleData(rs) Then
                fs.DeleteFile (xFiles(k)), True
                RaiseEvent FileProcessed(1)
            Else
                
            End If
        End If
    Next k
        
    Set fs = Nothing
End Sub

Private Function SaveSaleData(rs As ADODB.Recordset) As Boolean
Dim rsSL As ADODB.Recordset
Dim rsP As ADODB.Recordset
Dim msg As String
''''''''''    If Not oDB.IsConnected Then
''''''''''        If oDB.dbConnect("", "") <> 0 Then
''''''''''            msg = "Can't connect to database!"
''''''''''            GoTo EH
''''''''''        End If
''''''''''    End If
    On Error GoTo EH
    oDB.dbConn.BeginTrans
    rs.ActiveConnection = oDB.dbConn
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
    oDB.dbConn.CommitTrans
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
        oDB.dbConn.RollbackTrans
        SaveSaleData = True
        Err.Clear
        GoTo MEX
    End If
    If msg <> "" Then
        MsgBox "Can't save Sales Data!" & vbLf & msg
    Else
        oDB.dbConn.RollbackTrans
        MsgBox "Can't save Sales Data!" & vbLf & Err.Description
    End If

End Function
'~~~~~~~~~~~HANDLING PRODUCT UPDATES~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub SendUpdates()
Dim sFile As String
Dim xFile() As String
Dim i As Integer, k As Integer
Dim fs As New FileSystemObject

    On Error Resume Next
    ReDim xFile(0)
    For i = 0 To UBound(ClientList)
        With ClientList(i)
            sFile = Dir(sOutBoxRoot & "\" & .OutBox & "\*.UPD")
            
            Do While sFile <> ""
                If xFile(UBound(xFile)) <> "" Then
                    ReDim Preserve xFile(UBound(xFile) + 1)
                End If
                xFile(UBound(xFile)) = sFile
                sFile = Dir
            Loop
            For k = 0 To UBound(xFile)
                If xFile(k) <> "" Then
                    fs.CopyFile sOutBoxRoot & "\" & .OutBox & "\" & xFile(k), .pathname & "\" & xFile(k)
                    If Not Err Then
                        fs.DeleteFile sOutBoxRoot & "\" & .OutBox & "\" & xFile(k), True
                    Else
                        Err.Clear
                    End If

                End If
            Next k
        End With
        
    Next i
    Exit Sub
EH:

End Sub

Private Function SendRSFile(TillCode As String, ClientPath As String) As Boolean
Dim fs As New FileSystemObject
Dim rs As ADODB.Recordset
Dim sName As String

    If fs.FolderExists(ClientPath) Then
        sName = TillCode & ".rs"
        Set rs = CreateDisconnRecSet(False)
        If Not rs Is Nothing Then
            rs.Save sOutBoxRoot & "\" & sName, adPersistADTG
            On Error Resume Next
            fs.CopyFile sOutBoxRoot & "\" & sName, ClientPath & "\" & sName
            If Not Err Then
                fs.DeleteFile sOutBoxRoot & "\" & sName
                SendRSFile = True
            Else
                Err.Clear
            End If
        End If
        
    End If
    
    Set fs = Nothing
    
End Function

Public Function CreateDisconnRecSet(SaveToDisc As Boolean) As ADODB.Recordset
Dim oShapeDB As New z_DatabasePersist
Dim rs As New ADODB.Recordset
Dim sSQL As String
Dim msg As String

    If oShapeDB.dbConnect("sa", "CITROEN", "MSDataShape") <> 0 Then
        msg = "Failed to create database connection!"
        GoTo EH
    End If
    sSQL = "SHAPE {SELECT * FROM tExchange WHERE tExchange.Exchange_GUID = Null} " & _
           "APPEND ({SELECT * FROM tCSL WHERE tCSL.CSL_ExchangeID = Null} " & _
           "RELATE Exchange_GUID TO CSL_ExchangeID) AS rsSaleLine, " & _
           "({SELECT * FROM tPayment WHERE tPayment.ExchangeID = Null} " & _
           "RELATE Exchange_GUID TO ExchangeID) AS rsPayment"

'    sSQL = "SELECT * FROM tblCashSaleLine WHERE ExchangeID = Null"
    On Error GoTo EH
    rs.ActiveConnection = oShapeDB.dbConn
    rs.CursorLocation = adUseClient
    rs.Open sSQL, , adOpenStatic, adLockBatchOptimistic
    Set rs.ActiveConnection = Nothing
    Set CreateDisconnRecSet = rs
    
'    If SaveToDisc Then SaveRSToFile rs
    
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


