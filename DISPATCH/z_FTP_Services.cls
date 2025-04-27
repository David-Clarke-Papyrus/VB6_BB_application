VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_FTP_Services"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim fs As New FileSystemObject
Dim ADOConn As New ADODB.Connection

Public Function UploadFile(OwnOrSupplier As String, f As String, pPath As String, pErr As Long) As Boolean
    On Error GoTo errHandler
Dim FTPAddress As String
Dim FTPUser As String
Dim FTPPassword As String
Dim FTPFolder As String
Dim FTP1 As New FTPClass
Dim bResult As Boolean
Dim sFilenameForDestination As String

    GetFTPAddresss OwnOrSupplier, f, FTPAddress, FTPUser, FTPPassword, FTPFolder
    
    If FTPAddress > "" Then
        Check FTP1.OpenFTP(FTPAddress, FTPUser, FTPPassword, True), EXC_GENERAL, "Opening FTP"
        If FTPFolder > "" Then
            Check FTP1.SetCurrentFolder(FTPFolder), EXC_GENERAL, "Opening FTP"
        End If
        bResult = FTP1.PutFile(f, pPath, True)
        InsertToTransmissionTable fs.GetBaseName(f), "EDI using FTP", FTPAddress, IIf(bResult, "OK", "Failed"), err.LastDllError
        FTP1.CloseFTP
    Else
        If ExtractDocCodeFromFilename(pPath) > "" Then
            UpdateLog "Cant get FTP address for Acno - : " & ExtractDestinationFromFilename(pPath), ExtractDocCodeFromFilename(pPath)
        End If
        InsertToTransmissionTable fs.GetBaseName(f), "EDI using FTP", FTPAddress, "Failed", "Cant get FTP address for Acno - : " & ExtractDestinationFromFilename(pPath)
    End If
    
    UploadFile = True
    
    Exit Function
errHandler:
    ErrPreserve
    If err <> 0 Then
        If ExtractDocCodeFromFilename(pPath) > "" Then
            UpdateLog Description & " " & ExtractDestinationFromFilename(pPath), ExtractDocCodeFromFilename(pPath)
        End If
        InsertToTransmissionTable fs.GetBaseName(f), "EDI using FTP", FTPAddress, "Failed", Description
        UploadFile = False
        pErr = err
        Exit Function
    End If
End Function

Private Sub GetFTPAddresss(OwnOrSupplier As String, SupplierAcno As String, FTP_Site As String, FTP_User As String, FTPPassword As String, FTPInitialFolder As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim OpenResult As Integer
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    If OwnOrSupplier = "OWN" Then
        cmd.CommandText = "GetOwnFTPAddress"
    Else
        cmd.CommandText = "GetPOFTPAddress"
    End If
    
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Acno", adVarChar, , 50, SupplierAcno)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@FTPAddress", adVarChar, adParamOutput, 200)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@FTPUser", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@FTPPassword", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@FTPInitialFolder", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.Execute
    FTP_Site = FNS(cmd.Parameters(1))
    FTP_User = FNS(cmd.Parameters(2))
    FTPPassword = FNS(cmd.Parameters(3))
    FTPInitialFolder = FNS(cmd.Parameters(4))
    Set cmd = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTP_Services.GetFTPAddresss(OwnOrSupplier,SupplierAcno,FTP_Site,FTP_User,FTPPassword," & _
        "FTPInitialFolder)", Array(OwnOrSupplier, SupplierAcno, FTP_Site, FTP_User, FTPPassword, _
         FTPInitialFolder)
End Sub

Private Sub InsertToTransmissionTable(Docno As String, sMode As String, FTP_Site As String, Success As String, pErrors As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
    Set cmd = New ADODB.Command
    cmd.CommandText = "InsertToTransmissionTable"
    cmd.CommandType = adCmdStoredProc
   
    Set par = cmd.CreateParameter("@DOC", adVarChar, adParamInput, 50, Docno)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@DEST", adVarChar, adParamInput, 200, FTP_Site)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@Mode", adVarChar, adParamInput, 20, sMode)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@RES", adVarChar, adParamInput, 20, Success)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@Errors", adVarChar, adParamInput, 300, pErrors)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.Execute
    Set cmd = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTP_Services.InsertToTransmissionTable(Docno,sMode,FTP_Site,Success,pErrors)", _
         Array(Docno, sMode, FTP_Site, Success, pErrors)
End Sub

Private Function ExtractDestinationFromFilename(pIn As String) As String
    On Error GoTo errHandler
    ExtractDestinationFromFilename = Left(fs.GetBaseName(pIn), InStr(1, fs.GetBaseName(pIn), "_") - 1)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTP_Services.ExtractDestinationFromFilename(pIn)", pIn
End Function

Private Function ExtractDocCodeFromFilename(pIn As String) As String
    On Error GoTo errHandler
    ExtractDocCodeFromFilename = Right(fs.GetBaseName(pIn), Len(fs.GetBaseName(pIn)) - InStr(1, fs.GetBaseName(pIn), "_"))
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTP_Services.ExtractDocCodeFromFilename(pIn)", pIn
End Function

Private Sub UpdateLog(pMsg As String, pDocCode As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
    Set cmd = New ADODB.Command
    cmd.CommandText = "UpdateTransmissionLog"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@DOCCODE", adVarChar, , 15, pDocCode)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Msg", adVarChar, adParamInput, 1000, pMsg)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.Execute
    
    Set cmd = Nothing

    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTP_Services.UpdateLog(pMsg,pDocCode)", Array(pMsg, pDocCode)
End Sub

