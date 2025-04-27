VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Export"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim FTP1 As New FTPClass
Dim oTF As z_TextFile
Dim zip As FathZIPCtrl



Public Sub Component(pTF As z_TextFile)
    Set oTF = pTF
End Sub

Public Sub ControlPOSServerPolling(bON As Boolean)
Dim QI As MSMQQueueInfo
Dim POSmsg As MSMQMessage
Dim QPOS As MSMQQueue

    Set QI = New MSMQQueueInfo
    QI.FormatName = "DIRECT=OS:" & oPC.ServerMachineName & "\Private$\QPOS"
    Set QPOS = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)


    Set POSmsg = New MSMQMessage
    POSmsg.Delivery = MQMSG_DELIVERY_RECOVERABLE
    If bON = True Then
        POSmsg.Label = "STARTPOLLING"
    Else
        POSmsg.Label = "STOPPOLLING"
    End If
    POSmsg.send QPOS
    Set POSmsg = Nothing
    Set QI = Nothing
End Sub

Public Function SendNielsen(pFilename As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean

    bOK = True
    If oPC.Configuration.NielsenActive Then
        If SetupNielsenFTP Then
            If SendNielsenFiles(pFilename) = True Then
                oTF.WriteToTextFile "Sales data: " & pFilename & "  sent at " & Format(Now(), "HH:nn")
            Else
                oTF.WriteToTextFile "Unable to send sales data: " & pFilename & "  " & Format(Now(), "HH:nn")
                bOK = False
            End If
           CloseFTP
        Else
            oTF.WriteToTextFile "Failed to set up sales data FTP connection " & Format(Now(), "HH:nn")
            bOK = False
        End If
    Else
            bOK = False
            oTF.WriteToTextFile "Sales export is set off; no data sent " & Format(Now(), "HH:nn")
    End If
    
    SendNielsen = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.SendNielsen(pFilename)", pFilename
End Function

Public Function SendLoyalty() As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
    bOK = False
    If SetupLoyaltyFTP Then
        If SendLoyaltyFiles() = True Then
            bOK = True
            oTF.WriteToTextFile "Loyalty data sent at " & Format(Now(), "HH:nn")
        End If
        CloseFTP
    Else
        bOK = False
        oTF.WriteToTextFile "Failed to set up Loyalty data FTP connection " & Format(Now(), "HH:nn")
    End If
    SendLoyalty = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.SendLoyalty"
End Function
Public Function SendStockSharing() As Boolean
    On Error GoTo errHandler
    
Dim bOK As Boolean
    bOK = False
    If SetupStockSharingFTP() = True Then
        If SendStockSharingFiles() = True Then
            bOK = True
            oTF.WriteToTextFile "Stock sharing data sent at " & Format(Now(), "HH:nn")
        End If
        CloseFTP
        bOK = True
    End If
    SendStockSharing = bOK

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.SendStockSharing"
End Function
Public Function SendNielsenFiles(strFilename As String) As Boolean
    On Error GoTo errHandler
Dim oEx As New z_Export
Dim oFSO As New FileSystemObject
Dim fold
Dim oOP As New a_Operation
Dim res As Boolean
Dim s As String
Dim strFiles As String
Dim Files
Dim File
Dim strFolder As String
  
    If oFSO.FileExists(oPC.ServerRootFolder & "\Data\NielsenSales\SALES.zip") Then
        res = FTP1.PutFile(oPC.ServerRootFolder & "\Data\NielsenSales\SALES.zip", strFilename, True)
        If res = False Then
            oTF.WriteToTextFile "Cannot put file: " & oPC.ServerRootFolder & "\Data\NielsenSales\SALES.zip"
            oFSO.DeleteFile oPC.ServerRootFolder & "\Data\NielsenSales\SALES.ZIP"
            SendNielsenFiles = False
            Exit Function
        Else
            SendNielsenFiles = True
            oTF.WriteToTextFile "Putting file: " & oPC.ServerRootFolder & "\Data\NielsenSales\SALES.zip to " & strFilename
        End If
    Else
        SendNielsenFiles = False
        oTF.WriteToTextFile "No SALES.ZIP exists: " & oPC.ServerRootFolder & "\Data\NielsenSales\SALES.zip"
        Exit Function
    End If

    'Only delete when they have been successfully sent: Later we will only delete when confirmation has been received
    strFolder = oPC.ServerRootFolder & "\Data\NielsenSales"
    Set fold = oFSO.GetFolder(strFolder)
    If fold.Files.Count > 0 Then
       oFSO.DeleteFile strFolder & "\*.*"
    End If
    oTF.WriteToTextFile "End of sales files transmission: " & CStr(res)
    Exit Function
errHandler:
    ErrPreserve
    oTF.WriteToTextFile "End of sales files transmission - Failed: " & CStr(res)
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.SendNielsenFiles(strFilename)", strFilename
End Function
Public Function SendLoyaltyFiles() As Boolean
    On Error GoTo errHandler
Dim oFSO As New FileSystemObject
Dim fold
Dim res As Boolean
Dim strFolder As String
Dim F
Dim fc
Dim bOK As Boolean

    Set fold = oFSO.GetFolder(oPC.ServerRootFolder & "\Data\Loyalty\UP")
    Set fc = fold.Files
    For Each F In fc
        If UCase(Right(F.Name, 4)) = ".ZIP" Then
            res = FTP1.PutFile(F.Path, F.Name, True) ', EXC_GENERAL, "Error putting file:SendLoyalty files"
            If res = False Then
                oTF.WriteToTextFile "Cannot put file: " & F.Path
                SendLoyaltyFiles = False
                Exit Function
            Else
                oTF.WriteToTextFile "Putting file: " & F.Path
            End If
        End If
    Next
    SendLoyaltyFiles = True

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.SendLoyaltyFiles"
End Function
Public Sub ClearObsoleteStockSharingFilesFROMSERVER()

End Sub
Public Function SendStockSharingFiles() As Boolean
    On Error GoTo errHandler
Dim oFSO As New FileSystemObject
Dim fold, fc, F
Dim res As Boolean
Dim strFolder As String
  
   
    Set fold = oFSO.GetFolder(oPC.ServerRootFolder & "\Data\StockSharing\UP")
    Set fc = fold.Files
    For Each F In fc
        If UCase(Right(F.Name, 4)) = ".ZIP" Then
            res = FTP1.PutFile(F.Path, F.Name, True) ', EXC_GENERAL, "Error putting file:SendStockSharingFiles"
            If res = False Then
                oTF.WriteToTextFile "Cannot put file: " & F.Path
                SendStockSharingFiles = False
                Exit Function
            Else
                oTF.WriteToTextFile "Putting file: " & F.Path
            End If
        End If
    Next
    SendStockSharingFiles = True
Exit Function

   ' Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.SendStockSharingFiles"
End Function

Public Function SetupNielsenFTP() As Boolean
    On Error GoTo errHandler
Dim res As Boolean
    SetupNielsenFTP = False
    res = FTP1.OpenFTP(oPC.Configuration.FTPAddress, oPC.Configuration.FTPUsername, oPC.Configuration.FTPPassword)  ', oPC.Configuration.FTPPassive
    If res Then
        If oPC.Configuration.FTPDefaultFolder > "" Then
            res = FTP1.SetCurrentFolder(oPC.Configuration.FTPDefaultFolder) ', EXC_GENERAL, "Error setting FTP folder"
            If res = False Then
                oTF.WriteToTextFile "Cannot set current folder: " & oPC.Configuration.FTPDefaultFolder
                Exit Function
            End If
        End If
    Else
        oTF.WriteToTextFile "Cannot open Nielsen FTP site " & oPC.Configuration.FTPAddress
        Exit Function
    End If
    SetupNielsenFTP = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.SetupNielsenFTP"
End Function
Public Function SetupLoyaltyFTP()
    On Error GoTo errHandler
Dim res As Boolean
Dim ftpFile As FTPFileClass

    SetupLoyaltyFTP = False
    res = FTP1.OpenFTP(oPC.CentralFTPAddress, oPC.CentralFTPUsername, oPC.CentralFTPPassword)
    If res Then
        If oPC.CentralFTPFolder > "" Then
            res = FTP1.SetCurrentFolder(oPC.CentralFTPFolder & "/LOYALTY/UP") ', EXC_GENERAL, "Error setting FTP folder"
            If res = False Then
                oTF.WriteToTextFile "Cannot set current folder " & oPC.CentralFTPFolder & "/LOYALTY/UP/"
                Exit Function
            Else
            'clear old files
                For Each ftpFile In FTP1.Files
                     If Left(ftpFile.FileName, Len(oPC.ClientCode)) = oPC.ClientCode Then
                        res = FTP1.DeleteFile(ftpFile.FileName)
                        If res = False Then
                            oTF.WriteToTextFile "Cannot delete old  file " & ftpFile.FileName & "   " & Format(Now(), "HH:nn")
                        End If
                        oTF.WriteToTextFile "File " & ftpFile.FileName & " deleted from server: " & Format(Now(), "HH:nn")
                     End If
                Next
            End If
        End If
    Else
        oTF.WriteToTextFile "Cannot open Loyalty FTP site"
        Exit Function
    End If
    SetupLoyaltyFTP = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.SetupLoyaltyFTP"
End Function

Public Function SetupStockSharingFTP() As Boolean
    On Error GoTo errHandler
Dim res As Boolean
Dim ftpFile As FTPFileClass
    
    SetupStockSharingFTP = False
    res = FTP1.OpenFTP(oPC.CentralFTPAddress, oPC.CentralFTPUsername, oPC.CentralFTPPassword)
    If res Then
        If oPC.CentralFTPFolder > "" Then
            res = FTP1.SetCurrentFolder(oPC.CentralFTPFolder & "/STOCKSHARING/UP") ', EXC_GENERAL, "Error setting FTP folder"
            If res = False Then
                oTF.WriteToTextFile "Cannot set current folder " & oPC.CentralFTPFolder & "/STOCKSHARING/UP"
                Exit Function
            Else
            'clear old files
                For Each ftpFile In FTP1.Files
                     If Left(ftpFile.FileName, Len(oPC.ClientCode)) = oPC.ClientCode Then
                        res = FTP1.DeleteFile(ftpFile.FileName)
                        If res = False Then
                            oTF.WriteToTextFile "Cannot delete old  file " & ftpFile.FileName & "   " & Format(Now(), "HH:nn")
                        End If
                        oTF.WriteToTextFile "File " & ftpFile.FileName & " deleted from server: " & Format(Now(), "HH:nn")
                     End If
                Next
                
            End If
        End If
    Else
        oTF.WriteToTextFile "Cannot open Stocksharing FTP site"
        Exit Function
    End If
    SetupStockSharingFTP = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.SetupStockSharingFTP"
End Function

Public Sub CloseFTP()
    On Error GoTo errHandler
    FTP1.CloseFTP
        oTF.WriteToTextFile "Closing FTP: " & Format(Now(), "HH:nn")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.CloseFTP"
End Sub
Public Sub Connect()
    On Error GoTo errHandler
    
    If oPC.InternetDialup = True Then
        oTF.WriteToTextFile "Connecting at " & Format(Now(), "HH:nn") & oPC.Configuration.DUN
        Set fINET = New wininet
        lngResult = fINET.StartDUN(0, oPC.Configuration.DUN, True)
    End If

    Exit Sub
errHandler:
    ErrorIn "z_Export.Connect"
End Sub
Public Sub Hangup()
    On Error GoTo errHandler
    If oPC.InternetDialup = True Then
        fINET.Hangup
        oTF.WriteToTextFile "Disconnecting at " & Format(Now(), "HH:nn")
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.HangUp"
End Sub

Public Sub FetchLCResponses()
    On Error GoTo errHandler
Dim oFSO As New FileSystemObject
Dim strLocalPath As String
Dim strLocalFolder As String
Dim res As Boolean
Dim fold, fc, F
Dim ftpFile As FTPFileClass

'Set local receiving folder
    strLocalFolder = oPC.SharedFolderRoot & "\DATA\LOYALTY\DOWN\"
    If Not oFSO.FolderExists(strLocalFolder) Then
        oFSO.CreateFolder strLocalFolder
    End If
    
'Set remote FTP folder
    res = FTP1.OpenFTP(oPC.CentralFTPAddress, oPC.CentralFTPUsername, oPC.CentralFTPPassword) ', EXC_GENERAL, "Error opening FTP:e:FetchLCConfirmations"
    If res = False Then
        oTF.WriteToTextFile "Cannot open FTP site: " & oPC.CentralFTPAddress
        Exit Sub
    End If
    res = FTP1.SetCurrentFolder(oPC.CentralFTPFolder & "/LOYALTY/DOWN")    ', EXC_GENERAL, "Error setting FTP folder"
    If res = False Then
        oTF.WriteToTextFile "Cannot Set current folder: " & oPC.CentralFTPAddress & "/LOYALTY/DOWN"
        Exit Sub
    End If

'Fetch all .LCR (receipts) and .LCE (Edited files) from FTP
    For Each ftpFile In FTP1.Files
         If Left(ftpFile.FileName, 3) = "LCR" Or (Left(ftpFile.FileName, 3) = "LCE" And ftpFile.ModifyDate > oPC.Configuration.DateLastLCEditedReceived) Then '  Returns and Edited data ' Len(oPC.ClientCode)) <> oPC.ClientCode Then   'Don't fetch own files
            res = FTP1.GetFile(ftpFile.FileName, strLocalFolder & "\" & ftpFile.FileName, True)   ', EXC_GENERAL, "Error getting file:FetchLCConfirmations"
            If res = False Then
                oTF.WriteToTextFile "Cannot fetch file " & ftpFile.FileName & "   " & Format(Now(), "HH:nn")
            Else
                oTF.WriteToTextFile "File " & ftpFile.FileName & " fetched: " & Format(Now(), "HH:nn")
            End If
         End If
    Next
    oTF.WriteToTextFile "Loyalty club returns fetched at " & Format(Now(), "HH:nn")
    
'Distribute to subfolders "\Data\Loyalty\Receipts\" and "\Data\Loyalty\Edited\"
    Set fold = oFSO.GetFolder(strLocalFolder)
    Set fc = fold.Files
    For Each F In fc
        Select Case Left(F.Name, 3)
        Case "LCR"
            If oFSO.FileExists(oPC.SharedFolderRoot & "\Data\Loyalty\Receipts\" & F.Name) Then
                oFSO.DeleteFile oPC.SharedFolderRoot & "\Data\Loyalty\Receipts\" & F.Name
            End If
            F.Move oPC.SharedFolderRoot & "\Data\Loyalty\Receipts\" & F.Name
        Case "LCE"
            If oFSO.FileExists(oPC.SharedFolderRoot & "\Data\Loyalty\Edited\" & F.Name) Then
                oFSO.DeleteFile oPC.SharedFolderRoot & "\Data\Loyalty\Edited\" & F.Name
            End If
            F.Move oPC.SharedFolderRoot & "\Data\Loyalty\Edited\" & F.Name
        End Select
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.FetchLCResponses"
End Sub
Public Sub UpdateFromEditedLC()
    On Error GoTo errHandler
Dim oFSO As New FileSystemObject
Dim fold, fc, F
Dim res As Boolean
Dim strFolder As String
Dim dteLastDate As Date
   
    Set fold = oFSO.GetFolder(oPC.ServerRootFolder & "\Data\Loyalty\Edited")
    Set fc = fold.Files
    For Each F In fc
        If UCase(Right(F.Name, 4)) = ".ZIP" Then
            Set zip = CreateObject("FathZIP.FathZIPCtrl.1")
            zip.OpenZip (F.Path)
            zip.BasePath = oPC.SharedFolderRoot & "\Data\Loyalty\"
            zip.ExtractFile ("*.*")
            zip.Close
            Set zip = Nothing
            F.Delete
        End If
    Next
    Set fold = oFSO.GetFolder(oPC.ServerRootFolder & "\Data\Loyalty\Edited")
    Set fc = fold.Files
    dteLastDate = CDate(0)
    For Each F In fc
        If ExtractDateAndTimeFromFilename(F.Name) > dteLastDate Then
            dteLastDate = ExtractDateAndTimeFromFilename(F.Name)
        End If
        UpdateCustomerFromFile_Papyrus F.Path
        oTF.WriteToTextFile "Updating from file " & F.Path

    Next
    For Each F In fc
        On Error Resume Next
        F.Delete True
        On Error GoTo errHandler
    Next
    If dteLastDate > CDate(0) Then
        oPC.COShort.Execute "Update tNielsen SET N_DateLastLCEditedReceived = '" & ReverseDateTime(dteLastDate) & "'"
    End If
    Exit Sub

errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.SendStockSharingFiles"

End Sub
Private Function ExtractDateFromFilename(pFN As String, Optional pStart As Integer) As Date
Dim strDateTime As String
Dim iStart As Integer

    If pStart > 0 Then
        iStart = pStart
    Else
        iStart = InStr(1, pFN, ".") - 8
    End If
    strDateTime = Mid(pFN, iStart, 4) & "/" & Mid(pFN, iStart + 4, 2) & "/" & Mid(pFN, iStart + 6, 2)
    If IsDate(strDateTime) Then
        ExtractDateFromFilename = CDate(strDateTime)
    Else
        ExtractDateFromFilename = CDate(0)
    End If
End Function
Private Function ExtractDateAndTimeFromFilename(pFN As String, Optional pStart As Integer) As Date
Dim strDateTime As String
Dim iStart As Integer
    If pStart > 0 Then
        iStart = pStart
    Else
        iStart = 4
    End If
    strDateTime = Mid(pFN, iStart, 4) & "/" & Mid(pFN, iStart + 4, 2) & "/" & Mid(pFN, iStart + 6, 2) & " " & Mid(pFN, iStart + 8, 2) & ":" & Mid(pFN, iStart + 10, 2)
    If IsDate(strDateTime) Then
        ExtractDateAndTimeFromFilename = CDate(strDateTime)
    Else
        ExtractDateAndTimeFromFilename = CDate(0)
    End If
End Function

Private Sub UpdateCustomerFromFile_Papyrus(pFilePath As String)
    On Error GoTo errHandler
Dim oTF As z_TextFile
Dim strLine As String
Dim ar() As String
Dim cmd As ADODB.Command
Dim oLC As New z_LCManager
Dim lngTPID As Long
Dim arAdd() As String
Dim i, j As Integer
Dim strStoreCode As String

    Set oTF = New z_TextFile
    oTF.OpenTextFileToRead pFilePath
    Do While Not oTF.IsEOF
        ReDim arAdd(6)
        ReDim ar(20)
        strLine = oTF.ReadLinefromTextFile
        If strLine = "" Then Exit Do
        ar = Split(strLine, vbTab)

        i = 0
        j = 0
    
        If UBound(ar) > 22 Then
            strStoreCode = Trim(ar(23))
        Else
            strStoreCode = "0"
        End If
        If UBound(ar) > 20 Then
             oLC.InsertLoyaltyCustomer lngTPID, Trim(ar(17)), Trim(ar(0)), Trim(ar(2)), Trim(ar(1)), "", _
             CreateAddressee(Trim(ar(1)), Trim(ar(0)), Trim(ar(2)), ""), Trim(ar(4)), Trim(ar(5)), "", "", Trim(ar(6)), Trim(ar(7)), _
            Trim(ar(22)), Trim(ar(8)), PhoneFormat(Trim(ar(10)), oPC.DefaultAreaCode), PhoneFormat(Trim(ar(11)), oPC.DefaultAreaCode), PhoneFormat(Trim(ar(3)), oPC.DefaultAreaCode), Trim(ar(13)), _
             Mid(Trim(ar(21)), 1, 1) = "1", Mid(Trim(ar(19)), 1, 1) = "1", Mid(Trim(ar(20)), 1, 1) = "1", strStoreCode
        End If
    
    
    
    Loop
    oTF.CloseTextFile
    Set oTF = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.UpdateCustomerFromFile_Papyrus(pFilePath)", pFilePath
End Sub
Public Sub FetchStockSharing()
    On Error GoTo errHandler
Dim oFSO As New FileSystemObject
Dim F, fol, fc
Dim strLocalPath As String
Dim res As Boolean
Dim ftpFile As FTPFileClass

'MsgBox "fetchStockSharing"
    res = FTP1.OpenFTP(oPC.CentralFTPAddress, oPC.CentralFTPUsername, oPC.CentralFTPPassword)
    If res = False Then
        oTF.WriteToTextFile "Cannot open FTP " & oPC.CentralFTPAddress & "," & oPC.CentralFTPUsername & "," & oPC.CentralFTPPassword
    End If

    strLocalPath = oPC.SharedFolderRoot & "\DATA\STOCKSHARING\DOWN"
    If Not oFSO.FolderExists(strLocalPath) Then
        oFSO.CreateFolder strLocalPath
    End If
    'Clear all older files in the folder, not needed now
    Set fol = oFSO.GetFolder(strLocalPath)
    Set fc = fol.Files
    For Each F In fc
        On Error Resume Next
        F.Delete True
        On Error GoTo errHandler
    Next
    res = FTP1.SetCurrentFolder(oPC.CentralFTPFolder & "/STOCKSHARING/UP")   ', EXC_GENERAL, "Error setting FTP folder"
    If res = False Then
        oTF.WriteToTextFile "Cannot set current folder " & oPC.CentralFTPFolder & "/STOCKSHARING/UP"
        Exit Sub
    End If
    For Each ftpFile In FTP1.Files
         If Left(ftpFile.FileName, Len(oPC.ClientCode)) <> oPC.ClientCode Then   'Don't fetch own files
            res = FTP1.GetFile(ftpFile.FileName, strLocalPath & "\" & ftpFile.FileName, True)   ', EXC_GENERAL, "Error getting file:FetchStockSharing"
            If res = False Then
                oTF.WriteToTextFile "Cannot get file " & ftpFile.FileName
                Exit Sub
            End If
            oTF.WriteToTextFile "Stock sharing file " & ftpFile.FileName & "  fetched at " & Format(Now(), "HH:nn")
         End If
    Next

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.FetchStockSharing"
End Sub

Public Sub UpdateStockSharing()
    On Error GoTo errHandler
Dim F, fc, fol
Dim oFSO As New FileSystemObject
Dim strSQL As String
Dim strStoreCode As String
Dim dteFileDate As Date
Dim lngRec As Long
Dim lngRecs As Long
Dim oBF As New z_BF
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim lngTimeOut As Long
Dim strErrPos As String

    On Error Resume Next
    Set fol = oFSO.GetFolder(oPC.SharedFolderRoot & "\DATA\STOCKSHARING\DOWN")
    Set fc = fol.Files
    For Each F In fc
        If UCase(Right(F.Name, 4)) = ".ZIP" Then
            Set zip = CreateObject("FathZIP.FathZIPCtrl.1")
            zip.OpenZip (F.Path)
            zip.BasePath = oPC.SharedFolderRoot & "\DATA\STOCKSHARING\DOWN"
            zip.PreservePaths = False
            zip.ExtractFile ("*.*")
            zip.Close
            Set zip = Nothing
            F.Delete
        End If
    Next
    
    oTF.WriteToTextFile "Position 1"
    
strErrPos = "Pos 1"
    On Error GoTo errHandler
    Set fol = oFSO.GetFolder(oPC.SharedFolderRoot & "\DATA\STOCKSHARING\DOWN")
    Set fc = fol.Files
    For Each F In fc
        strStoreCode = Left(F.Name, InStr(1, F.Name, "_") - 1)
        LogSaveToFile "Filename = " & F.Name
        dteFileDate = ExtractDateFromFilename(F.Name)
        LogSaveToFile "dteFileDate = " & dteFileDate
        strSQL = "DELETE FROM dbo.tSS"
        lngTimeOut = oPC.COShort.CommandTimeout
    oTF.WriteToTextFile "Position 2"
        
strErrPos = "Pos 2"
        oPC.COShort.CommandTimeout = 0
        oPC.COShort.Execute strSQL, lngRecs
        strSQL = "BULK INSERT PBKS.dbo.tSS From '" & F.Path & "'" _
        & " WITH (FIELDTERMINATOR = '\t', ROWTERMINATOR = '\n')"
        oPC.COShort.Execute strSQL
    oTF.WriteToTextFile "Position 3"
strErrPos = "Pos 3"
        'Find local PIDs for all matches on tSS against the local DB using EAN and/or ISBN
        oPC.COShort.Execute "EXEC dbo.sp_UpdateSS_1"
        'If Bookfind is supported, insert records from Bookfind for non-matches
        oBF.InsertFromBookfind
    oTF.WriteToTextFile "Position 4"
        
strErrPos = "Pos 4"
        'Update the tSS records that have not been matched against the newly inserted records
        oPC.COShort.Execute "EXEC dbo.sp_UpdateSS_2"
        
        'Update all tStorep records to 0 or null before updating ( in case records with values for foreign branch have no new record to update them - we can't keep the old values
    oTF.WriteToTextFile "Position 5"
        
strErrPos = "Pos 5"
        
        'update qtys of matching records by ISBN
        Set cmd = New ADODB.Command
        Set prm = cmd.CreateParameter("@STORECODE", adVarChar, , 3, strStoreCode)
        cmd.Parameters.Append prm
        cmd.CommandText = "sp_UpdateSS_3"
        cmd.CommandType = adCmdStoredProc
        cmd.ActiveConnection = oPC.COShort
        cmd.CommandTimeout = 0
        cmd.Execute
        Set cmd = Nothing
strErrPos = "Pos 6"
        oPC.COShort.Execute "UPDATE tSTORE SET STORE_LastSharedDate = '" & ReverseDateTime(dteFileDate) & "'"
        oTF.WriteToTextFile "UPDATING Stock sharing tables " & F.Name & Format(Now(), "HH:nn")
        oPC.COShort.CommandTimeout = lngTimeOut
    Next
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.UpdateStockSharing", , , , "strErrPos", Array(strErrPos)
End Sub
Public Sub DeleteReceipted()
    On Error GoTo errHandler
Dim oFSO As New FileSystemObject
Dim fold, fc, F
Dim strLocalPath As String
Dim strLocalUpPath As String
Dim oTF As z_TextFile
Dim strLine As String
Dim ar() As String
Dim i As Integer
Dim dteLatestDate As Date

    strLocalPath = oPC.SharedFolderRoot & "\DATA\LOYALTY\Receipts"
    strLocalUpPath = oPC.SharedFolderRoot & "\DATA\LOYALTY\UP"
    Set fold = oFSO.GetFolder(strLocalPath)
    Set fc = fold.Files
    dteLatestDate = CDate(0)
    For Each F In fc
        Set oTF = New z_TextFile
        oTF.OpenTextFileToRead F.Path
        Do While Not oTF.IsEOF
            strLine = oTF.ReadLinefromTextFile
            ar() = Split(strLine, ",")
            For i = 0 To UBound(ar)
                If oFSO.FileExists(strLocalUpPath & "\" & ar(i)) Then
                    On Error Resume Next  ' dont stop the whole thing because of this
                    Set F = oFSO.GetFile(strLocalUpPath & "\" & ar(i))
                    If F.Date > dteLatestDate Then
                        dteLatestDate = F.DateCreated
                    End If
                    oFSO.DeleteFile strLocalUpPath & "\" & ar(i)
                    oTF.WriteToTextFile "Deleting file " & strLocalUpPath & "\" & ar(i)
                    On Error GoTo errHandler
                End If
            Next i
        Loop
        oTF.CloseTextFile
        Set oTF = Nothing
    Next
    For Each F In fc
        If oFSO.FileExists(F.Path) Then
            oFSO.DeleteFile F.Path  'Deletes receipt file
        End If
    Next
    If CDate(dteLatestDate) > 0 Then
        oPC.COShort.Execute "Update tNielsen Set N_LastDateLCReceivedAtCentral = '" & ReverseDate(dteLatestDate) & "' FROM tNielsen"
    End If
 '   oPC.Configuration.BeginEdit
 '   oPC.Configuration.LastDateLCSent = dteLatestDate
 '   oPC.Configuration.ApplyEdit str
    
    Exit Sub
errHandler:
    ErrPreserve
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.DeleteReceipted"
End Sub
Private Function InsertLoyaltyCustomer(pTPID As Long, pACNO As String, pLastName As String, pTitle As String, pFirstname As String, _
pInitials As String, pADDRESSEE As String, pADDL1 As String, pADDL2 As String, pADDL3 As String, pADDL4 As String, pADDL5 As String, pADDL6 As String, _
pCountryName As String, pPCode As String, pPhone1 As String, pPhone2 As String, pMobile As String, pEmail As String, _
pBooksale As Boolean, pBookpromotion As Boolean, pBookLaunch As Boolean, pStoreCode As String, pIDNUM As String) As Long
    On Error GoTo errHandler

Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_InsertLoyaltyCustomer"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@ACNO", adVarChar, adParamInput, 10)
    cmd.Parameters.Append par
    par.Value = pACNO
    Set par = cmd.CreateParameter("@Lastname", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = pLastName
    Set par = cmd.CreateParameter("@Title", adVarChar, adParamInput, 20)
    cmd.Parameters.Append par
    par.Value = pTitle
    Set par = cmd.CreateParameter("@Firstname", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = pFirstname
    Set par = cmd.CreateParameter("@Initials", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = pInitials
    Set par = cmd.CreateParameter("@ADDRESSEE", adVarChar, adParamInput, 150)
    cmd.Parameters.Append par
    par.Value = pADDRESSEE
    Set par = cmd.CreateParameter("@ADDL1", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = pADDL1
    Set par = cmd.CreateParameter("@ADDL2", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = pADDL2
    Set par = cmd.CreateParameter("@ADDL3", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = pADDL3
    Set par = cmd.CreateParameter("@ADDL4", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = pADDL4
    Set par = cmd.CreateParameter("@ADDL5", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = pADDL5
    Set par = cmd.CreateParameter("@ADDL6", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = pADDL6
    Set par = cmd.CreateParameter("@CountryName", adVarChar, adParamInput, 30)
    cmd.Parameters.Append par
    par.Value = pCountryName
    Set par = cmd.CreateParameter("@PCode", adVarChar, adParamInput, 10)
    cmd.Parameters.Append par
    par.Value = pPCode
    Set par = cmd.CreateParameter("@Phone1", adVarChar, adParamInput, 20)
    cmd.Parameters.Append par
    par.Value = pPhone1
    Set par = cmd.CreateParameter("@Phone2", adVarChar, adParamInput, 20)
    cmd.Parameters.Append par
    par.Value = pPhone2
    Set par = cmd.CreateParameter("@Mobile", adVarChar, adParamInput, 20)
    cmd.Parameters.Append par
    par.Value = pMobile
    Set par = cmd.CreateParameter("@Email", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = pEmail
    Set par = cmd.CreateParameter("@Booksale", adBoolean, adParamInput)
    cmd.Parameters.Append par
    par.Value = pBooksale
    Set par = cmd.CreateParameter("@Bookpromotion", adBoolean, adParamInput)
    cmd.Parameters.Append par
    par.Value = pBookpromotion
    Set par = cmd.CreateParameter("@Booklaunch", adBoolean, adParamInput)
    cmd.Parameters.Append par
    par.Value = pBookLaunch
    Set par = cmd.CreateParameter("@Source", adVarChar, adParamInput, 50)
    cmd.Parameters.Append par
    par.Value = pStoreCode
    'source
    'Pref1,2,3
    'DateCreated
    'DateModified
    'ID
    'Loyalty type (presently '1')
    
    cmd.ActiveConnection = oPC.COShort
    cmd.Execute
    pTPID = cmd.Parameters("@TPID").Value
    Set cmd = Nothing
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.InsertLoyaltyCustomer(pTPID,pACNO,pLastName,pTitle,pFirstname,pInitials," & _
        "pADDRESSEE,pADDL1,pADDL2,pADDL3,pADDL4,pADDL5,pADDL6,pCountryName,pPCode,pPhone1,pPhone2,pMobile," & _
        "pEmail,pBooksale,pBookpromotion,pBookLaunch,pStoreCode,pIDNUM)", Array(pTPID, pACNO, pLastName, _
         pTitle, pFirstname, pInitials, pADDRESSEE, pADDL1, pADDL2, pADDL3, pADDL4, pADDL5, pADDL6, _
         pCountryName, pPCode, pPhone1, pPhone2, pMobile, pEmail, pBooksale, pBookpromotion, pBookLaunch, _
         pStoreCode, pIDNUM)
End Function



