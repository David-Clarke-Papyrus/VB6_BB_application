VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Extracting sales from Wordstock"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResults 
      BackColor       =   &H00E3F9FD&
      Height          =   2175
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   660
      Width           =   5115
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTimeOut As Long
Dim strSQL As String

Dim res As Long

Public Sub DoWork()
    On Error GoTo errHandler
Dim res As Boolean
'Fetch file from Wordstock position using FTP
    res = FetchFile
        MsgBox "Pos 1" & strARG
    If res = False Then
        Exit Sub
    End If
        MsgBox "Pos 2" & strARG
    
    ExportLoyaltySales
        MsgBox "Pos 3" & strARG
    
'Transmit prepared file to Central
    If strARG <> "SB" Then
        lg "Creating export file"
        MsgBox "Creating export file, strARG= :" & strARG
        SendLoyalty
    End If
    res = SendByQueue
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.DoWork"
End Sub

Private Function FetchFile() As Boolean
    On Error GoTo errHandler
Dim strFN As String
Dim strCommand As String
Dim X As Long
Dim fso As New FileSystemObject
Dim strPos As String
Dim f As File
Dim fol As Folder


    lg "FTP source : " & gFTPSourceAddress & vbCrLf & "FTP user name : " & gFTPSourceUsername & vbCrLf & "FTP password : " & gFTPSourcePassword

    FetchFile = True
    
    gRes = FTP1.OpenFTP(gFTPSourceAddress, gFTPSourceUsername, gFTPSourcePassword, True)
    If gRes = False Then
        oTF.WriteToLogandsave "Cannot open FTP site: " & gFTPSourceAddress, strLogPath
        lg "Failed to open FTP site"
        FetchFile = False
        Exit Function
    Else
        oTF.WriteToLogandsave "Opening FTP site " & gFTPSourceAddress & " . . ." & Format(Now, "HH:NN"), strLogPath
        lg "Opened FTP site"
    End If
    
    lg "Fetching files. . ." & gFTPSourceAddress
    gRes = FTP1.SetCurrentFolder(gFTPSourceFolder)
'    Set fol = fso.GetFolder(DownloadFolder)
'    For Each F In fol
'        oTF.WriteToLogandsave "Checking: " & Backupfolder & "\" & F.Name
'        If fso.FileExists(Backupfolder & "\" & F.Name) Then
'            fso.DeleteFile Backupfolder & "\" & F.Name
'        End If
'        oTF.WriteToLogandsave "Moving: " & DownloadFolder & "\" & F.Name & " to " & Backupfolder
'        fso.MoveFile DownloadFolder & "\" & F.Name, Backupfolder
'    Next
    oTF.WriteToLogandsave "Filecount: " & CStr(FTP1.files.Count), strLogPath
    For Each ftpFile In FTP1.files
        lg ". . . " & ftpFile.FileName
        gRes = FTP1.GetFile(ftpFile.FileName, DownloadFolder & "\F.TXT", True)
        lg "FTP source file name: " & ftpFile.FileName
        If gRes = False Then
            oTF.WriteToLogandsave "Cannot Get file: " & ftpFile.FileName, strLogPath
            Exit Function
        Else
            strCommand = "C:\PBKS\Executables\SED.BAT"
            res = F_7_AB_1_ShellAndWaitSimple(strCommand)
            ImportFile
            fso.CopyFile DownloadFolder & "\F.TXT", Backupfolder & ftpFile.FileName
            gRes = FTP1.DeleteFile(ftpFile.FileName)
            oTF.WriteToLogandsave ftpFile.FileName & " successful import", strLogPath
        End If
    Next
    Exit Function
    FTP1.CloseFTP
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.FetchFile"
End Function
Private Sub ImportFile()
    On Error GoTo errHandler
Dim fol
Dim fc
Dim f As File
Dim oTmp As z_TextFile
Dim rs As New ADODB.Recordset
Dim rsOut As New ADODB.Recordset
Dim str As String
Dim str2 As String
Dim ar() As String
Dim strPos As String
Dim i As Integer

        Set oTmp = New z_TextFile
        
        lngTimeOut = cn.CommandTimeout
        cn.CommandTimeout = 0
        
        cn.Execute "DELETE FROM tRAW"
        cn.Execute "DELETE FROM tRAWDATA"
        cn.Execute "DELETE FROM tEXCHANGE"
        strSQL = "BULK INSERT PBKS_WSTOCK.dbo.tRAW From 'c:\PBKS\DOWNLOADFOLDER\fOUT.TXT'" & " WITH (FIELDTERMINATOR = '**', ROWTERMINATOR = '\n')"
        cn.Execute strSQL
        
        cn.CommandTimeout = lngTimeOut
        rsOut.Open "SELECT * FROM tRawData", cn, adOpenDynamic, adLockOptimistic
        rs.Open "SELECT * FROM tRAW", cn, adOpenDynamic, adLockOptimistic
        Do While Not rs.EOF
            str = rs.Fields(0)
            ar() = Split(str, "|")
            rsOut.AddNew
                For i = 0 To UBound(ar)
                    rsOut.Fields(i) = ar(i)
                Next i
            rsOut.Update
            rs.MoveNext
        Loop
        
    On Error Resume Next
        cn.Execute "DROP TABLE tCSL"
        cn.Execute "DROP TABLE tPayments"
 strPos = "Pos 4"
    On Error GoTo errHandler
        cn.Execute "INSERT tEXCHANGE (EXCH_NUMBER) SELECT EXCH_NUMBER FROM vExtractExchangeNumbers"
 strPos = "Pos 5"
        cn.Execute "UPDATE tEXCHANGE SET EXCH_SALEVALUE = LineTotal FROM tEXCHANGE JOIN vEXCHVALUE ON EXCH_NUMBER = EXCHNUM"
 strPos = "Pos 6"
        cn.Execute "UPDATE tEXCHANGE SET EXCH_CHANGEGIVEN = CHANGEGiven FROM tEXCHANGE JOIN vEXCHCHANGEGIVEN ON EXCH_NUMBER = EXCHNUM"
 strPos = "Pos 7"
        cn.Execute "UPDATE tEXCHANGE SET EXCH_DISCOUNTVALUE = DiscountValue FROM tEXCHANGE JOIN vEXCHDISCOUNT ON EXCH_NUMBER = EXCHNUM"
 strPos = "Pos 8"
        cn.Execute "UPDATE tEXCHANGE SET EXCH_VATVALUE = VATvalue FROM tEXCHANGE JOIN vEXCHVATVALUE ON EXCH_NUMBER = EXCHNUM"
 strPos = "Pos 9"
        cn.Execute "UPDATE tEXCHANGE SET EXCH_TYPE = 'S',EXCH_ACNO = ACNO,EXCH_SALEDATE = SALEDATE,EXCH_STAFFMEMBER = SMID,EXCH_TILLPOINT = TILLPOINT " _
            & "FROM tEXCHANGE JOIN vGROUPVALUES on EXCH_NUMBER = EXCHNUM"
 strPos = "Pos 10"
        cn.Execute "SELECT *  into tCSL FROM vExtractSales"
 strPos = "Pos 11"
        cn.Execute "SELECT *  into tPayments FROM vExtractReceipts"
        
        Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ImportFile", , , , "strPos", Array(strPos)
End Sub


Private Sub lg(pText As String)
    On Error GoTo errHandler
    txtResults = txtResults & vbCrLf & pText
    txtResults.Refresh
    txtResults.SelStart = Len(txtResults) - 1
    txtResults.SelLength = 0
    oTF.WriteToLogandsave pText, strLogPath
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.lg(pText)", pText
End Sub

Public Sub ExportLoyaltySales()
    On Error GoTo errHandler
Dim strCommand As String
Dim strSQL As String
Dim strFilePathText As String
Dim strFilePathZip As String
Dim zip
Dim strFolder As String
Dim dte As Date
Dim strSalFilePathText As String

    dte = Now()
    strSalFilePathText = strFolder & gStoreCode & "_SA" & ReverseDateTimeStripped(dte) & ".ZIP"

    strSQL = "SELECT * FROM PBKS_WSTOCK.dbo.vLoyaltySales_EXCHANGES"
    strCommand = "bcp """ & strSQL & """ queryout """ & "\PBKS\Exch.txt" & """ -eBCPError.sal -c -q  -Usa -P" & gPassword & " -S " & strServerName
            res = F_7_AB_1_ShellAndWaitSimple(strCommand)

    strSQL = "SELECT * FROM PBKS_WSTOCK.dbo.vLoyaltySales_CSLS"
    strCommand = "bcp """ & strSQL & """ queryout """ & "\PBKS\CSL.txt" & """ -eBCPError.sal -c -q  -Usa -P" & gPassword & " -S " & strServerName
            res = F_7_AB_1_ShellAndWaitSimple(strCommand)

    strSQL = "SELECT * FROM PBKS_WSTOCK.dbo.vLoyaltySales_PAYMENTS"
    strCommand = "bcp """ & strSQL & """ queryout """ & "\PBKS\Pay.txt" & """ -eBCPError.sal -c -q  -Usa -P" & gPassword & " -S " & strServerName
            res = F_7_AB_1_ShellAndWaitSimple(strCommand)

Dim fso As New FileSystemObject
    If fso.GetFolder(strSharedServerFolder & "\Data\Loyalty\UP\").files.Count > 0 Then
        fso.DeleteFile strSharedServerFolder & "\Data\Loyalty\UP\*.*", True
    End If
    
'Zip file and delete
    Set zip = CreateObject("FathZIP.FathZIPCtrl.1")
    strFolder = strSharedServerFolder & "\Data\Loyalty\UP\"
    strFilePathZip = strFolder & strSalFilePathText
    zip.CreateZip strFilePathZip, ""
    zip.preservepaths = False
    
    strFilePathText = "\PBKS\Exch.txt"
    If oFSO.FileExists(strFilePathText) Then
        zip.AddFile strFilePathText, ""
        oFSO.DeleteFile strFilePathText
    End If
    
    strFilePathText = "\PBKS\CSL.txt"
    If oFSO.FileExists(strFilePathText) Then
        zip.AddFile strFilePathText, ""
        oFSO.DeleteFile strFilePathText
    End If
    
    strFilePathText = "\PBKS\Pay.txt"
    If oFSO.FileExists(strFilePathText) Then
        zip.AddFile strFilePathText, ""
        oFSO.DeleteFile strFilePathText
    End If
    
    zip.Close
    Set zip = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ExportLoyaltySales"
End Sub

Function ReverseDateTimeStripped(pDate As Date) As String
    On Error GoTo errHandler
Dim str As String
  str = Replace(ReverseDateTime(pDate), "-", "")
  str = Replace(str, ":", "")
  ReverseDateTimeStripped = Replace(str, " ", "")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ReverseDateTimeStripped(pDate)", pDate
End Function
Function ReverseDateTime(pDate As Date) As String
    On Error GoTo errHandler
  ReverseDateTime = Format(pDate, "yyyy-mm-dd HH:nn")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ReverseDateTime(pDate)", pDate
End Function

Private Sub ExtractTotalsFromRaw()
    On Error GoTo errHandler
Dim cmd As ADODB.Command

    Set cmd = New ADODB.Command
    cmd.CommandText = "CalculateExchangeValues"
    cmd.CommandType = adCmdStoredProc
    
    
    cmd.ActiveConnection = cn
    cmd.Execute
    Set cmd = Nothing
    

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ExtractTotalsFromRaw"
End Sub
Public Function SendLoyalty() As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim res
    bOK = False
    If gDialup = True Then
        Set fINET = New wininet
        res = fINET.StartDUN(0, gConnectionName, True)
    End If

    
    If SetupLoyaltyFTP Then
        If SendLoyaltyFiles() = True Then
            bOK = True
            oTF.WriteToLogandsave "Loyalty data sent at " & Format(Now(), "HH:nn"), strLogPath
        End If
        CloseFTP
    Else
        bOK = False
        oTF.WriteToLogandsave "Failed to set up Loyalty data FTP connection " & Format(Now(), "HH:nn"), strLogPath
    End If
    SendLoyalty = bOK
    If gDialup = True Then
        res = fINET.Hangup
    End If


    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.SendLoyalty"
End Function
Public Function SendLoyaltyFiles() As Boolean
    On Error GoTo errHandler
Dim oFSO As New FileSystemObject
Dim fold
Dim res As Boolean
Dim strFolder As String
Dim f
Dim fc
Dim bOK As Boolean

    Set fold = oFSO.GetFolder(strSharedServerFolder & "\Data\Loyalty\UP")
    Set fc = fold.files
    For Each f In fc
        If UCase(Right(f.Name, 4)) = ".ZIP" Then
            res = FTP1.PutFile(f.Path, f.Name, True) ', EXC_GENERAL, "Error putting file:SendLoyalty files"
            If res = False Then
                oTF.WriteToLogandsave "Cannot put file: " & f.Path, strLogPath
                SendLoyaltyFiles = False
                Exit Function
            Else
                oTF.WriteToLogandsave "Putting file: " & f.Path, strLogPath
            End If
        End If
    Next
    SendLoyaltyFiles = True

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.SendLoyaltyFiles"
End Function

Public Function SetupLoyaltyFTP()
    On Error GoTo errHandler
Dim res As Boolean
Dim ftpFile As FTPFileClass


    SetupLoyaltyFTP = False
    lg "FTP Target : " & gFTPTargetAddress & vbCrLf & "FTP user name : " & gFTPTargetUsername & vbCrLf & "FTP password : " & gFTPTargetPassword
    res = FTP1.OpenFTP(gFTPTargetAddress, gFTPTargetUsername, gFTPTargetPassword)
    If res Then
        If gFTPTargetFolder > "" Then
            res = FTP1.SetCurrentFolder(gFTPTargetFolder & "/LOYALTY/UP") ', EXC_GENERAL, "Error setting FTP folder"
            If res = False Then
                oTF.WriteToLogandsave "Cannot set current folder " & gFTPTargetFolder & "/LOYALTY/UP/", strLogPath
                Exit Function
            Else
            'clear old files
                For Each ftpFile In FTP1.files
                     If Left(ftpFile.FileName, Len(gStoreCode)) = gStoreCode Then
                        res = FTP1.DeleteFile(ftpFile.FileName)
                        If res = False Then
                            oTF.WriteToLogandsave "Cannot delete old  file " & ftpFile.FileName & "   " & Format(Now(), "HH:nn"), strLogPath
                        End If
                        oTF.WriteToLogandsave "File " & ftpFile.FileName & " deleted from server: " & Format(Now(), "HH:nn"), strLogPath
                     End If
                Next
            End If
        End If
    Else
        oTF.WriteToLogandsave "Cannot open Loyalty FTP site: FTP Target : " & gFTPTargetAddress & vbCrLf & "FTP user name : " & gFTPTargetUsername & vbCrLf & "FTP password : " & gFTPTargetPassword, strLogPath

        Exit Function
    End If
    SetupLoyaltyFTP = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.SetupLoyaltyFTP"
End Function

Public Sub CloseFTP()
    On Error GoTo errHandler
    FTP1.CloseFTP
        oTF.WriteToLogandsave "Closing FTP: " & Format(Now(), "HH:nn"), strLogPath
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.CloseFTP"
End Sub

Private Function SendByQueue() As Integer
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim par As ADODB.Parameter

    Set cmd.ActiveConnection = cn
    cmd.CommandText = "_SendCQ"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Res", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ErrMsg", adVarChar, adParamOutput, 200)
    cmd.Parameters.Append par
    
    cmd.Execute
    oTF.WriteToLogandsave "Sent to queue" & Format(Now(), "HH:nn"), strLogPath
    
    SendByQueue = (cmd.Parameters(0))
    Set cmd = Nothing
    

    Exit Function
errHandler:
    ErrPreserve
    oTF.WriteToLogandsave "Error in sending by queue:" & Error & Format(Now(), "HH:nn"), strLogPath
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.SendByQueue"
End Function


