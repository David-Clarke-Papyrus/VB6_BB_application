VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Loyalty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oTF As z_TextFile
Dim strSQL As String
Dim dteSince As Date
Private rs As ADODB.Recordset
Private oPC As a_connection
Private oFSO As New FileSystemObject
Dim dteLastSuccessful As Date
Dim dte As Date

Public Sub Component(pPC As a_connection)
    On Error GoTo errHandler
    Set oPC = pPC
    dteLastSuccessful = oPC.Configuration.LastDateLCSent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Loyalty.Component(pPC)", pPC
End Sub
Public Sub CreateLoyaltyExtractionFile()
    On Error GoTo errHandler
Dim strFilePathText As String
Dim strSalFilePathText As String
Dim strFilePathZip As String
Dim zip
Dim f As File
Dim strFolder As String
Dim strPos As String

    dte = Now
strPos = "Pos 1"
    strFolder = oPC.ServerRootFolder & "\Data\Loyalty\UP\"
    If Not oFSO.FolderExists(strFolder) Then
        oFSO.CreateFolder strFolder
    End If
strPos = "Pos 2"
'Append New customers
    strFilePathText = strFolder & oPC.ClientCode & "_LC" & ReverseDateTimeStripped(dte) & ".CUS"
    If oFSO.FileExists(strFilePathText) Then
        oFSO.DeleteFile strFilePathText, True
    End If
strPos = "Pos 3"
    AppendNewCustomers strFilePathText
strPos = "Pos 4"
'Append New sales
    strSalFilePathText = strFolder & oPC.ClientCode & "_SA" & ReverseDateTimeStripped(dte) & ".ZIP"
    If oFSO.FileExists(strSalFilePathText) Then
        oFSO.DeleteFile strSalFilePathText, True
    End If
strPos = "Pos 5"
    ExportLoyaltySales strSalFilePathText
strPos = "Pos 6"
    oPC.COShort.Execute "UPDATE tNIELSEN SET N_LastDateLCExtraction = GetDate()"  'This controls from where the next extraction occurs
strPos = "Pos 7"
'Zip file and delete
    Set zip = CreateObject("FathZIP.FathZIPCtrl.1")
    strFilePathZip = strFolder & oPC.ClientCode & "_LC" & ReverseDateTimeStripped(dte) & ".zip"
    zip.CreateZip strFilePathZip, ""
    zip.preservepaths = False
    If oFSO.FileExists(strFilePathText) Then
        Set f = oFSO.GetFile(strFilePathText)
        If f.Size > 0 Then
            zip.AddFile strFilePathText, ""
        End If
        Set f = Nothing
    End If
strPos = "Pos 8"
    If oFSO.FileExists(strSalFilePathText) Then
        Set f = oFSO.GetFile(strSalFilePathText)
        If f.Size > 0 Then
            zip.AddFile strSalFilePathText, ""
        End If
        Set f = Nothing
    End If
    zip.Close
    
strPos = "Pos 9"
    Set zip = Nothing
    If oFSO.FileExists(strFilePathText) Then oFSO.DeleteFile strFilePathText
strPos = "Pos 9a"
    If oFSO.FileExists(strSalFilePathText) Then oFSO.DeleteFile strSalFilePathText
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Loyalty.CreateLoyaltyExtractionFile", , , , "strPos", Array(strPos)
End Sub

Private Sub AppendNewCustomers(pTextFilename As String)
Dim strCommand As String
Dim oFSO As New FileSystemObject
    strSQL = "SELECT * FROM PBKS.dbo.vLCustomers"
    
    strCommand = "bcp """ & strSQL & """ queryout """ & pTextFilename & """ -c -q -Usa -Pcar -S" & oPC.ServerName
    'MsgBox "AppendNewCustomers   " & strCommand
    LogSaveToFile "z_Loyalty.AppendNewCustomers" & strCommand
    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Loyalty.AppendNewCustomers"
End Sub

Public Sub ExportLoyaltySales(pFilename)
Dim strCommand As String
Dim strSQL As String
Dim strFilePathText As String
Dim strFilePathZip As String
Dim zip
Dim strFolder As String

    strSQL = "SELECT * FROM PBKS.dbo.vLoyaltySales_EXCHANGES"
    strCommand = "bcp """ & strSQL & """ queryout """ & "\PBKS\Exch.txt" & """ -eBCPError.sal -c -q  -Usa -Pcar -S " & oPC.ServerName
    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide

    strSQL = "SELECT * FROM PBKS.dbo.vLoyaltySales_CSLS"
    strCommand = "bcp """ & strSQL & """ queryout """ & "\PBKS\CSL.txt" & """ -eBCPError.sal -c -q  -Usa -Pcar -S " & oPC.ServerName
    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide

    strSQL = "SELECT * FROM PBKS.dbo.vLoyaltySales_PAYMENTS"
    strCommand = "bcp """ & strSQL & """ queryout """ & "\PBKS\Pay.txt" & """ -eBCPError.sal -c -q  -Usa -Pcar -S " & oPC.ServerName
    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide

'Zip file and delete
    Set zip = CreateObject("FathZIP.FathZIPCtrl.1")
    strFolder = oPC.ServerRootFolder & "\Data\Loyalty\UP\"
    strFilePathZip = pFilename
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
End Sub




Public Sub CreateStockSharingExtractionFile()
    On Error GoTo errHandler
Dim strFilePathText As String
Dim strFilePathZip As String
Dim zip
Dim strFolder As String
Dim dte As Date
Dim fol
Dim fc
Dim f

    dte = Date
    strFolder = oPC.ServerRootFolder & "\Data\StockSharing\UP\"
    
    strFilePathText = strFolder & oPC.ClientCode & "_SS" & ReverseDateStripped(dte) & ".txt"     '& ReverseDateStripped(dte)
    If Not oFSO.FolderExists(strFolder) Then
        oFSO.CreateFolder strFolder
    Else
        'clear all existing files
        Set fol = oFSO.GetFolder(strFolder)
        Set fc = fol.Files
        For Each f In fc
            On Error Resume Next
            f.Delete
            On Error GoTo errHandler
        Next
    End If
    strFilePathZip = strFolder & oPC.ClientCode & "_SS" & ReverseDateStripped(dte) & ".zip"
'Create empty file
    If oFSO.FileExists(strFilePathText) Then
        oFSO.DeleteFile strFilePathText, True
    End If
    
    ExportStockSharing strFilePathText
    
'Zip file and delete
    Set zip = CreateObject("FathZIP.FathZIPCtrl.1")
    zip.CreateZip strFilePathZip, ""
    zip.preservepaths = False
    If oFSO.FileExists(strFilePathText) Then
        zip.AddFile strFilePathText, ""
    End If
    zip.Close
    If oFSO.FileExists(strFilePathText) Then
        oFSO.DeleteFile strFilePathText
    End If
    Set zip = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Loyalty.CreateStockSharingExtractionFile"
End Sub



Public Sub UpdateStock()
    On Error GoTo errHandler
Dim oFSO As New FileSystemObject
Dim fold, fc, f
Dim zip
Dim strPath As String
Dim strCommand As String
Dim strSource As String

'Unzip each file
    strPath = oPC.SharedFolderRoot & "\Data\StockSharing\Down"
    Set fold = oFSO.GetFolder(strPath)
    Set fc = fold.Files
    For Each f In fc
        'Unzip into DOWN folder
        If UCase(Right(f.Name, 4)) = ".ZIP" Then
            Set zip = CreateObject("FathZIP.FathZIPCtrl.1")
            zip.basepath = strPath
            zip.OpenZip (f.Path)
            zip.extractFile "*.*"
            zip.Close
            If oFSO.FileExists(f.Path) Then
                oFSO.DeleteFile f.Path
            End If
            Set zip = Nothing
        End If
    Next

'use bcp to load into tSS and update tStoreP from tSS
    Set fold = oFSO.GetFolder(strPath)
    Set fc = fold.Files
    For Each f In fc
        strSource = Left(f.Name, 3)
        oPC.COShort.Execute "Delete FROM tSS"
        strCommand = "bcp PBKS.dbo.tSS in """ & f.Path & """ -c -q -Usa -Pcar -S" & oPC.ServerName
        F_7_AB_1_ShellAndWaitSimple strCommand, vbHide
        oPC.COShort.Execute "Update tSS Set SS_Source = '" & strSource & "'"
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Loyalty.UpdateStock"
End Sub

Private Function ExportStockSharing(pFilename)
    On Error GoTo errHandler
Dim strCommand As String
Dim oFSO As New FileSystemObject

    strCommand = "bcp PBKS.dbo.vStockSharing_1 out """ & pFilename & """ -c -q -Usa -Pcar -S" & oPC.ServerName
    'MsgBox "ExportStockSharing" & strCommand
    LogSaveToFile "z_Loyalty.ExportStockSharing(pFilename)" & strCommand
    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide
    

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Loyalty.ExportStockSharing(pFilename)", pFilename, , , "strCommand", Array(strCommand)
End Function
