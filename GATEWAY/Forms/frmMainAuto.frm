VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Papyrus II  Data gateway"
   ClientHeight    =   2085
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   6825
   ControlBox      =   0   'False
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      Caption         =   "Starting extraction . . . "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   1020
      TabIndex        =   0
      Top             =   615
      Width           =   4830
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oLC As z_Loyalty
Attribute oLC.VB_VarHelpID = -1
Dim WithEvents oSplit As z_Split
Attribute oSplit.VB_VarHelpID = -1
Dim XA As XArrayDB
Dim pCLose As Boolean
Dim oTF As New z_TextFile
Dim lngCount As Long
Dim bWorkDone As Boolean
Dim oEx As z_Export



Private Sub oSplit_Report(pMsg As String)
    oTF.WriteToTextFile pMsg
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Load", , EA_NORERAISE
    HandleErrorQuiet pCLose
    If pCLose Then Unload Me
End Sub
Private Sub ClearOldLogs(pDirection As String)
    On Error GoTo errHandler
Dim oFSO As New FileSystemObject
Dim fol, fc, f
Dim strDirection As String
    strDirection = UCase(pDirection) & "LOG"
    Set fol = oFSO.GetFolder(oPC.SharedFolderRoot)
    Set fc = fol.files
    For Each f In fc
        If UCase(Left(f.Name, Len(strDirection))) = strDirection Then
            If DateDiff("d", f.DateCreated, Date) > 7 Then
                f.Delete True
            End If
        End If
    Next
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ClearOldSendLogs(pDirection)", pDirection
End Sub

Public Sub DoWork()
Dim oSQL As z_SQL
Dim lngOpID As Long
Dim strFilename As String
Dim lngRecords As Long
Dim bUpdateLCE As Boolean
Dim f
Dim oFSO As New FileSystemObject
Dim dteMostRecent As Date

    On Error GoTo errHandler
    'oPC.OpenDBShort
        Set oEx = New z_Export
        If UCase(strCL) = "SEND" Then
            ClearOldLogs "SEND"
            Me.lblProgress.Caption = "SENDING . . . "
            Me.Refresh
            DoEvents
            Set oSQL = New z_SQL
            oEx.Component oTF
            oTF.OpenTextFile oPC.SharedFolderRoot & "\SENDLOG" & Format(Date, "yyyymmdd") & ".txt"
            oTF.WriteToTextFile "Connecting  . . ." & Format(Now, "HH:NN")
                
            oEx.Connect
'----------------------------
            If oPC.Configuration.NielsenActive Then
                lngOpID = oSQL.StartOperation(Date, 0, NielsenSales)
                'Extract all data
                        oTF.WriteToTextFile "Preparing sales data  . . ." & Format(Now, "HH:NN")
                Set oSplit = New z_Split
                oSplit.ExportNielsentoFile dteMostRecent
                Set oSplit = Nothing
                
                If oEx.SendNielsen(oPC.ClientCode & Format(Now(), "yyyymmddhhnn") & ".ZIP") = True Then
                    oPC.COShort.Execute "Update tNielsen Set N_LastDateSalesSent = '" & ReverseDate(dteMostRecent) & "'"
                    oSQL.CompleteOperation lngOpID, True
                Else
                    oSQL.CompleteOperation lngOpID, False
                End If
            Else
                        oTF.WriteToTextFile "Sales data export INACTIVE" & Format(Now, "HH:NN")
            End If
'----------------------------
'            If oPC.Configuration.StockSharingACtive Then
'                        oTF.WriteToTextFile "Preparing stocksharing data  . . ." & Format(Now, "HH:NN")
'                Set oLC = New z_Loyalty
'                oLC.Component oPC
'                oLC.CreateStockSharingExtractionFile
'                Set oLC = Nothing
'
'                lngOpID = oSQL.StartOperation(Date, 0, StockSharing)
'                If oEx.SendStockSharing = True Then
'                    oSQL.CompleteOperation lngOpID, True
'                Else
'                    oSQL.CompleteOperation lngOpID, False
'                End If
'            Else
'                        oTF.WriteToTextFile "Stock sharing INACTIVE" & Format(Now, "HH:NN")
'            End If
            
        'Transmit all prepared files
            Me.Refresh
'-----------------------------
'            If oPC.Configuration.LoyaltySchemeActive Then
'        'Get receipts for files received at Central and delete the files from the local folder so they don't get sent again
'                            oTF.WriteToTextFile "Fetching receipts  . . ." & Format(Now, "HH:NN")
'
'                oEx.FetchLCResponses 'get the .LCR (receipts) and .LCE (edited records) files
'                oEx.DeleteReceipted
'
'        'Send all changes to loyalty customers to Central
'                Set oLC = New z_Loyalty
'                oLC.Component oPC
'                            oTF.WriteToTextFile "Preparing loyalty data  . . ." & Format(Now, "HH:NN")
'                oLC.CreateLoyaltyExtractionFile
'                Set oLC = Nothing
'
'                lngOpID = oSQL.StartOperation(Date, 0, LoyaltyScheme)
'                If oEx.SendLoyalty() Then
'                    oSQL.CompleteOperation lngOpID, True
'                Else
'                    oSQL.CompleteOperation lngOpID, False
'                End If
                
        'Update local database from the fetched .LCE files (only if Backup of DB taken)
'                Set f = oFSO.GetFile(oPC.SharedFolderRoot & "\BU\PBKS.BAK")
'                If Not f Is Nothing Then
'                    If DateDiff("d", f.DateLastModified, Date) < 1 Then
'                        oEx.UpdateFromEditedLC  'this should only happen if a backup has been made by the dayend run
'                    Else
'                        oTF.WriteToTextFile "Cannot update LCE - no backup taken." & Format(Now, "HH:NN")
'                    End If
'                End If
'            Else
'                            oTF.WriteToTextFile "Loyalty scheme INACTIVE" & Format(Now, "HH:NN")
'            End If
'-----------------------------
'            If oPC.Configuration.AuditingActive Then
'
'
'
'
'            End If
            
            oEx.Hangup
            oTF.CloseTextFile
            Set oSQL = Nothing
            Set oEx = Nothing
            Set oSplit = Nothing
            Set oTF = Nothing
'----------------------------
'----------------------------
        ElseIf UCase(strCL) = "FETCH" Then
            lblProgress.Caption = "FETCHING . . . "
            Me.Refresh
            DoEvents
            ClearOldLogs "FETCH"
            
            oEx.Component oTF
            oTF.OpenTextFile oPC.SharedFolderRoot & "\FETCHLOG" & Format(Date, "yyyymmdd") & ".txt"
            oTF.WriteToTextFile "Connecting  . . ." & Format(Now, "HH:NN")
                
            
            oEx.Connect
            FetchFromCentral
            bWorkDone = True
            oEx.Hangup
            
            oEx.UpdateStockSharing
            
            
            oTF.CloseTextFile
        End If
        Set oEx = Nothing
        Set oTF = Nothing
        
        oPC.DisconnectDBShort
    Exit Sub
errHandler:
    ErrPreserve
    If Not oEx Is Nothing Then
        oEx.Hangup
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.DoWork", , EA_NORERAISE
    oPC.DisconnectDBShort
    HandleErrorQuiet pCLose
    If pCLose Then Unload Me
End Sub

Private Sub FetchFromCentral()
    On Error GoTo errHandler

  '  MsgBox "FetchFromCentral"
    If oPC.Configuration.StockSharingACtive Then
        oEx.FetchStockSharing
    End If
      
    
    Exit Sub
errHandler:
    ErrorIn "frmMain.FetchFromCentral"
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
    Set XA = New XArrayDB
    bWorkDone = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Initialize", , EA_NORERAISE
    HandleErrorQuiet pCLose
    If pCLose Then Unload Me
End Sub


