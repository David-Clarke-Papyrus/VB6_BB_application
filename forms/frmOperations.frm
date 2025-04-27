VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C8B9B3&
   Caption         =   "Papyrus II:  Management console v1.3"
   ClientHeight    =   6855
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10275
   FontTransparent =   0   'False
   Icon            =   "frmOperations.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10275
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   6555
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "a"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12488
            Key             =   "b"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "c"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1980
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   6210
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   5700
      Visible         =   0   'False
      Width           =   540
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   330
      Left            =   7860
      TabIndex        =   2
      Top             =   6255
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9585
      Top             =   5025
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":03E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":0446
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":04A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":0502
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":0560
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":05BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":061C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":067A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":06D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":0736
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":0794
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":07F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":0850
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":08AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":090C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":096A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":09C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperations.frx":0A26
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer objT 
      Enabled         =   0   'False
      Interval        =   65000
      Left            =   4470
      Top             =   5535
   End
   Begin MSComctlLib.ListView lvwOperations 
      Height          =   5910
      Left            =   45
      TabIndex        =   1
      Top             =   330
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   10425
      SortKey         =   5
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date started"
         Object.Width           =   3599
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ended"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Result"
         Object.Width           =   1412
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Operator"
         Object.Width           =   3598
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "srt"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Recent operations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   135
      TabIndex        =   0
      Top             =   75
      Width           =   1740
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuLocalBU 
         Caption         =   "Backup only to local disk"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup only to removable device"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBUNow 
         Caption         =   "Dayend now with backup to local disk"
      End
      Begin VB.Menu mnuDEBU 
         Caption         =   "Dayend now with backup to removable device"
         Visible         =   0   'False
      End
      Begin VB.Menu mnude 
         Caption         =   "Dayend without backup (not recommended)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuhyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExportImport 
         Caption         =   "Export and import (accounting)"
      End
      Begin VB.Menu mnuBrowseImportsAndExports 
         Caption         =   "Browse imports and exports (accounting)"
      End
      Begin VB.Menu mnuhyphen1a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImportFromFlatFiles 
         Caption         =   "Importing data from files (take-on)"
         Begin VB.Menu mnuImportPT 
            Caption         =   "Import product type master data from file"
         End
         Begin VB.Menu mnuImportStockCategories 
            Caption         =   "Import stock categories master data from file"
         End
         Begin VB.Menu mnuImportStock 
            Caption         =   "Import stock from file"
         End
         Begin VB.Menu mnuImportSuppliers 
            Caption         =   "Import suppliers from file"
         End
         Begin VB.Menu mnuImportCustomers 
            Caption         =   "Import customers from file"
         End
         Begin VB.Menu mnuImportSalesOrders 
            Caption         =   "Import sales orders"
         End
         Begin VB.Menu mnuImportPO 
            Caption         =   "Import purchase order"
         End
         Begin VB.Menu mnuImportGRNs 
            Caption         =   "Import GRNs"
         End
         Begin VB.Menu mnuImportWSSales 
            Caption         =   "Import Wordstock sales from file"
         End
      End
      Begin VB.Menu mnuExportToFiles 
         Caption         =   "Export data to files"
         Begin VB.Menu mnuExportPTMaster 
            Caption         =   "Export product type master"
         End
         Begin VB.Menu mnuExportPTAssignments 
            Caption         =   "Export product type assignments"
         End
         Begin VB.Menu mnuExportCategoryMaster 
            Caption         =   "Export category master"
         End
         Begin VB.Menu mnuExportCategoryAssignments 
            Caption         =   "Export category assignments"
         End
         Begin VB.Menu mnuExportMBMaster 
            Caption         =   "Export multibuy master"
         End
         Begin VB.Menu mnuExportMBAssignments 
            Caption         =   "Export multibuy assignments"
         End
         Begin VB.Menu mnuhyphen1aa 
            Caption         =   "-"
         End
         Begin VB.Menu mnuExportFromViews 
            Caption         =   "Export data from views"
         End
      End
      Begin VB.Menu mnuImportDatafromFiles 
         Caption         =   "Import data from files"
         Begin VB.Menu mnuImportPTMaster 
            Caption         =   "Import product type master"
         End
         Begin VB.Menu mnuImportPTAssignments 
            Caption         =   "Import product type assignments"
         End
         Begin VB.Menu mnuImportCategoryMaster 
            Caption         =   "Import category  master"
         End
         Begin VB.Menu mnuImportCatassignments 
            Caption         =   "Import category assignments"
         End
         Begin VB.Menu mnuImportMBMaster 
            Caption         =   "Import multibuy master"
         End
         Begin VB.Menu mnuImportMBAssignments 
            Caption         =   "Insert multibuy assignments"
         End
         Begin VB.Menu mnuImportAged 
            Caption         =   "Import aged balances from file"
         End
         Begin VB.Menu mnuWebmasterList 
            Caption         =   "Import web master list"
         End
         Begin VB.Menu mnuImportPrices 
            Caption         =   "Import selling prices from file"
         End
      End
      Begin VB.Menu mnuBookfindFeed 
         Caption         =   "Import Bookfind feed"
      End
      Begin VB.Menu mnuhyphen1b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDiscounts 
         Caption         =   "Manage product discounts"
      End
      Begin VB.Menu mnuCancelSO 
         Caption         =   "Mark purchase orders as Cancelled"
      End
      Begin VB.Menu mnuCancelCO 
         Caption         =   "Mark customer orders as Cancelled"
      End
      Begin VB.Menu mnuCancelApp 
         Caption         =   "Mark appros as Cancelled"
      End
      Begin VB.Menu mnudeletions 
         Caption         =   "Deletions"
         Begin VB.Menu mnuRemoveObsoleteCO 
            Caption         =   "Remove obsolete customer orders"
         End
         Begin VB.Menu mnuRemoveQuotations 
            Caption         =   "Remove quotations"
         End
         Begin VB.Menu mnuCasualReassign 
            Caption         =   "Reassign old invoices for casual customers to a/c 'CASUAL'"
         End
         Begin VB.Menu mnuRemoveDuplicates 
            Caption         =   "Remove product duplicates automatically"
         End
         Begin VB.Menu mnuWRitedowns 
            Caption         =   "Write-downs"
         End
      End
      Begin VB.Menu mnuSwap 
         Caption         =   "Swap supplier for selected publishers"
      End
      Begin VB.Menu mnuSectionCheck 
         Caption         =   "Section check"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "Configuration"
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBrief 
         Caption         =   "Copy database to briefcase"
      End
      Begin VB.Menu mnuBriefcaseInstall 
         Caption         =   "Install Test database from briefcase"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuDataman 
      Caption         =   "Data management"
      Begin VB.Menu mnuInit 
         Caption         =   "Initialize system data"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRemoveOldData 
         Caption         =   "Remove obsolete data"
      End
      Begin VB.Menu mnuWash 
         Caption         =   "Wash data against Bookdata"
      End
      Begin VB.Menu mnuBFPUB 
         Caption         =   "Import Bookdata Publisher data"
      End
      Begin VB.Menu mnuREAS 
         Caption         =   "Reassign EAN and product codes"
      End
      Begin VB.Menu mnuRemoveArticle 
         Caption         =   "Clean up titles and 'The','An' etc."
      End
      Begin VB.Menu mnuImportLCMaster 
         Caption         =   "Import master loyalty customer list (from Central)"
      End
      Begin VB.Menu mnuResetVATRATE 
         Caption         =   "Set all VAT rate to default on Products and documents"
      End
      Begin VB.Menu mnuPrimePastelExport 
         Caption         =   "Prime Pastel export"
      End
      Begin VB.Menu mnuRebuildindexes 
         Caption         =   "Rebuild indexes"
      End
      Begin VB.Menu mnuClearLocks 
         Caption         =   "Clear all locks on order fulfilment"
      End
      Begin VB.Menu mnuInitCosts 
         Caption         =   "Initialize cost values"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPOETA_reset 
         Caption         =   "re-set ETA dates of POs"
      End
   End
   Begin VB.Menu mnuDataTr 
      Caption         =   "Data transmission"
      Begin VB.Menu mnuPastelControl 
         Caption         =   "Pastel transmission control"
      End
      Begin VB.Menu mnuTransmissionControl 
         Caption         =   "Transmission control"
      End
   End
   Begin VB.Menu mnuMonthend 
      Caption         =   "Month end"
      Begin VB.Menu mnuPrepareStatements 
         Caption         =   "Prepare statements"
      End
      Begin VB.Menu mnuPrintStatements 
         Caption         =   "Print statements"
      End
      Begin VB.Menu mnuRunMonthEnd 
         Caption         =   "Change periods"
      End
      Begin VB.Menu mnuShowAccounts 
         Caption         =   "Show accounts"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Show operation status"
      Begin VB.Menu mnuFullreport 
         Caption         =   "Print full report"
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "Display full report"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nRet         As Long
Private nMainhWnd    As Long

'Amazon web services subscription: 04BT5EGZMMVCPMN3C6R2

Private cOperations As c_Operations
Dim WithEvents oBatch As z_Batch
Attribute oBatch.VB_VarHelpID = -1
Dim WithEvents oSQL As z_SQL
Attribute oSQL.VB_VarHelpID = -1
Dim oST As a_Statements
Dim bStartscheduler As Boolean
Dim dteNominalNextUPdate As Date
Dim fWait As Boolean
Dim WithEvents oDE As z_Dayend
Attribute oDE.VB_VarHelpID = -1
Dim bDoUpdate As Boolean
Dim oXML As zXML


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo errHandler
    If MsgBox("You want to close Papyrus II Console?", vbQuestion + vbYesNo, "Application closing") = vbNo Then
        Cancel = True
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_QueryUnload(Cancel,UnloadMode)", Array(Cancel, UnloadMode)
    HandleError
End Sub

Private Sub mnuBookfindFeed_Click()
    On Error GoTo errHandler
Dim f As New frmImportFromFile
    f.Component enBookfindFeed
    f.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBookfindFeed_Click"
    HandleError
End Sub

Private Sub mnubrief_Click()
    On Error GoTo errHandler
Dim oBU As z_PBKSBackup
Dim fs As New FileSystemObject
Dim strFilefolder As String
Dim strFileName As String

    strFilefolder = GetSetting(App.EXEName, "Console", "Briefcasefolder", "c:\PBKS\BU")
    strFileName = GetSetting(App.EXEName, "Console", "BriefcaseFilename", "PBKS.BAK")
    CD1.DialogTitle = "Save live database to file"
    CD1.InitDir = strFilefolder
    CD1.FLAGS = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer
    CD1.CancelError = True
    If Right(strFileName, 3) = "BAK" Then
        CD1.Filter = "Raw SQL Server file |*.BAK|Zipped files (*.ZIP)|*.ZIP"
    Else
        CD1.Filter = "Zipped files (*.ZIP)|*.ZIP|Raw SQL Server file |*.BAK"
    End If
    CD1.ShowOpen
    If Err = 32755 Then
        Exit Sub
    ElseIf Err <> 0 Then
        GoTo errHandler
    End If
    If CD1.FileName = "" Then
        MsgBox "You must specify an existing file name!", vbInformation, "Invalid filename"
    Else
        strFileName = CD1.FileName
    End If
    
    SaveSetting App.EXEName, "Console", "Briefcasefolder", fs.GetParentFolderName(strFileName)
    SaveSetting App.EXEName, "Console", "BriefcaseFilename", strFileName

    Me.SB1.Panels(2).Text = "Backing up database . . . "
    Set oBU = New z_PBKSBackup
    Screen.MousePointer = vbHourglass
    
    oBU.BackupToBriefcase strFileName
    
    Me.SB1.Panels(2).Text = ""
    Screen.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnubrief_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnubrief_Click"
    HandleError
End Sub


Private Sub mnuBrowseImportsAndExports_Click()
Dim rs As New ADODB.Recordset
Dim f As New frmBrowseImportExports
    
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDB
'-------------------------------
    
    rs.Open "SELECT * FROM tExportToAccountingMaster Order BY ExportDate DESC", oPC.CO, adOpenStatic
    f.Component rs
    f.Show

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
End Sub

Private Sub mnuClearLocks_Click()
    On Error GoTo errHandler
Dim oB As New z_Batch
    If MsgBox("You want to clear all locks on order fulfilement? This should only be done when everyone has finished working with their order fulfilment forms." _
    & vbCrLf & "All prepared fulfilment lists will be cleared (but can be regenerated).", vbQuestion + vbOKCancel, "Warning") = vbCancel Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    oB.RunProc "ClearAllOrderFulfilmentLocks", "", "Order fulfilment locks cleared"
    Screen.MousePointer = vbDefault
    MsgBox "Order fulfilment locks cleared.", vbInformation, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuClearLocks_Click"
    HandleError
End Sub


Private Sub mnuExportFromViews_Click()
Dim frm As New frmExportFromViews

    frm.Show vbModal
    
End Sub

'Private Sub mnuBriefcaseInstall_Click()
'Dim frm As New frmInstallFromBriefcase
'
'    frm.Show vbModal
'
'End Sub



Private Sub mnuExportImport_Click()
    On Error GoTo errHandler
Dim frm As frmImportExport

    Set frm = New frmImportExport
    frm.Show vbModal
    
    Unload frm
    DoEvents
    Screen.MousePointer = vbHourglass
    UpdateScreenData
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuExportImport_Click"
    HandleError
End Sub


Private Sub mnuExportMBAssignments_Click()
    ExportToFile "MBAssignments"
End Sub

Private Sub mnuExportMBMaster_Click()
    ExportToFile "MBMaster"
End Sub

Private Sub mnuExportPTMaster_Click()
    ExportToFile "PTMaster"
End Sub

Private Sub mnuExportPTAssignments_Click()
    ExportToFile "PTAssignments"
End Sub

Private Sub mnuExportCategoryMaster_Click()
    ExportToFile "CatMaster"
End Sub

Private Sub mnuExportCategoryAssignments_Click()
    ExportToFile "CatAssignments"
End Sub


Private Sub ExportToFile(sType As String)
Dim strFileName As String
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim fs As New FileSystemObject
Dim sSP As String
Dim sMsg As String
Dim sPath As String
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDB
'-------------------------------

    Select Case sType
        Case "PTMaster"
            sSP = "ExtractPTMaster_XML"
            sMsg = "product types"
        Case "PTAssignments"
            sSP = "ExtractPTDetail_XML"
            sMsg = "product type assignments"
        Case "CatMaster"
            sSP = "ExtractCatMaster_XML"
            sMsg = "categories"
        Case "CatAssignments"
            sSP = "ExtractCatDetail_XML"
            sMsg = "category assignments"
        Case "MBMaster"
            sSP = "ExtractMBMaster_XML"
            sMsg = "multibuys"
        Case "MBAssignments"
            sSP = "ExtractMBDetail_XML"
            sMsg = "multibuy assignments"
        Case "Inventory"
            sSP = "ExtractInventory_XML"
            sMsg = "Inventory"
        Case "Suppliers"
            sSP = "ExtractSuppliers_XML"
            sMsg = "Suppliers"
        Case "Customer"
            sSP = "ExtractCustomer_XML"
            sMsg = "Customer"
        
    End Select
    
    If MsgBox("You are exporting " & sMsg & " to a file?", vbYesNo, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
    sPath = GetSetting("PBKS", "ImportExportFolder", "PTMaster", "c:\PBKS")
    CD1.InitDir = oPC.SharedFolderRoot
    CD1.FLAGS = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer
    CD1.CancelError = False
    CD1.Filter = "Text Files (*.XML)|*.XML"
    CD1.InitDir = sPath
    CD1.ShowOpen
    If CD1.FileName = "" Then
        MsgBox "You must specify a file name!", vbInformation, "Invalid filename"
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        strFileName = CD1.FileName
    End If
    
    SaveSetting "PBKS", "ImportExportFolder", "PTMaster", fs.GetParentFolderName(strFileName)
   
    Screen.MousePointer = vbHourglass

'Extract data
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = sSP
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Path", adVarChar, adParamInput, 1000, fs.GetParentFolderName(fs.GetAbsolutePathName(strFileName)))
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@Filename", adVarChar, adParamInput, 100, fs.GetFileName(fs.GetAbsolutePathName(strFileName)))
    cmd.Parameters.Append par
    cmd.Execute
    
    Set cmd = Nothing
    
    Screen.MousePointer = vbDefault

    MsgBox "Export complete.", vbOKOnly, "Status"

End Sub

Private Sub mnuImportCatassignments_Click()
    ImportFromFile "CatAssignments"
End Sub

Private Sub mnuImportCategoryMaster_Click()
    ImportFromFile "CatMaster"
End Sub

Private Sub mnuImportMBAssignments_Click()
    ImportFromFile "MBAssignments"
End Sub

Private Sub mnuImportMBMaster_Click()
    ImportFromFile "MBMaster"
End Sub

Private Sub mnuImportPrices_Click()
    On Error GoTo errHandler
    
Dim f As New frmImportFromFile
    f.Component enSellingPricesUpdate
    f.Show
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportPrices_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuImportPTAssignments_Click()
    ImportFromFile "PTAssignments"
End Sub

Private Sub mnuImportPTMaster_Click()
    ImportFromFile "PTMaster"
End Sub

Private Sub ImportFromFile(sType As String)
Dim strFileName As String
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim fs As New FileSystemObject
Dim sSP As String
Dim sMsg As String
Dim sPath As String

    Select Case sType
        Case "PTMaster"
            sSP = "ImportPTMaster_XML"
            sMsg = "product types"
        Case "PTAssignments"
            sSP = "ImportPTDetail_XML"
            sMsg = "product type assignments"
        Case "CatMaster"
            sSP = "ImportCatMaster_XML"
            sMsg = "categories"
        Case "CatAssignments"
            sSP = "ImportCatDetail_XML"
            sMsg = "category assignments"
        Case "MBMaster"
            sSP = "ImportMBMaster_XML"
            sMsg = "multibuys"
        Case "MBAssignments"
            sSP = "ImportMBDetail_XML"
            sMsg = "multibuy assignments"
        Case "SellingPrices"
            sSP = "ImportSellingprices_XML"
            sMsg = "Selling prices"
        
    End Select
    
    If MsgBox("You are importing " & sMsg & " from a file?", vbYesNo, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
    sPath = GetSetting("PBKS", "ImportExportFolder", "PTMaster", "c:\PBKS")
    CD1.InitDir = oPC.SharedFolderRoot
    CD1.FLAGS = cdlOFNHideReadOnly Or cdlOFNFileMustExist Or cdlOFNExplorer
    CD1.CancelError = True
    CD1.Filter = "Text Files (*.XML)|*.XML"
    CD1.InitDir = sPath
    CD1.CancelError = False
    CD1.ShowOpen
    If CD1.FileName = "" Then
        MsgBox "You must specify a file name!", vbInformation, "Invalid filename"
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        strFileName = CD1.FileName
    End If
    
    SaveSetting "PBKS", "ImportExportFolder", "PTMaster", fs.GetParentFolderName(strFileName)
   
    Screen.MousePointer = vbHourglass

'Import data
    '--------------
    OpenResult = oPC.OpenDBSHort
    '--------------
    Set cmd = New ADODB.Command
    cmd.CommandText = sSP
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FilePath", adVarChar, adParamInput, 1000, strFileName)
    cmd.Parameters.Append par
    
   
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandTimeout = 0
    cmd.Execute
    
    Set cmd = Nothing
    
    Screen.MousePointer = vbDefault
'    --------------
    If OpenResult = 0 Then oPC.DisconnectDBShort  'if the recent open command actually opened a connection then close it
'    --------------

    MsgBox "Import complete.", vbOKOnly, "Status"

End Sub


Private Sub mnuImportAged_Click()
    On Error GoTo errHandler
Dim oB As z_Batch
Dim frm As New frmConfirmMEImport_1
Dim strFileName As String

    If MsgBox("Confirm you wish to IMPORT customers' balances and interest charges from Accounting into Papyrus.", vbYesNo, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
  'Find the file containing the New month's balances and interest

    CD1.InitDir = oPC.SharedFolderRoot
    CD1.FLAGS = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer
    CD1.CancelError = True
    CD1.Filter = "Text Files (*.txt)|*.txt"
    CD1.ShowOpen
    If CD1.FileName = "" Then
        MsgBox "You must specify an existing file name!", vbInformation, "Invalid filename"
    Else
        strFileName = CD1.FileName
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    Set oB = New z_Batch
    oB.ImportDebtorsME gSTAFFID, strFileName
    Screen.MousePointer = vbDefault
    frm.Show vbModal
    If frm.Cancelled = True Then
        Unload frm
        Set oB = Nothing
        Exit Sub
    End If

    oB.CreateBFBALTransaction frm.ImportDate, frm.PeriodDescription
    Unload frm
    Set oB = Nothing
    
    Screen.MousePointer = vbDefault

    MsgBox "Import complete.", vbOKOnly, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportAged_Click"
    HandleError
End Sub

Private Sub mnuImportCustomers_Click()
    On Error GoTo errHandler
Dim f As New frmImportFromFile
    f.Component encustomerImport
    f.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportCustomers_Click"
    HandleError
End Sub

Private Sub mnuImportGRNs_Click()
    On Error GoTo errHandler
Dim f As New frmImportFromFile
    f.Component enGRNImport
    f.Show

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportGRNs_Click"
    HandleError
End Sub

Private Sub mnuImportLCMaster_Click()
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
Dim strFilefolder As String
Dim strFileName As String
Dim fs As New FileSystemObject
Dim oSQL As New z_SQL

    If MsgBox("In order to avoid unnecessarily filling queues at Central, database triggers will be turned off while this import is done. Ensure you are doing this operation at a time when other operations are not being done by other users. Continue?", vbInformation + vbOKCancel, "Warning") = vbCancel Then
        Exit Sub
    End If
    strFilefolder = GetSetting(App.EXEName, "Console", "MasterCustomerListFolder", "c:\PBKS\BU")
    strFileName = GetSetting(App.EXEName, "Console", "MasterCustomerList", "")
    CD1.DialogTitle = "Insert and update customers from external file"
    CD1.InitDir = strFilefolder
    CD1.FLAGS = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer
    CD1.CancelError = True
    CD1.Filter = "Text files (*.TXT)|*.TXT"
    CD1.ShowOpen
    If Err = 32755 Then
        Exit Sub
    ElseIf Err <> 0 Then
        GoTo errHandler
    End If
    If CD1.FileName = "" Then
        MsgBox "You must specify an existing file name!", vbInformation, "Invalid filename"
    Else
        strFileName = CD1.FileName
    End If
    
    SaveSetting App.EXEName, "Console", "MasterCustomerListFolder", fs.GetParentFolderName(strFileName)
    SaveSetting App.EXEName, "Console", "MasterCustomerList", strFileName


    oSQL.RunProc "SwitchTriggers", Array(oPC.DatabaseName, "disable"), ""
    Set oTF = New z_TextFile
    oTF.OpenTextFileToRead strFileName
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
    oSQL.RunProc "SwitchTriggers", Array(oPC.DatabaseName, "enable"), ""
    
    MsgBox "Import of loyalty customers is complete.", vbInformation + vbOKOnly, "Status"

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuImportLCMaster_Click"
    Exit Sub
errHandler:
    ErrPreserve
    oSQL.RunProc "SwitchTriggers", Array(oPC.DatabaseName, "enable"), ""
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportLCMaster_Click"
    HandleError
    Resume
End Sub

Private Sub mnuImportPO_Click()
    On Error GoTo errHandler
Dim f As New frmImportFromFile
    f.Component enPurchaseOrdersImport
    f.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportPO_Click"
    HandleError
End Sub

Private Sub mnuImportPT_Click()
    On Error GoTo errHandler
Dim f As New frmImportFromFile
    f.Component enPTImport
    f.Show

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportPT_Click"
    HandleError
End Sub


Private Sub mnuImportSalesOrders_Click()
    On Error GoTo errHandler
Dim f As New frmImportFromFile
    f.Component enSalesOrdersImport
    f.Show

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportSalesOrders_Click"
    HandleError
End Sub

Private Sub mnuImportStock_Click()
    On Error GoTo errHandler
Dim f As New frmImportFromFile
    f.Component enStockImport
    f.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportStock_Click"
    HandleError
End Sub

Private Sub mnuImportStockCategories_Click()
    On Error GoTo errHandler
Dim f As New frmImportFromFile
    f.Component enStockCategoriesImport
    f.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportStockCategories_Click"
    HandleError
End Sub

Private Sub mnuImportSuppliers_Click()
    On Error GoTo errHandler
Dim f As New frmImportFromFile
    f.Component enSupplierImport
    f.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportSuppliers_Click"
    HandleError
End Sub

Private Sub mnuImportWSSales_Click()
    On Error GoTo errHandler
Dim f As New frmImportFromFile
    f.Component enWSSalesImport
    f.Show

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportWSSales_Click"
    HandleError
End Sub


Private Sub mnuLocalBU_Click()
    On Error GoTo errHandler
Dim oBU As z_PBKSBackup
Dim fs As New FileSystemObject

    SB1.Panels(2).Text = "Backing up database . . . "
    Screen.MousePointer = vbHourglass
    DoEvents
    Set oBU = New z_PBKSBackup
    Check oBU.BackupToLocal, EXC_GENERAL, "Backup was not successful. Contact support person"
    Check (fs.FileExists(oPC.BackupFolder & "PBKS.BAK") And fs.FileExists(oPC.BackupFolder & "PBKSMASTER.BAK")), EXC_GENERAL, "Backup was not successful.Contact support person"
    Set oBU = Nothing
    SB1.Panels(2).Text = ""
    Screen.MousePointer = vbDefault
    
    MsgBox "Local backup complete.", vbOKOnly + vbInformation, "Status"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuLocalBU_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuLocalBU_Click"
    HandleError
End Sub

Private Sub mnuBackup_Click()
    On Error GoTo errHandler
    Me.SB1.Panels(2).Text = "Backing up database . . . "
    Set oDE = New z_Dayend
    Screen.MousePointer = vbHourglass
    DoEvents
    oDE.Backup
    SB1.Panels(2).Text = ""
    MsgBox "Backup to removable disk complete.", vbOKOnly + vbInformation, "Status"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuBackup_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBackup_Click"
    HandleError
End Sub

Private Sub mnuBUNow_Click()
    On Error GoTo errHandler
Dim oBU As z_PBKSBackup
Dim fs As New FileSystemObject
Dim oB As New z_Batch
    If SecurityControlforSupervisor Then
        SB1.Panels(2).Text = "Running dayend . . . "
        oB.RunPapyrusSchedulerDayend
                
'        Set oBU = New z_PBKSBackup
'        Check oBU.BackupToLocal, EXC_GENERAL, "Backup was not successful.Contact support person"
'        Check (fs.FileExists(oPC.BackupFolder & "PBKS.BAK") And fs.FileExists(oPC.BackupFolder & "PBKSMASTER.BAK")), EXC_GENERAL, "Backup was not successful.Contact support person"
'        Set oBU = Nothing
'
'        SB1.Panels(2).Text = "Running dayend . . . "
'        Set oDE = New z_Dayend
'        oDE.DayendUpdate gSTAFFID
        SB1.Panels(2).Text = ""
        MsgBox "Dayend and backup finished: Copy your backup file to a safe place'"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBUNow_Click"
    HandleError
End Sub

Private Sub mnuDE_Click()
    On Error GoTo errHandler

    If SecurityControlforSupervisor Then
        Set oDE = New z_Dayend
        oDE.DayendUpdate gSTAFFID
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuDE_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuDE_Click"
    HandleError
End Sub

Private Sub mnuDEBU_Click()
    On Error GoTo errHandler
    
    If SecurityControlforSupervisor Then
        Set oDE = New z_Dayend
        oDE.RunDayend gSTAFFID
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuDEBU_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuDEBU_Click"
    HandleError
End Sub

Private Sub mnuDiscounts_Click()
    On Error GoTo errHandler
Dim frm As frmProductMarketing

    If SecurityControlforSupervisor Then
        Set frm = New frmProductMarketing
        frm.Show vbModal
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuDiscounts_Click"
    HandleError
End Sub




Private Sub mnuPastelControl_Click()
    On Error GoTo errHandler
Dim f As New frmAccountingExportManagement
    f.Show

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPastelControl_Click"
    HandleError
End Sub

Private Sub mnuPOETA_reset_Click()
Dim f As New frmPOETAChange
    f.Show vbModal
End Sub

Private Sub mnuPrepareStatements_Click()
    On Error GoTo errHandler
    
    If MsgBox("Confirm you want to prepare the statements." & vbCrLf & "This could take a few minutes.", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set oST = New a_Statements
    
    
    oST.PrepareStatements oPC.getProperty("OnlyActiveAccounts") = "TRUE"
    Set oBatch = Nothing
    
    Screen.MousePointer = vbDefault
    MsgBox "Statement files have been created. They are ready to print", vbInformation, "Status"
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuPrepareStatements_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPrepareStatements_Click"
    HandleError
End Sub


Private Sub mnuPrintStatements_Click()
    On Error GoTo errHandler
Dim oFS As New FileSystemObject
Dim fol, fil, f
    If MsgBox("Confirm you want to print the statements.", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Me.SB1.Panels(2) = "Statements are being printed . . . "
    DoEvents
    Set fol = oFS.GetFolder(oPC.SharedFolderRoot & "\Statements\")
    Set fil = fol.Files
    For Each f In fil
        Set oXML = New zXML
        oXML.PrintXML oPC.SharedFolderRoot & "\Statements\" & f.Name, oPC.SharedFolderRoot & "\TEMP", oPC.SharedFolderRoot & "\Templates\", oPC.LocalFolder & "\Executables", True
        Set oXML = Nothing
    Next
    Screen.MousePointer = vbDefault
    Me.SB1.Panels(2) = ""
    MsgBox "The statements have finished printing.", vbInformation, "Status"
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPrintStatements_Click"
    HandleError
End Sub

Private Sub mnuRebuildindexes_Click()
    On Error GoTo errHandler
Dim oDMO As New z_SQLDMO
    Screen.MousePointer = vbHourglass
    Me.SB1.Panels(2).Text = "Rebuilding indexes . . ."
    Me.Refresh
    oDMO.RebuildIndexes
    Screen.MousePointer = vbDefault
    Me.SB1.Panels(2).Text = ""

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuRebuildindexes_Click"
    HandleError
End Sub

Private Sub mnuRemoveArticle_Click()
    On Error GoTo errHandler
    
    If MsgBox("Confirm you want to clean up the titles by removing the preceding article (The,An,A) from the title and storing separately." & vbCrLf & "Note this is normally done automatically and there should generally be no need to do this.", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass

    oPC.CO.Execute "Execute sp_Cleanup"
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuRemoveArticle_Click"
    HandleError
End Sub
Private Sub mnuPrimePastelExport_Click()
    On Error GoTo errHandler
    
    If MsgBox("Confirm you want to prime the Pastel export." & vbCrLf & "Note normally is only done once.", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass

    oPC.CO.Execute "Execute PrimePastelExport"
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPrimePastelExport_Click"
    HandleError
End Sub

Private Sub mnuRemoveObsoleteCO_Click()
    On Error GoTo errHandler
Dim ofrm As New frmRemoveObsoleteCustOrders

    If SecurityControlforSupervisor Then
        ofrm.Show vbModal
        Unload ofrm
        Set ofrm = Nothing
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuRemoveObsoleteCO_Click"
    HandleError
End Sub
Private Sub mnuRemoveQuotations_Click()
Dim ofrm As New frmRemoveObsoleteQuotations

    If SecurityControlforSupervisor Then
        ofrm.Show vbModal
        Unload ofrm
        Set ofrm = Nothing
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuRemoveQuotations_Click"
    HandleError

End Sub
Private Sub mnuCasualReassign_Click()
Dim ofrm As New frmReAssignOldInvoices

    If SecurityControlforSupervisor Then
        ofrm.Show vbModal
        Unload ofrm
        Set ofrm = Nothing
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCasualReassign_Click"
    HandleError

End Sub


Private Sub mnuRunMonthEnd_Click()
    On Error GoTo errHandler
Dim frm As New frmPeriodSwitch
    frm.Show vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuRunMonthEnd_Click"
    HandleError
End Sub


Private Sub mnuTransmissionControl_Click()
    On Error GoTo errHandler
Dim frm As New frmTransmissionControl
    frm.Show vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuTransmissionControl_Click"
    HandleError
End Sub

Private Sub mnuWebmasterList_Click()
Dim oB As z_Batch
Dim frm As New frmConfirmMEImport_1
Dim strFileName As String

    
  'Find the file containing the New month's balances and interest

    CD1.InitDir = oPC.SharedFolderRoot
    CD1.FLAGS = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer
    CD1.CancelError = True
    CD1.Filter = "Text Files (*.txt)|*.txt"
    CD1.ShowOpen
    If CD1.FileName = "" Then
        MsgBox "You must specify an existing file name!", vbInformation, "Invalid filename"
    Else
        strFileName = CD1.FileName
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    Set oB = New z_Batch
    oB.ImportWebMasterList strFileName
    Screen.MousePointer = vbDefault

   ' oB.PrepareAllDataForWebMasterList
    
    Unload frm
    
    Set oB = Nothing
    
    Screen.MousePointer = vbDefault

    MsgBox "Import complete.", vbOKOnly, "Status"

End Sub

Private Sub mnuWRitedowns_Click()
    On Error GoTo errHandler
Dim frm As New frmWriteDown
    frm.Component
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuWRitedowns_Click"
    HandleError
End Sub

Private Sub oBatch_Progress(lngPos As Long, lngMax As Long)
    On Error GoTo errHandler
    If lngPos Mod 100 = 0 Then
        Me.SB1.Panels(2).Text = "       Record " & CStr(lngPos) & " of " & CStr(lngMax)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oBatch_Progress(lngPos,lngMax)", Array(lngPos, lngMax)
    HandleError
End Sub
Private Sub oBatch_ProgressB(lngPos As Long, lngMax As Long, pMsg As String)
    On Error GoTo errHandler
        Me.SB1.Panels(2).Text = pMsg & CStr(lngPos) & " of " & CStr(lngMax)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oBatch_ProgressB(lngPos,lngMax,pMsg)", Array(lngPos, lngMax, pMsg)
    HandleError
End Sub

Private Sub oBatch_Status(msg As String)
    On Error GoTo errHandler
    If msg = "DUP" Then
        MsgBox "The operation cannot complete because there are duplicate product codes in the database." _
        & vbCrLf & "Use the Reports application to report the duplicates, correct them, then restart this operation.", vbInformation, "Can't do this"
    Else
        Me.SB1.Panels(2).Text = msg
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oBatch_Status(msg)", msg
    HandleError
End Sub

Private Sub oDE_Status(pMsg As String, pErr As Boolean)
    On Error GoTo errHandler
        
    Screen.MousePointer = vbDefault
    
    Me.SB1.Panels(2).Text = pMsg
    Me.Refresh
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.oDE_Status(pMsg)", pMsg, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oDE_Status(pMsg,pErr)", Array(pMsg, pErr)
    HandleError
End Sub
Private Sub oDE_COMPLETE()
    On Error GoTo errHandler
    UpdateScreenData
    Set oDE = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oDE_COMPLETE"
    HandleError
End Sub

Private Sub mnuInit_Click()
    On Error GoTo errHandler
    If SecurityControlforSupervisor Then
        oPC.CO.Execute "EXEC PBKS_INITIALIZEDATA"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuInit_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuInit_Click"
    HandleError
End Sub

Private Sub UpdateScreenData()
    On Error GoTo errHandler
    Me.SB1.Refresh
    Set cOperations = Nothing
    Set cOperations = New c_Operations
    cOperations.Load
    FillOperationsList
    Set cOperations = Nothing
    MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.UpdateScreenData"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.UpdateScreenData"
End Sub

Private Sub lvwOperations_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
    Cancel = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.lvwOperations_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
'         EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.lvwOperations_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString)
    HandleError
End Sub

Private Sub lvwOperations_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.lvwOperations_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.lvwOperations_BeforeLabelEdit(Cancel)", Cancel
    HandleError
End Sub

Private Sub lvwOperations_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
Dim objOp As New a_Operation
Dim lngResult As Long

   If Button = 2 Then

        PopupMenu mnuPrint
   End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.lvwOperations_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
'         EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.lvwOperations_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y)
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler

    GetThunder
    If UBound(arCommandLine) > 0 Then
    If arCommandLine(1) <> "N" Then
        BackColor = vbRed
    Else
        Me.BackColor = &HC8B9B3
    End If
    End If
    bStartscheduler = False
    objT.Enabled = bStartscheduler
    
    Me.SB1.Panels("a") = "Last day-end: " & oPC.Configuration.LastUpdateDateF & "   "
    Me.SB1.Panels("b") = ""  '   " & oPC.NewQuotation
    Me.SB1.Panels("c") = "   " & IIf(oPC.DatabaseName <> "PBKS", "Server:" & oPC.servername & ", Database:" & oPC.DatabaseName, "Server:" & oPC.servername)
    SB1.Panels("b").ToolTipText = SB1.Panels("b").Text
    
    Set cOperations = New c_Operations
    cOperations.Load
'    mnuExportImport.Visible = oPC.Configuration.AccountingApplicationName <> "NONE"
    mnuMonthend.Visible = oPC.RunsAccountsTF
    Me.Refresh
    FillOperationsList
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Load"
    HandleError
End Sub



Private Sub mnuBFPUB_Click()
    On Error GoTo errHandler
    If SecurityControlforSupervisor Then
        Set oBatch = New z_Batch
        oBatch.createpublisherlist
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuBFPUB_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBFPUB_Click"
    HandleError
End Sub

Private Sub mnuCancelSO_Click()
    On Error GoTo errHandler
Dim ofrm As New frmMarkOrdersAsCancelled
    If SecurityControlforSupervisor Then
        ofrm.Show vbModal
        Unload ofrm
        Set ofrm = Nothing
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuCancelSO_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCancelSO_Click"
    HandleError
End Sub
Private Sub mnuCancelCO_Click()
    On Error GoTo errHandler
Dim ofrm As New frmMarkCustOrdersAsCancelled
    If SecurityControlforSupervisor Then
        ofrm.Show vbModal
        Unload ofrm
        Set ofrm = Nothing
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuCancelCO_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCancelCO_Click"
    HandleError
End Sub
Private Sub mnuCancelAPP_Click()
    On Error GoTo errHandler
Dim ofrm As New frmMarkApprosAsCancelled
    If SecurityControlforSupervisor Then
        ofrm.Show vbModal
        Unload ofrm
        Set ofrm = Nothing
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuCancelAPP_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCancelAPP_Click"
    HandleError
End Sub
Private Sub mnuConfig_Click()
    On Error GoTo errHandler
Dim frm As New frmConfiguration
    If SecurityControlforSupervisor Then
        frm.Component oPC.Configuration
        frm.Show vbModal
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuConfig_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuConfig_Click"
    HandleError
End Sub

Private Sub mnuDisplay_Click()
    On Error GoTo errHandler
Dim objOp As New a_Operation
Dim lngResult As Long
Dim str As String

    objOp.Load lngResult, val(lvwOperations.SelectedItem.Key)
    str = IIf(objOp.TypeName = "DailySales", "Run on: " & objOp.StartedAt & vbCrLf & "for dayend dated: " & objOp.NominalDate, " - done on: " & objOp.StartedAt)
    MsgBox "Report for operation: " & Trim(objOp.TypeDesc) & " " & str & vbCrLf & objOp.Fullreport & vbCrLf & "Operation result: " & IIf(objOp.Result = 1, "Successful", "Failed")
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuDisplay_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuDisplay_Click"
    HandleError
End Sub

Private Sub mnuExit_Click()
    On Error GoTo errHandler
    oPC.DisConnect
    
    Set oBatch = Nothing
    Set oDE = Nothing
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuExit_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuExit_Click"
    HandleError
End Sub

Private Sub mnuFullreport_Click()
    On Error GoTo errHandler
Dim objOp As New a_Operation
Dim lngResult As Long
Dim rpt As New rptStatus
Dim str As String
    objOp.Load lngResult, val(lvwOperations.SelectedItem.Key)
    str = IIf(objOp.TypeName = "DailySales", "Run on: " & objOp.StartedAt & vbCrLf & "for dayend dated: " & objOp.NominalDate, " - done on: " & objOp.StartedAt)
    rpt.lblHead = "Report for operation: " & Trim(objOp.TypeName) & str
    rpt.txt1 = Trim(objOp.Fullreport)
    rpt.txtResult = "Operation " & IIf(objOp.Result = 1, "SUCCESSFUL", "FAILED")
    
    rpt.PrintReport False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuFullreport_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuFullreport_Click"
    HandleError
End Sub

'Private Sub mnuRestoreFromBackup_Click()
'    Restore
'End Sub

'Private Sub mnuStrip_Click()
'Dim oBatch As z_Batch
'Dim lngResult As Long
'Dim oOp As a_Operation
'Dim strTemp As String
'Dim frm As frmConfirmStrip
'    Set frm = New frmConfirmStrip
'    frm.Show vbModal
'    If frm.GetResult = 1 Then
'        Set oOp = New a_Operation
'        oOp.BeginEdit
'        oOp.OperatorID = frm.OperatorID
'        Unload frm
'        Me.MousePointer = vbHourglass
'        strTemp = Me.SB1.Panels(2).Text
'        Me.SB1.Panels(2).Text = "Removing 'A', 'The','Die' etc. from start of titles"
'        oOp.StartedAt = Now()
'        oOp.TypeID = SimplifyTitles
'        Set oBatch = New z_Batch
'        lngResult = oBatch.RunProc("StripArticlesFromTitlesStart_1", Array(), "Removing 'A', 'The','Die' etc. from start of titles")
'        lngResult = oBatch.RunProc("StripArticlesFromTitlesStart_2", Array(), "Removing 'A', 'The','Die' etc. from start of titles")
'        lngResult = oBatch.RunProc("StripArticlesFromTitlesStart_3", Array(), "Removing 'A', 'The','Die' etc. from start of titles")
'        lngResult = oBatch.RunProc("StripArticlesFromTitlesStart_4", Array(), "Removing 'A', 'The','Die' etc. from start of titles")
'        lngResult = oBatch.RunProc("StripArticlesFromTitlesStart_5", Array(), "Removing 'A', 'The','Die' etc. from start of titles")
'        oOp.Endedat = Now()
'        oOp.NominalDate = Date
'        oOp.Result = 1
'        oOp.ApplyEdit lngResult
'        Set oOp = Nothing
'        Set oBatch = Nothing
'        Me.SB1.Panels(2).Text = strTemp
'        Me.MousePointer = vbDefault
'        Set cOperations = Nothing
'        Set cOperations = New c_Operations
'        cOperations.Load
'        FillOperationsList
'        MsgBox "Finished", vbOKOnly, "Status"
'    End If
'End Sub

Private Sub mnuRemoveOldData_Click()
    On Error GoTo errHandler
Dim frm As New frmCleanOldData
Dim frmS As New frmSecurity
Dim strName As String
Dim lngOperatorID As Long
    If MsgBox("Before removing obsolete data you should make a special copy of the database. Contact Papyrus Services.", vbCritical + vbOKCancel, "Warning") = vbCancel Then Exit Sub
    If SecurityControlforSupervisor Then
        Me.SB1.Panels(2).Text = "Removing obsolete data . . . "
        frm.Component lngOperatorID
        frm.Show vbModal
        Set cOperations = Nothing
        Set cOperations = New c_Operations
        cOperations.Load
        FillOperationsList
        SB1.Panels(2).Text = ""
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuRemoveOldData_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuRemoveOldData_Click"
    HandleError
End Sub

Private Sub mnuResetVATRATE_Click()
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

    If MsgBox("You want to set all VAT rate to default? (" & oPC.Configuration.VATRateF & ")", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    If SecurityControlforSupervisor Then
        Screen.MousePointer = vbHourglass
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.CO
        cmd.CommandTimeout = 600
          
        cmd.CommandText = "ResetVATRATE"
        cmd.CommandType = adCmdStoredProc
        
        cmd.Execute
        Screen.MousePointer = vbDefault
        MsgBox "Done"
    End If
    Exit Sub
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuResetVATRATE_Click"
    HandleError
End Sub
'Private Sub mnuInitCosts_Click()
'Dim cmd As adodb.Command
'Dim prm As adodb.Parameter
'
'    If MsgBox("You want to set Product costs to a percentage of SP where they are presently ? (" & oPC.Configuration.VATRateF & ")", vbQuestion + vbYesNo) = vbNo Then
'        Exit Sub
'    End If
'
'    If SecurityControlforSupervisor Then
'        Screen.MousePointer = vbHourglass
'        Set cmd = New adodb.Command
'        cmd.ActiveConnection = oPC.CO
'        cmd.CommandTimeout = 600
'
'        cmd.CommandText = "ResetVATRATE"
'        cmd.CommandType = adCmdStoredProc
'
'        cmd.Execute
'        Screen.MousePointer = vbDefault
'        MsgBox "Done"
'    End If
'    Exit Sub
'
'End Sub

Private Sub mnuSwap_Click()
    On Error GoTo errHandler
Dim frm As frmChangeSupplier
    Set frm = New frmChangeSupplier
    frm.Show vbModal
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuSwap_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSwap_Click"
    HandleError
End Sub

Private Sub mnuWash_Click()
    On Error GoTo errHandler
Dim frm As frmWash

    Set oBatch = New z_Batch
    If Not InStr(1, oPC.Configuration.LookupSeq, "BF") > 0 Then
        MsgBox "This application does not use Bookfind"
        Exit Sub
    End If
    If MsgBox("This procedure will take some hours and should be run when trading has ended." & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    If SecurityControlforSupervisor Then
        Set frm = New frmWash
        frm.Show vbModal
        If frm.Cancelled Then
            Unload frm
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        PB1.Visible = True
        oBatch.UpdateFromBookfind frm.Author = 1, _
        frm.Title = 1, frm.Subtitle = 1, _
        frm.Availability = 1, _
        frm.Binding = 1, frm.Edition = 1, frm.SupplierCode = 1, frm.Publishername = 1, frm.SeriesTitle = 1, _
        frm.PublicationDate = 1, _
        frm.UKPrice = 1, frm.RRP = 1, frm.BIC = 1, frm.BookStatus = 1, gSTAFFID
        Screen.MousePointer = vbDefault
        PB1.Visible = False
        Unload frm
        Set oBatch = Nothing
        UpdateScreenData
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuWash_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuWash_Click"
    HandleError
End Sub

Private Sub mnuReas_Click()
    On Error GoTo errHandler
Dim lngRecordsUPdated As Long
Dim oSQL As New z_SQL
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

    If MsgBox("This procedure will renumber the non-ISBN codes in the database, sequencing them to remove gaps. It will take some time." & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    If SecurityControlforSupervisor Then

        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.CO
        cmd.CommandTimeout = 0
          
        cmd.CommandText = "ReassignCodes"
        cmd.CommandType = adCmdStoredProc
        
        Set prm = cmd.CreateParameter("@MaxUsed", adInteger, adParamOutput)
        cmd.Parameters.Append prm
        prm.Value = lngRecordsUPdated
        cmd.Execute
        MsgBox "Last number used: " & cmd.Parameters(0) & vbCrLf & "Print the report for duplicate EAN codes using the Reports application and correct any duplicates found.", vbInformation, "Procedure ended"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuReas_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuReas_Click"
    HandleError
End Sub
Private Sub mnuRemoveDuplicates_Click()
    On Error GoTo errHandler
Dim lngRecordsUPdated As Long
Dim oSQL As New z_SQL
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

    If MsgBox("This procedure should only be executed if you are familiar with it." & vbCrLf & "Continue?", vbYesNo + vbQuestion, "Warning") = vbNo Then
        Exit Sub
    End If
    If SecurityControlforSupervisor Then

        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.CO
          
        cmd.CommandText = "MergeDuplicateProducts"
        cmd.CommandType = adCmdStoredProc
        
        cmd.Execute
        MsgBox "Done"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuRemoveDuplicates_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuRemoveDuplicates_Click"
    HandleError
End Sub

Private Sub FillOperationsList()
    On Error GoTo errHandler
Dim objItem As d_operation
Dim itmList As ListItem
Dim lngIndex As Long

    Me.lvwOperations.ListItems.Clear
    For lngIndex = 1 To cOperations.Count
        With objItem
            Set objItem = cOperations.Item(lngIndex)
            Set itmList = lvwOperations.ListItems.Add(Key:=Format$(objItem.ID) & " K")
            With itmList
                .Text = objItem.StartedAtFormatted
                .SubItems(1) = objItem.EndedatFormatted
                .SubItems(2) = objItem.TypeName
                .SubItems(3) = objItem.ResultName
               ' .SubItems(4) = objItem.NominalDateFormatted
                .SubItems(4) = objItem.OperatorName
                .SubItems(5) = Format(objItem.StartedAt, "yyyy/mm/dd hh:mm")
            End With
        End With
    Next

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.FillOperationsList"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.FillOperationsList"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If oPC.Connected Then
        If bStartscheduler Then
            If MsgBox("Closing Console will stop the dayend scheduler, " & vbCrLf & "do you want to close?", vbExclamation + vbYesNo + vbDefaultButton2, "Warning") = vbNo Then
                Cancel = True
            End If
        End If
        oPC.DisConnect
        Set cOperations = Nothing
    End If
    If Not frmWS Is Nothing Then
        Unload frmWS
    End If
    Set oPC = Nothing
    Set Constructor = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Unload(Cancel)", Cancel
    HandleError
End Sub

'Public Sub StartScheduler(pStart As Boolean)
'    On Error GoTo errHandler
'    bStartscheduler = pStart
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.StartScheduler(pStart)", pStart
'End Sub
'Private Function GetActualUpdateTime(pNextUPdate) As Date
'    On Error GoTo errHandler
'Dim iHoursToAdd As Integer
'Dim iMinsToAdd As Integer
'
'    iHoursToAdd = DatePart("h", oPC.Configuration.UpdateWindowStart)
'    iMinsToAdd = DatePart("n", oPC.Configuration.UpdateWindowStart)
'    If DatePart("h", oPC.Configuration.UpdateWindowStart) < 12 Then
'        iHoursToAdd = 12 + iHoursToAdd
'    End If
'    pNextUPdate = DateAdd("h", iHoursToAdd, CDate(DatePart("yyyy", pNextUPdate) & "-" & DatePart("m", pNextUPdate) & "-" & DatePart("d", pNextUPdate)))
'    GetActualUpdateTime = DateAdd("n", iMinsToAdd, pNextUPdate)
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.GetActualUpdateTime(pNextUPdate)", pNextUPdate
'End Function
'Private Function GetNextWorkingDay(DIW As Integer, pLastDate As Date) As Date
'    On Error GoTo errHandler
'    Select Case DIW
'    Case 5
'        If Weekday(pLastDate, vbMonday) = 5 Then
'            GetNextWorkingDay = DateAdd("d", 3, pLastDate)
'        Else
'            GetNextWorkingDay = DateAdd("d", 1, pLastDate)
'        End If
'    Case 6
'        If Weekday(pLastDate, vbMonday) = 6 Then
'            GetNextWorkingDay = DateAdd("d", 2, pLastDate)
'        Else
'            GetNextWorkingDay = DateAdd("d", 1, pLastDate)
'        End If
'    Case 7
'            GetNextWorkingDay = DateAdd("d", 1, pLastDate)
'    End Select
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.GetNextWorkingDay(DIW,pLastDate)", Array(DIW, pLastDate)
'End Function
'Private Function GetPreviousWorkingDay(DIW As Integer, pLastDate As Date) As Date
'    On Error GoTo errHandler
'    Select Case DIW
'    Case 5
'        If Weekday(pLastDate, vbMonday) = 1 Then
'            GetPreviousWorkingDay = DateAdd("d", -3, pLastDate)
'        ElseIf Weekday(pLastDate, vbMonday) = 7 Then
'            GetPreviousWorkingDay = DateAdd("d", -2, pLastDate)
'        Else
'            GetPreviousWorkingDay = DateAdd("d", -1, pLastDate)
'        End If
'    Case 6
'        If Weekday(pLastDate, vbMonday) = 1 Then
'            GetPreviousWorkingDay = DateAdd("d", 2, pLastDate)
'        Else
'            GetPreviousWorkingDay = DateAdd("d", -1, pLastDate)
'        End If
'    Case 7
'            GetPreviousWorkingDay = DateAdd("d", -1, pLastDate)
'    End Select
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.GetPreviousWorkingDay(DIW,pLastDate)", Array(DIW, pLastDate)
'End Function
'Private Function WorkingDay(pDate As Date)
'    On Error GoTo errHandler
'    If Weekday(pDate, vbMonday) >= 1 And Weekday(pDate, vbMonday) <= oPC.Configuration.DaysInWeek Then
'        WorkingDay = True
'    Else
'        WorkingDay = False
'    End If
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.WorkingDay(pDate)", pDate
'End Function
'
'Function InWindow() As Boolean
'    On Error GoTo errHandler
'Dim dteFrom As Date
'Dim dteTo As Date
'
'    If fWait = True Then
'        InWindow = False
'        Exit Function
'    End If
'    InWindow = False
'    dteFrom = oPC.Configuration.UpdateWindowStart
'    dteTo = oPC.Configuration.UpdateWindowEnd
'    If dteFrom > dteTo Then  'the times are from night to morning
'        If DatePart("h", Now()) * 60 + DatePart("n", Now()) > DatePart("h", dteFrom) * 60 + DatePart("n", dteFrom) _
'        Or DatePart("h", Now()) * 60 + DatePart("n", Now()) < DatePart("h", dteTo) * 60 + DatePart("n", dteTo) Then
'            InWindow = True
'        End If
'    Else
'        If DatePart("h", Now()) * 60 + DatePart("n", Now()) > DatePart("h", dteFrom) * 60 + DatePart("n", dteFrom) _
'        And DatePart("h", Now()) * 60 + DatePart("n", Now()) < DatePart("h", dteTo) * 60 + DatePart("n", dteTo) Then
'            InWindow = True
'        End If
'    End If
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.InWindow"
'End Function
Private Sub GetThunder()
    On Error GoTo errHandler
Dim hIcon As Long
    
    nRet = GetWindowLong(Me.hWnd, GWL_HWNDPARENT)
    Do While nRet
       nMainhWnd = nRet
       nRet = GetWindowLong(nMainhWnd, GWL_HWNDPARENT)
    Loop
    ' set the icon
 '   Set Me.Icon = Picture1.Picture
    ' get a handle to ICON_BIG
    hIcon = SendMessage(Me.hWnd, WM_GETICON, ICON_BIG, ByVal 0)
    ' send ICON_BIG to the main window
    SendMessage nMainhWnd, WM_SETICON, ICON_BIG, ByVal hIcon

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.GetThunder"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.GetThunder"
End Sub

Private Sub oZ_Status(msg As String)
    On Error GoTo errHandler
    Me.SB1.Panels(2).Text = msg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oZ_Status(msg)", msg
    HandleError
End Sub



