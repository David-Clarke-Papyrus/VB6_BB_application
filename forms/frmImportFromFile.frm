VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportFromFile 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Import records from file"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5355
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C4BCA4&
      Caption         =   "bcp run results"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3075
      Width           =   1575
   End
   Begin VB.CommandButton cmdOpenLog 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Open log"
      Height          =   405
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4740
      Width           =   1800
   End
   Begin VB.TextBox lblResults 
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   1680
      Width           =   4995
   End
   Begin VB.CommandButton cmdErrors 
      BackColor       =   &H00C4BCA4&
      Caption         =   "View Insert errors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3540
      Width           =   5025
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4035
      Picture         =   "frmImportFromFile.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4530
      Width           =   1110
   End
   Begin VB.CommandButton cmdSkippedReport 
      BackColor       =   &H00C4BCA4&
      Caption         =   "View log of rows skipped"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4005
      Width           =   5025
   End
   Begin VB.CommandButton cmdImportErrors 
      BackColor       =   &H00C4BCA4&
      Caption         =   "View data import errors (bcp)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3075
      Width           =   3435
   End
   Begin VB.CheckBox chkAppend 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Append"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   180
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmdImport 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Import"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3015
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1185
      Width           =   1575
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00C4BCA4&
      Caption         =   "···"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4725
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   495
      Width           =   465
   End
   Begin VB.TextBox txtFilePath 
      Height          =   690
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4545
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2610
      Top             =   930
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Find file to import"
      Filter          =   "*.csv,*.txt"
      InitDir         =   "PBKS_S_"
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Results"
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
      Left            =   225
      TabIndex        =   5
      Top             =   1410
      Width           =   1470
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Path to tab-delimited file."
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
      Height          =   285
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   3630
   End
End
Attribute VB_Name = "frmImportFromFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFilename As String
Dim fold
Dim fc
Dim f
Dim oFSO As New FileSystemObject
Dim strCommand As String
Dim oSQL As New z_SQL
Dim Res As Long
Dim strErrorFilePath As String

Dim enImportType As enumImportType
Public Function FileName() As String
    FileName = strFilename
End Function
Public Sub component(pImportType As enumImportType)
    On Error GoTo errHandler
    enImportType = pImportType
    Select Case enImportType
    Case enStockImport
        Me.Caption = "Import stock data from file"
    Case encustomerImport
        Me.Caption = "Import customer data from file"
    Case enSupplierImport
        Me.Caption = "Import supplier data from file"
    Case enStockCategoriesImport
        Me.Caption = "Import stock category data from file"
    Case enSalesOrdersImport
        Me.Caption = "Import sales orders data from file"
    Case enPurchaseOrdersImport
        Me.Caption = "Import purchase orders data from file"
    Case enGRNImport
        Me.Caption = "Import goods received notes data from file"
    Case enPTImport
        Me.Caption = "Import product type data from file"
    Case enWSSalesImport
        Me.Caption = "Import Wordstock sales from file"
    Case enClipboardImport
        Me.Caption = "Import Data to Papyrus clipboard"
    Case enBookfindFeed
        Me.Caption = "Import Bookfind feed"
    Case enBankStatement
        Me.Caption = "Import bank statement"
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportFromFile.component(pImportType)", pImportType
End Sub

Private Sub CancelButton_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportFromFile.CancelButton_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdGo_Click()
    On Error GoTo errHandler
    CD1.InitDir = GetSetting("PBKS", "ImportFromFile", "SourceFolder", oPC.SharedFolderRoot)
    CD1.ShowOpen
    strFilename = CD1.FileName
    txtFilePath = strFilename
    SaveSetting "PBKS", "ImportFromFile", "SourceFolder", oFSO.GetParentFolderName(strFilename)
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportFromFile.cmdGo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdImport_Click()
    On Error GoTo errHandler
Dim oSQL As New z_SQL
Dim pMsg As String
Dim pErrorFilePath As String

    MsgBox "Ensure that no other application has the file to import open (e.g. Openoffice or Excel). " & vbCrLf & "If the import fails you should check the log file using the button at the bottom of this form", vbInformation, "Warning"
    Screen.MousePointer = vbHourglass
    Select Case enImportType
    Case enStockImport
        If Not oFSO.FileExists(oPC.GetProperty("StockInputFormatFilePath")) Then
            MsgBox "Template file: " & oPC.GetProperty("StockInputFormatFilePath") & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        oSQL.ImportFromFile "Stock", oPC.GetProperty("StockInputFormatFilePath"), txtFilePath, pMsg, pErrorFilePath
    Case encustomerImport
        If Not oFSO.FileExists(oPC.GetProperty("CustomerInputFormatFilePath")) Then
            MsgBox "Template file: " & oPC.GetProperty("CustomerInputFormatFilePath") & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        oSQL.ImportFromFile "Customer", oPC.GetProperty("CustomerInputFormatFilePath"), txtFilePath, pMsg, pErrorFilePath
    Case enSupplierImport
        If Not oFSO.FileExists(oPC.GetProperty("SupplierInputFormatFilePath")) Then
            MsgBox "Template file: " & oPC.GetProperty("SupplierInputFormatFilePath") & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        oSQL.ImportFromFile "Supplier", oPC.GetProperty("SupplierInputFormatFilePath"), txtFilePath, pMsg, pErrorFilePath
    Case enStockCategoriesImport
        If Not oFSO.FileExists(oPC.GetProperty("StockCategoryInputFormatFilePath")) Then
            MsgBox "Template file: " & oPC.GetProperty("StockCategoryInputFormatFilePath") & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        oSQL.ImportFromFile "StockCategory", oPC.GetProperty("StockCategoryInputFormatFilePath"), txtFilePath, pMsg, pErrorFilePath
    Case enSalesOrdersImport
        If Not oFSO.FileExists(oPC.GetProperty("SalesOrderInputFormatFilePath")) Then
            MsgBox "Template file: " & oPC.GetProperty("SalesOrderInputFormatFilePath") & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        oSQL.ImportFromFile "SalesOrder", oPC.GetProperty("SalesOrderInputFormatFilePath"), txtFilePath, pMsg, pErrorFilePath
    Case enPurchaseOrdersImport
        If Not oFSO.FileExists(oPC.GetProperty("PurchaseOrderInputFormatFilePath")) Then
            MsgBox "Template file: " & oPC.GetProperty("PurchaseOrderInputFormatFilePath") & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        oSQL.ImportFromFile "PurchaseOrder", oPC.GetProperty("PurchaseOrderInputFormatFilePath"), txtFilePath, pMsg, pErrorFilePath
    Case enGRNImport
        If Not oFSO.FileExists(oPC.GetProperty("GRNInputFormatFilePath")) Then
            MsgBox "Template file: " & oPC.GetProperty("GRNInputFormatFilePath") & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        oSQL.ImportFromFile "GRN", oPC.GetProperty("GRNInputFormatFilePath"), txtFilePath, pMsg, pErrorFilePath
    Case enPTImport
        If Not oFSO.FileExists(oPC.GetProperty("PTInputFormatFilePath")) Then
            MsgBox "Template file: " & oPC.GetProperty("PTInputFormatFilePath") & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        oSQL.ImportFromFile "PT", oPC.GetProperty("PTInputFormatFilePath"), txtFilePath, pMsg, pErrorFilePath
    Case enWSSalesImport
        If Not oFSO.FileExists(oPC.GetProperty("WordstockSalesInputFormatFilePath")) Then
            MsgBox "Template file: " & oPC.GetProperty("WordstockSalesInputFormatFilePath") & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        oSQL.ImportFromFile "WSSales", oPC.GetProperty("WordstockSalesInputFormatFilePath"), txtFilePath, pMsg, pErrorFilePath
    Case enClipboardImport
        If Not oFSO.FileExists(oPC.GetProperty("ClipboardInputFormatFilePath")) Then
            MsgBox "Template file: " & oPC.GetProperty("ClipboardInputFormatFilePath") & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        oSQL.ImportFromFile "Clipboard", oPC.GetProperty("ClipboardInputFormatFilePath"), txtFilePath, pMsg, pErrorFilePath
    Case enBookfindFeed
        If Not oFSO.FileExists(oPC.GetProperty("BookFindFeedInputFormatFilePath")) Then
            MsgBox "Template file: " & oPC.GetProperty("BookFindFeedInputFormatFilePath") & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        oSQL.ImportFromFile "BookfindFeed", oPC.GetProperty("BookFindFeedInputFormatFilePath"), txtFilePath, pMsg, pErrorFilePath
    Case enBankStatement
        oSQL.ImportFromFile "BankStatement", oPC.SharedFolderRoot & "\Templates\ImportSBStatementFormat.XML", txtFilePath, pMsg, pErrorFilePath
    Case enSellingPricesUpdate
        If MsgBox("You will be updating selling prices in bulk. You should have made a backup of your database before continuing." & vbCrLf & "Continue?", vbExclamation + vbYesNo, "Warning") = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        oSQL.ImportFromFile "SellingPricesUpdate", oPC.SharedFolderRoot & "\Templates\Format_SellingPricesUpdate.XML", txtFilePath, pMsg, pErrorFilePath
    
    End Select
    strErrorFilePath = pErrorFilePath
    cmdImportErrors.Enabled = oFSO.FileExists(strErrorFilePath)
    lblResults.text = pMsg
    If Not oFSO.FolderExists(oPC.SharedFolderRoot & "\Logs") Then
        oFSO.CreateFolder oPC.SharedFolderRoot & "\Logs"
    End If
    Me.Command1.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportFromFile.cmdImport_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdImportErrors_Click()
    On Error GoTo errHandler
Dim sExec As String

    sExec = GetPDFExecutable(strErrorFilePath)
    If oFSO.FileExists(strErrorFilePath) Then Shell sExec & " " & strErrorFilePath, vbNormalFocus
    If oFSO.FileExists(strErrorFilePath & ".Error.TXT") Then Shell sExec & " " & strErrorFilePath & ".Error.TXT", vbNormalFocus

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportFromFile.cmdImportErrors_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdOpenLog_Click()
    On Error GoTo errHandler
    cmdFindLogFile_Click
    Shell "NOTEPAD.EXE '" & oPC.SharedFolderRoot & "\Logs\bcp_Results.txt'", vbNormalFocus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportFromFile.cmdOpenLog_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdFindLogFile_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject

    strFilename = GetSetting("PBKS", "SB", "LOGFILEPATH", "")
    If fs.GetBaseName(strFilename) <> "ERRORLOG" Then
        CD1.DialogTitle = "Open SQL Server log file"
        CD1.DefaultExt = ""
        CD1.InitDir = "c:\Program files\Microsoft SQL SERVER"
        CD1.ShowOpen
        strFilename = CD1.FileName
        SaveSetting "PBKS", "SB", "LOGFILEPATH", strFilename
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportFromFile.cmdFindLogFile_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSkippedReport_Click()
    On Error GoTo errHandler
Dim strCommand As String
Dim strSQL As String
Dim sExec As String
Dim strSkippedFilePath As String
    
    Select Case enImportType
    Case enStockImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportStock WHERE IS_RowInsertStatus = 'S'"
    Case encustomerImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportCustomer WHERE ICUS_RowInsertStatus = 'S'"
    Case enSupplierImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportSupplier WHERE ISUP_RowInsertStatus = 'S'"
    Case enStockCategoriesImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportStockCategory WHERE ISC_RowInsertStatus = 'S'"
    Case enSalesOrdersImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportSalesOrder WHERE ISO_RowInsertStatus = 'S'"
    Case enGRNImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportGRNs WHERE IGR_RowInsertStatus = 'S'"
    Case enPTImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportProductTypes WHERE IPT_RowInsertStatus = 'S'"
    Case enWSSalesImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportWordstockSales WHERE WS_RowInsertStatus = 'S'"
    Case enClipboardImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportClipboard WHERE PC_RowInsertStatus = 'S'"
    Case enBookfindFeed
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportBookfindFeed WHERE BFF_RowInsertStatus = 'S'"
    End Select
    
    strSkippedFilePath = oPC.SharedFolderRoot & "\Logs\SkippedRows" & Format(Now(), "ddmmyy_HHNNSS") & ".TXT"
    strCommand = "bcp """ & strSQL & """ queryout """ & strSkippedFilePath & """ -eBCPError.sal -c -t\, -q  -Usa -Pcar -S " & oPC.servername
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand, vbHide, 10000)
    sExec = GetPDFExecutable(strSkippedFilePath)
     If sExec = "" Then
              MsgBox "There is no application set on this computer to open the file: " & strSkippedFilePath & ". The document cannot be displayed", vbOKOnly, "Can't do this"
     Else
            If oFSO.FileExists(strSkippedFilePath) Then
                Shell sExec & " " & strSkippedFilePath, vbNormalFocus
            End If
     End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportFromFile.cmdSkippedReport_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdErrors_Click()
    On Error GoTo errHandler
Dim strCommand As String
Dim strSQL As String
Dim sExec As String
Dim strErrorFilePath As String
    
    Select Case enImportType
    Case enStockImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportStock WHERE IS_RowInsertStatus = 'E'"
    Case encustomerImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportCustomer WHERE ICUS_RowInsertStatus = 'E'"
    Case enSupplierImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportSupplier WHERE ISUP_RowInsertStatus = 'E'"
    Case enStockCategoriesImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportStockCategory WHERE ISC_RowInsertStatus = 'E'"
    Case enSalesOrdersImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportSalesOrder WHERE ISO_RowInsertStatus = 'E'"
    Case enPurchaseOrdersImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportPurchaseOrder WHERE IPO_RowInsertStatus = 'E'"
    Case enGRNImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportGRNs WHERE IGR_RowInsertStatus = 'E'"
    Case enPTImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportProductTypes WHERE IPT_RowInsertStatus = 'E'"
    Case enWSSalesImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportWordstockSales WHERE WS_RowInsertStatus = 'E'"
    Case enClipboardImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportClipboard WHERE PC_RowInsertStatus = 'E'"
    Case enBookfindFeed
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportBookfindFeed WHERE BFF_RowInsertStatus = 'E'"
    Case enSellingPricesUpdate
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tUpdateSellingPrices WHERE ISPR_RowInsertStatus IN ('E','C','S')"
    End Select
    
    strErrorFilePath = oPC.SharedFolderRoot & "\Logs\ErrorRows" & Format(Now(), "ddmmyy_HHNNSS") & ".TXT"
    strCommand = "bcp """ & strSQL & """ queryout """ & strErrorFilePath & """ -eBCPError.sal -c -t\, -q  -Usa -Pcar -S " & oPC.servername
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand, vbHide, 10000)
    sExec = GetPDFExecutable(strErrorFilePath)
    If oFSO.FileExists(strErrorFilePath) Then Shell sExec & " " & strErrorFilePath, vbNormalFocus

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportFromFile.cmdErrors_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Command1_Click()
    On Error GoTo errHandler
    cmdFindLogFile_Click
    Shell "NOTEPAD.EXE '" & Replace(strErrorFilePath, "BulkInsertErrors", "bcp_command_results"), vbNormalFocus

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportFromFile.Command1_Click", , EA_NORERAISE
    HandleError
End Sub
