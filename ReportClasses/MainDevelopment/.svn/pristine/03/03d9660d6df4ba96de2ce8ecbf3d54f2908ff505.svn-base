VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportStock 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Import stock records from file"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   10
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
      Picture         =   "frmImportStock.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   4005
      Width           =   5025
   End
   Begin VB.CommandButton cmdImportErrors 
      BackColor       =   &H00C4BCA4&
      Caption         =   "View data translation error log (Bulk Import)"
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
      TabIndex        =   7
      Top             =   3075
      Width           =   5025
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
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   975
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
      Left            =   4695
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   495
      Width           =   465
   End
   Begin VB.TextBox txtFilePath 
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4545
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2595
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
      TabIndex        =   6
      Top             =   1365
      Width           =   1470
   End
   Begin VB.Label lblResults 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   135
      TabIndex        =   5
      Top             =   1620
      Width           =   5025
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Path to null delimited file."
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
Attribute VB_Name = "frmImportStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFileName As String
Dim fold
Dim fc
Dim f
Dim oFSO As New FileSystemObject
Dim strCommand As String
Dim oSQL As New z_SQL
Dim res As Long
Dim strErrorFilePath As String

Dim enImportType As enumImportType

Public Sub component(pImportType As enumImportType)
    enImportType = pImportType
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub


Private Sub cmdGo_Click()
    CD1.ShowOpen
    strFileName = CD1.FileName
    txtFilePath = strFileName
End Sub

Private Sub cmdImport_Click()
Dim oSQL As New z_SQL
Dim pMsg As String
Dim pErrorFilePath As String

    Select Case enImportType
    Case enStockImport
        If Not oFSO.FileExists(oPC.getProperty("StockInputFormatFilePath")) Then
            MsgBox "Template file: " & oPC.getProperty("StockInputFormatFilePath") & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Exit Sub
        End If
        oSQL.ImportStock txtFilePath, pMsg, pErrorFilePath
    Case encustomerImport
        If Not oFSO.FileExists(oPC.getProperty("CustomerInputFormatFilePath")) Then
            MsgBox "Template file: " & oPC.getProperty("CustomerInputFormatFilePath") & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Exit Sub
        End If
        oSQL.ImportCustomer txtFilePath, pMsg, pErrorFilePath
    Case enSupplierImport
        If Not oFSO.FileExists(oPC.getProperty("SupplierInputFormatFilePath")) Then
            MsgBox "Template file: " & oPC.getProperty("SupplierInputFormatFilePath") & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Exit Sub
        End If
        oSQL.ImportSupplier txtFilePath, pMsg, pErrorFilePath
    Case enStockCategoriesImport
        If Not oFSO.FileExists(oPC.getProperty("StockCategoriesInputFormatFilePath")) Then
            MsgBox "Template file: " & oPC.getProperty("StockCategoriesInputFormatFilePath") & " does not exist. Cannot continue", vbInformation + vbOKOnly, "Can't do this"
            Exit Sub
        End If
        oSQL.StockCategories txtFilePath, pMsg, pErrorFilePath
    End Select
    strErrorFilePath = pErrorFilePath
    cmdImportErrors.Enabled = oFSO.FileExists(strErrorFilePath)
    lblResults.Caption = pMsg
    If Not oFSO.FolderExists(oPC.SharedFolderRoot & "\Logs") Then
        oFSO.CreateFolder oPC.SharedFolderRoot & "\Logs"
    End If
End Sub




Private Sub cmdImportErrors_Click()
Dim sExec As String

    sExec = GetPDFExecutable(strErrorFilePath)
    If oFSO.FileExists(strErrorFilePath) Then Shell sExec & " " & strErrorFilePath, vbNormalFocus
    If oFSO.FileExists(strErrorFilePath & ".Error.TXT") Then Shell sExec & " " & strErrorFilePath & ".Error.TXT", vbNormalFocus

End Sub


Private Sub cmdSkippedReport_Click()
Dim strCommand As String
Dim strSQL As String
Dim sExec As String
Dim strSkippedFilePath As String
    
    Select Case enImportType
    Case enStockImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportStock WHERE IS_RowInsertStatus = 'S'"
    Case encustomerImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportCustomer WHERE IS_RowInsertStatus = 'S'"
    Case enSupplierImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportSupplier WHERE IS_RowInsertStatus = 'S'"
    Case enStockCategoriesImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tStockCategories WHERE IS_RowInsertStatus = 'S'"
    End Select
    
    strSkippedFilePath = oPC.SharedFolderRoot & "\Logs\SkippedRows" & Format(Now(), "ddmmyy_HHNNSS") & ".TXT"
    strCommand = "bcp """ & strSQL & """ queryout """ & strSkippedFilePath & """ -eBCPError.sal -c -t\, -q  -Usa -Pcar -S " & oPC.ServerName
    res = F_7_AB_1_ShellAndWaitSimple(strCommand)
    sExec = GetPDFExecutable(strSkippedFilePath)
    If oFSO.FileExists(strSkippedFilePath) Then Shell sExec & " " & strSkippedFilePath, vbNormalFocus

End Sub
Private Sub cmdErrors_Click()
Dim strCommand As String
Dim strSQL As String
Dim sExec As String
Dim strErrorFilePath As String
    
    Select Case enImportType
    Case enStockImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportStock WHERE IS_RowInsertStatus = 'E'"
    Case encustomerImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportCustomer WHERE IS_RowInsertStatus = 'E'"
    Case enSupplierImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tImportSupplier WHERE IS_RowInsertStatus = 'E'"
    Case enStockCategoriesImport
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.tStockCategories WHERE IS_RowInsertStatus = 'E'"
    End Select
    
    strErrorFilePath = oPC.SharedFolderRoot & "\Logs\ErrorRows" & Format(Now(), "ddmmyy_HHNNSS") & ".TXT"
    strCommand = "bcp """ & strSQL & """ queryout """ & strErrorFilePath & """ -eBCPError.sal -c -t\, -q  -Usa -Pcar -S " & oPC.ServerName
    res = F_7_AB_1_ShellAndWaitSimple(strCommand)
    sExec = GetPDFExecutable(strErrorFilePath)
    If oFSO.FileExists(strErrorFilePath) Then Shell sExec & " " & strErrorFilePath, vbNormalFocus

End Sub

