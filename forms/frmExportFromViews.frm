VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExportFromViews 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Export from view"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   8055
   StartUpPosition =   1  'CenterOwner
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
      Left            =   6735
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1125
      Width           =   465
   End
   Begin VB.TextBox txtDelimiter 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   405
      TabIndex        =   5
      Text            =   ","
      Top             =   1920
      Width           =   510
   End
   Begin VB.TextBox txtOutputPath 
      Height          =   315
      Left            =   405
      TabIndex        =   3
      Text            =   "txtOutputPath"
      Top             =   1140
      Width           =   6225
   End
   Begin VB.CommandButton cmdLoadData 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Export data"
      Height          =   345
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2535
      Width           =   1185
   End
   Begin VB.ComboBox cboReport_View 
      Height          =   315
      Left            =   405
      TabIndex        =   0
      Text            =   "cboDatabase_View"
      Top             =   495
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   4050
      Top             =   255
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Separator (empty yields tab-delimied file)"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   405
      TabIndex        =   6
      Top             =   1680
      Width           =   3630
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Output file path"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   405
      TabIndex        =   4
      Top             =   900
      Width           =   1635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Database view"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   390
      TabIndex        =   2
      Top             =   270
      Width           =   1635
   End
End
Attribute VB_Name = "frmExportFromViews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tl As New z_TextListSimple
Dim strFileName As String
Dim oSQL As z_SQL
Dim lngErr As Long
Dim strErrorMessage As String
Dim oFSO As FileSystemObject

Private Sub cmdLoadData_Click()
    Set oSQL = New z_SQL
    Screen.MousePointer = vbHourglass
    oSQL.ExportDataToFile cboReport_View, strFileName, IIf(txtDelimiter > "", Left(txtDelimiter, 1), ""), oPC.servername, lngErr, strErrorMessage
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    Set oFSO = New FileSystemObject
    tl.Load sltAdhocQueries
    LoadComboFromTextListSimple Me.cboReport_View, tl

End Sub
Private Sub cmdGo_Click()
    On Error GoTo errHandler
    CD.DefaultExt = ".txt"
    CD.DialogTitle = "Save export file"
    CD.InitDir = GetSetting("PBKS", "ExportToFile", "TargetFolder", oPC.SharedFolderRoot)
    CD.CancelError = True
    CD.Filter = "*.txt"
    On Error Resume Next
    CD.ShowSave
    If Err.Number = cdlCancel Then
      On Error GoTo 0
      Exit Sub
    Else
      On Error GoTo 0
    End If
    
    
    strFileName = CD.Filename
    txtOutputPath = strFileName
    SaveSetting "PBKS", "ExportToFile", "ExportToFile", oFSO.GetParentFolderName(strFileName)
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExportFromViews.cmdGo_Click", , EA_NORERAISE
    HandleError
End Sub

