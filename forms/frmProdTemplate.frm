VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmProdTemplate 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Preview"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7815
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4605
      Width           =   1620
   End
   Begin VB.TextBox txtHeading 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   390
      Left            =   3945
      TabIndex        =   5
      Top             =   4245
      Width           =   3450
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Page layout"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   705
      Left            =   135
      TabIndex        =   2
      Top             =   3945
      Width           =   3225
      Begin VB.OptionButton optLandscape 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Landscape"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   1650
         TabIndex        =   4
         Top             =   315
         Width           =   1500
      End
      Begin VB.OptionButton optPortrait 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Portrait"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   150
         TabIndex        =   3
         Top             =   315
         Value           =   -1  'True
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Print preview"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7815
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4065
      Width           =   1620
   End
   Begin TrueOleDBGrid60.TDBGrid GT 
      Height          =   3600
      Left            =   135
      OleObjectBlob   =   "frmProdTemplate.frx":0000
      TabIndex        =   0
      Top             =   135
      Width           =   9315
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3375
      Top             =   4305
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".DOC"
      DialogTitle     =   "Save to new file"
      Filter          =   "Word Document (*.Doc)|*.doc"
   End
   Begin VB.Label lblCount 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   7500
      TabIndex        =   7
      Top             =   3750
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Heading for report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   3945
      TabIndex        =   6
      Top             =   3945
      Width           =   1935
   End
End
Attribute VB_Name = "frmProdTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XB As XArrayDB
Dim XA As XArrayDB

Public Sub component(pXA As XArrayDB, pXB As XArrayDB)
    On Error GoTo errHandler
    Set XB = pXB
    Set XA = pXA
    Me.Caption = "Preview export data"
    lblCount.Caption = XB.UpperBound(1) & " rows"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProdTemplate.component(pXA,pXB)", Array(pXA, pXB)
End Sub


Private Sub cmdExport_Click()
    On Error GoTo errHandler
Dim fs As New Scripting.FileSystemObject
    
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\Export files") Then
        fs.CreateFolder oPC.SharedFolderRoot & "\Export files"""
    End If
    CD1.InitDir = oPC.SharedFolderRoot & "\Export files"
    CD1.DefaultExt = ".txt"
    CD1.Filter = "*.txt"
    CD1.FLAGS = cdlOFNCreatePrompt
    CD1.ShowSave
    If fs.FileExists(CD1.FileName) Then
        MsgBox "You must specify a new file name!", vbInformation, "Invalid filename"
    Else
        If GT.SelBookmarks.Count > 0 Then
           ' GT.ExportToDelimitedFile CD1.FileName, dbgSelectedRows, "~"
        Else
           ' GT.ExportToDelimitedFile CD1.FileName, dbgAllRows, "~"
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProdTemplate.cmdExport_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    GT.PrintInfo.PageHeader = "\t" & FNS(Me.txtHeading)
 '   GT.PrintInfo.PageFooterFont =
    GT.PrintInfo.PageFooter = "\tPage:  \p of page \P"
    GT.PrintInfo.PreviewCaption = "test Caption"
    If optPortrait = True Then
        GT.PrintInfo.SettingsOrientation = 1
    Else
        GT.PrintInfo.SettingsOrientation = 2
    End If
    GT.PrintInfo.PrintPreview 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProdTemplate.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    Me.GT.Array = XB
    Me.GT.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProdTemplate.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        TOP = 300
        Left = 200
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProdTemplate.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub GT_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    XB.QuickSort XB.LowerBound(1), XB.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex) 'XTYPE_INTEGER
    GT.Refresh

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProdTemplate.GT_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case UCase(XA.Value(ColIndex + 1, 6))
        Case "DATE"
            GetRowType = 4
        Case "CHAR"
            GetRowType = 9
        Case "NUM", "CURR"
            GetRowType = 11
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProdTemplate.GetRowType(ColIndex)", ColIndex
End Function

