VERSION 5.00
Begin VB.Form frmFrontDeskSales 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Batch selection"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   6270
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5010
      Picture         =   "frmFrontDeskSales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2190
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   3990
      Picture         =   "frmFrontDeskSales.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2190
      Width           =   1000
   End
   Begin VB.Frame fraPreviewPrint 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   4320
      TabIndex        =   4
      Top             =   240
      Width           =   1665
      Begin VB.OptionButton optPrint 
         BackColor       =   &H00D3D3CB&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   135
         TabIndex        =   7
         Top             =   525
         Width           =   1065
      End
      Begin VB.OptionButton optPreview 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Pre&view"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   360
         Left            =   135
         TabIndex        =   6
         Top             =   165
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton optCSV 
         BackColor       =   &H00D3D3CB&
         Caption         =   "&CSV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   150
         TabIndex        =   5
         Top             =   870
         Width           =   1065
      End
   End
   Begin VB.ComboBox cboTo 
      ForeColor       =   &H8000000D&
      Height          =   345
      ItemData        =   "frmFrontDeskSales.frx":0714
      Left            =   150
      List            =   "frmFrontDeskSales.frx":0716
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1725
      Width           =   3945
   End
   Begin VB.ComboBox cboFrom 
      ForeColor       =   &H8000000D&
      Height          =   345
      ItemData        =   "frmFrontDeskSales.frx":0718
      Left            =   150
      List            =   "frmFrontDeskSales.frx":071A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   3945
   End
   Begin VB.Label Label19 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   630
      TabIndex        =   1
      Top             =   1350
      Width           =   555
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select range"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   4335
   End
End
Attribute VB_Name = "frmFrontDeskSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents oRpts As z_reports
Attribute oRpts.VB_VarHelpID = -1
Dim oTxtList As z_TextList
Dim enPrevPrintCSV As enumReportPresentation
Public Property Get ReportPresentation() As enumReportPresentation
    ReportPresentation = enPrevPrintCSV
End Property
Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashSales.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim blnPrint As Boolean
Dim blnNoRecordsReturned As Boolean
Dim strErrMsg As String
Dim lngTPID As Long
    
    If optPrint Then
        enPrevPrintCSV = enPrint
    ElseIf optPreview Then
        enPrevPrintCSV = enPreview
    Else
        enPrevPrintCSV = enCSV
    End If
        
    Me.MousePointer = vbHourglass
    strErrMsg = oRpts.FrontDeskSales(oTxtList.Key(cboFrom), oTxtList.Key(cboTo), cboFrom, cboTo, blnNoRecordsReturned, enPrevPrintCSV)
    If strErrMsg > "" Then
        MsgBox strErrMsg, vbOKOnly, "ERROR"
    ElseIf blnNoRecordsReturned Then
        MsgBox "No records returned", vbOKOnly, "Papyrus II Reports - Status"
    End If
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashSales.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Me.Width = 7000
    Me.Height = 4800
    
    Set oRpts = New z_reports
    Set oTxtList = New z_TextList
    oTxtList.Load ltCS, ReverseDate(DateAdd("m", -2, Date))
    LoadCombo Me.cboFrom, oTxtList
    LoadCombo Me.cboTo, oTxtList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashSales.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set oRpts = Nothing
    Set oTxtList = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashSales.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

