VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPeriodDialogue_2 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Period selection"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
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
   ScaleHeight     =   3015
   ScaleWidth      =   7290
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Style"
      ForeColor       =   &H8000000D&
      Height          =   1065
      Left            =   150
      TabIndex        =   11
      Top             =   1470
      Width           =   3900
      Begin VB.OptionButton optLDC 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Uses last delivered cost (Ex VAT)"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   255
         TabIndex        =   13
         Top             =   315
         Width           =   3555
      End
      Begin VB.OptionButton optWeighted 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Uses weighted average cost (Ex VAT)"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   255
         TabIndex        =   12
         Top             =   645
         Value           =   -1  'True
         Width           =   3555
      End
   End
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
      Left            =   5955
      Picture         =   "frmPeriodDialogue_2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1950
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   4935
      Picture         =   "frmPeriodDialogue_2.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1950
      Width           =   1000
   End
   Begin VB.CheckBox chkNP 
      BackColor       =   &H00D3D3CB&
      Caption         =   "New page per section"
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   285
      TabIndex        =   7
      Top             =   1065
      Visible         =   0   'False
      Width           =   2175
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
      Left            =   5280
      TabIndex        =   4
      Top             =   180
      Width           =   1665
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
         TabIndex        =   8
         Top             =   870
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
         TabIndex        =   5
         Top             =   525
         Width           =   1065
      End
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   495
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   49348611
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   2385
      TabIndex        =   2
      Top             =   480
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   49348611
      CurrentDate     =   37421
   End
   Begin VB.Label lblAnd 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   1965
      TabIndex        =   3
      Top             =   540
      Width           =   555
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select period between"
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
      Height          =   330
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   3915
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPeriodDialogue_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bPreview As Boolean
Dim dteFrom As Date
Dim dteTo As Date
Dim bCancel As Boolean
Public enPrevPrintCSV As enumReportPresentation
Dim strCostSetting As String

Public Property Get CostSetting() As String
    CostSetting = strCostSetting
End Property

Public Property Get ReportPresentation() As enumReportPresentation
    ReportPresentation = enPrevPrintCSV
End Property
Public Sub Component(pMsg As String, pOneOrTwoDates As Integer, pDefaultDate As Date, Optional pShowchkNP As Boolean)
    lblDescription.Caption = pMsg
    Me.chkNP.Visible = pShowchkNP
    Me.dtpFrom.Value = pDefaultDate
    If pOneOrTwoDates = 1 Then
        dtpTo.Visible = False
        lblAnd.Visible = False
    End If
        
End Sub
Public Sub Component2(Optional pShowchkNP As Boolean)
    Me.chkNP.Visible = pShowchkNP
End Sub

Private Sub cmdClose_Click()
    bCancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    
    strCostSetting = IIf(Me.optWeighted = True, "W", "LPD")
    SaveSetting "PBKS", "Reports", "PeriodDialog_2", strCostSetting
    
    If optPrint Then
        enPrevPrintCSV = enPrintout
    ElseIf optPreview Then
        enPrevPrintCSV = enPreview
    Else
        enPrevPrintCSV = enCSV
    End If
    dteFrom = Me.dtpFrom.Value
    dteTo = Me.dtpTo.Value
    bCancel = False
    Me.Hide
    
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrorIn "frmPeriodDialogue.cmdOK_Click"
End Sub

Private Sub Form_Load()
'    Me.Width = 7000
'    Me.Height = 3400
    
    dtpFrom.Value = DateAdd("w", -1, Date)
    dtpTo.Value = Date
    
    strCostSetting = GetSetting("PBKS", "Reports", "OSApprosCost", "W")
    If strCostSetting = "W" Then
        Me.optWeighted = True
    Else
        Me.optLDC = True
    End If
    
End Sub


Public Property Get FromDate() As Date
    FromDate = dteFrom
End Property
Public Property Get ToDate() As Date
    ToDate = dteTo
End Property
Public Property Get CancelReport() As Boolean
    CancelReport = bCancel
End Property
Public Property Get Preview() As Boolean
    Preview = bPreview
End Property
Public Property Get NP() As Boolean
    NP = (chkNP = 1)
End Property


