VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCommissionDialogue 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Period selection"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCommissionDialogue.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   5835
   StartUpPosition =   1  'CenterOwner
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
      Left            =   4620
      Picture         =   "frmCommissionDialogue.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3060
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   3600
      Picture         =   "frmCommissionDialogue.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3060
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
      Left            =   270
      TabIndex        =   6
      Top             =   2415
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   870
         Width           =   1065
      End
   End
   Begin VB.ComboBox cboSalesRep 
      Height          =   345
      Left            =   270
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   525
      Width           =   2910
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   225
      TabIndex        =   1
      Top             =   1680
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   171704323
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   2895
      TabIndex        =   2
      Top             =   1680
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   171704323
      CurrentDate     =   37421
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales rep"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   285
      TabIndex        =   5
      Top             =   225
      Width           =   975
   End
   Begin VB.Label lblAnd 
      BackColor       =   &H00D3D3CB&
      Caption         =   "and"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   2430
      TabIndex        =   3
      Top             =   1725
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
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   1215
      Width           =   5460
   End
End
Attribute VB_Name = "frmCommissionDialogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bPreview As Boolean
Dim dteFrom As Date
Dim dteTo As Date
Dim mStaffID As Long
Dim bCancel As Boolean
Dim tlRep As z_TextList

Dim enPrevPrintCSV As enumReportPresentation
Public Property Get ReportPresentation() As enumReportPresentation
    ReportPresentation = enPrevPrintCSV
End Property
Public Sub Component(pMsg As String, pOneOrTwoDates As Integer, pDefaultDate As Date)
    lblDescription.Caption = pMsg
    Me.dtpFrom.Value = pDefaultDate
    If pOneOrTwoDates = 1 Then
        dtpTo.Visible = False
        lblAnd.Visible = False
    End If
        
End Sub
Public Property Get STAFFID() As Long
    STAFFID = mStaffID
End Property


Private Sub cboSalesRep_Click()
    mStaffID = tlRep.Key(cboSalesRep)
End Sub

Private Sub cmdClose_Click()
    bCancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    
    If optPrint Then
        enPrevPrintCSV = enPrint
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
    Set tlRep = New z_TextList
    Me.Width = 6000
    Me.Height = 4200
    
    dtpFrom.Value = DateAdd("w", -1, Date)
    dtpTo.Value = Date
    
    tlRep.Load ltSalesRep, , "<NONE>"
    LoadCombo cboSalesRep, tlRep

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

