VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStatementControl 
   Caption         =   "Statement control"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6645
   Icon            =   "frmStatementControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCLose 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Close"
      Height          =   375
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2535
      Width           =   900
   End
   Begin VB.CommandButton cmdEmail 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Email"
      Height          =   375
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2535
      Width           =   900
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Preview"
      Height          =   375
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2535
      Width           =   900
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Print"
      Height          =   375
      Left            =   1215
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2535
      Width           =   900
   End
   Begin VB.CheckBox chkNonZeroOnly 
      Alignment       =   1  'Right Justify
      Caption         =   "Show statements with a non-zero balance only"
      ForeColor       =   &H8000000D&
      Height          =   390
      Left            =   345
      TabIndex        =   7
      Top             =   2055
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date controls"
      ForeColor       =   &H8000000D&
      Height          =   1725
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   6165
      Begin MSComCtl2.DTPicker dtpStatementDate 
         Height          =   345
         Left            =   1380
         TabIndex        =   1
         Top             =   345
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   609
         _Version        =   393216
         Format          =   221839361
         CurrentDate     =   39980
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   345
         Left            =   1935
         TabIndex        =   3
         Top             =   1020
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   609
         _Version        =   393216
         Format          =   221839361
         CurrentDate     =   39980
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   345
         Left            =   3855
         TabIndex        =   5
         Top             =   1020
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   609
         _Version        =   393216
         Format          =   221839361
         CurrentDate     =   39980
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "to"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Top             =   1065
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Show transactions from"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   75
         TabIndex        =   4
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Statement date"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   75
         TabIndex        =   2
         Top             =   390
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmStatementControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bPreview As Boolean
Dim bPrint As Boolean
Dim bEmail As Boolean
Dim dteStatementDate As Date
Dim dteFrom As Date
Dim dteTo As Date

Public Sub component(pSingleCustomer As Boolean)
    On Error GoTo errHandler
    If pSingleCustomer Then
        chkNonZeroOnly.Value = 0
        chkNonZeroOnly.Enabled = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStatementControl.component(pSingleCustomer)", pSingleCustomer
End Sub

Public Property Get NonZeroOnly() As Boolean
    On Error GoTo errHandler
    NonZeroOnly = (chkNonZeroOnly.Value = 1)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStatementControl.NonZeroOnly"
End Property

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    bPreview = False
    bPrint = False
    bEmail = False
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStatementControl.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEmail_Click()
    On Error GoTo errHandler
    bPreview = False
    bPrint = False
    bEmail = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStatementControl.cmdEmail_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo errHandler
    bPreview = True
    bPrint = False
    bEmail = False
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStatementControl.cmdPreview_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    bPrint = True
    bPreview = False
    bEmail = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStatementControl.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Public Property Get StatementDate() As Date
    StatementDate = dteStatementDate
End Property
Public Property Get FromDate() As Date
    FromDate = dteFrom
End Property
Public Property Get ToDate() As Date
    ToDate = dteTo
End Property

