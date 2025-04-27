VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTRDate 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Invoice date"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   405
      Left            =   420
      TabIndex        =   3
      Top             =   690
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   -2147483635
      CalendarTitleForeColor=   -2147483635
      Format          =   221839360
      CurrentDate     =   38967
      MaxDate         =   73415
      MinDate         =   36526
   End
   Begin VB.CommandButton cmdCurrentdate 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Today"
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
      Height          =   375
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   690
      Width           =   690
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1530
      Picture         =   "frmTRDate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1470
      Width           =   1000
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Select document date"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   360
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   4185
   End
End
Attribute VB_Name = "frmTRDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCurrentdate_Click()
    On Error GoTo errHandler
    DTP1 = Date
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTRDate.cmdCurrentdate_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTRDate.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub
Public Sub component(pDate As Date)
    On Error GoTo errHandler
    Me.DTP1 = pDate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTRDate.component(pDate)", pDate
End Sub



Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        TOP = (Screen.Height - Me.Height) / 2
        Left = (Screen.Width - Me.Width) / 2
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTRDate.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Public Property Get InvoiceDate() As Date
    InvoiceDate = DTP1
End Property
