VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPeriodSwitch 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Month end controls"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Periods"
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
      Height          =   2040
      Left            =   105
      TabIndex        =   2
      Top             =   90
      Width           =   5640
      Begin VB.TextBox txtPeriods 
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1275
         Left            =   165
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   705
         Width           =   2355
      End
      Begin VB.CommandButton OKButton 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Set"
         CausesValidation=   0   'False
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
         Left            =   4425
         Picture         =   "frmPeriodSwitch.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   195
         Width           =   1000
      End
      Begin VB.CommandButton cmdAdvanced 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Advanced"
         CausesValidation=   0   'False
         Default         =   -1  'True
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
         Left            =   3060
         Picture         =   "frmPeriodSwitch.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1245
         Width           =   2430
      End
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   390
         Left            =   2715
         TabIndex        =   5
         Top             =   270
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   221839361
         UpDown          =   -1  'True
         CurrentDate     =   36526
         MinDate         =   -73046
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of start of new period"
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
         Height          =   390
         Left            =   15
         TabIndex        =   6
         Top             =   300
         Width           =   2700
      End
   End
   Begin VB.CommandButton cmdME 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Month end procedure"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   150
      Picture         =   "frmPeriodSwitch.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2220
      Width           =   1995
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
      Left            =   4605
      Picture         =   "frmPeriodSwitch.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2655
      Width           =   1110
   End
End
Attribute VB_Name = "frmPeriodSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim ob As New z_Batch

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cmdAdvanced_Click()
Dim frm As New frmPeriods
Dim oSQL As New z_SQL
    frm.Show vbModal
    Me.txtPeriods.text = oSQL.GetAccountingPeriods
    
End Sub

Private Sub cmdME_Click()
Dim oSQL As New z_SQL
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    oSQL.RunProc "MonthEnd", Array(), "Running month-end"
    Screen.MousePointer = vbDefault
    MsgBox "Month end procedure complete", vbInformation, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPeriodSwitch.cmdME_Click"
End Sub

Private Sub Form_Load()
Dim oSQL As New z_SQL
    Me.DTP1 = Date
    Me.txtPeriods.text = oSQL.GetAccountingPeriods
End Sub

Private Sub OKButton_Click()
    On Error GoTo errHandler
Dim strSQL As String
Dim strDescription As String
Dim oSQL As New z_SQL

    oSQL.RunProc "SwitchPeriod", Array(DTP1), "Switching periods"
    Me.txtPeriods.text = oSQL.GetAccountingPeriods
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPeriodSwitch.OKButton_Click"
End Sub
