VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrintPreviousActions 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Print previous actions"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00D7D1BF&
      Caption         =   "Print"
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
      Left            =   1680
      Picture         =   "frmPrintPreviousActions.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2445
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker DT1 
      Height          =   405
      Left            =   1380
      TabIndex        =   1
      Top             =   1530
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   -2147483635
      CalendarTitleForeColor=   -2147483635
      Format          =   221839361
      CurrentDate     =   37744
      MaxDate         =   73415
      MinDate         =   36892
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "You can print a history of actions taken on overdue purchase or customer orders. Select the date since which you wish to report."
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
      Height          =   960
      Left            =   435
      TabIndex        =   0
      Top             =   225
      Width           =   3735
   End
End
Attribute VB_Name = "frmPrintPreviousActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ar As arCustReport2
Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
    Set ar = New arCustReport2
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open "SELECT * FROM vCOLActions", oPC.COShort, adOpenForwardOnly, adLockOptimistic
    Unload Me
    ar.component rs
    ar.Show
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintPreviousActions.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
