VERSION 5.00
Begin VB.Form frmAlert 
   Caption         =   "Alert"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C4BCA4&
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
      Left            =   3795
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmAlertLoad.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.TextBox txtMsg 
      Height          =   1350
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmAlertLoad.frx":038A
      Top             =   1350
      Width           =   4485
   End
   Begin VB.Frame frTo 
      Caption         =   "To"
      ForeColor       =   &H8000000D&
      Height          =   630
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   4500
      Begin VB.OptionButton optSM 
         Caption         =   "Staff member"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   195
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.OptionButton optCust 
         Caption         =   "Customer"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   435
         TabIndex        =   1
         Top             =   195
         Width           =   1110
      End
   End
   Begin VB.Label lblDestination 
      Height          =   315
      Left            =   135
      TabIndex        =   3
      Top             =   840
      Width           =   4470
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTPID As Long
Dim strAcno As String
Public Sub component(Optional TPID As Long, Optional TPNAME As String, Optional TPACNO As String)
    On Error GoTo errHandler
    lblDestination.Caption = TPNAME & " " & TPACNO
    lngTPID = TPID
    strAcno = TPACNO
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAlert.component(TPID,TPNAME,TPACNO)", Array(TPID, TPNAME, TPACNO)
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errHandler
Dim oSQL As New z_SQL

    oSQL.LoadAlert "C", FNS(txtMsg), strAcno
    MsgBox "Message sent", vbOKOnly, "Alert"
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAlert.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub
