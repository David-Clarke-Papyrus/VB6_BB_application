VERSION 5.00
Begin VB.Form frmCloseSession 
   Caption         =   "Close day session"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   2625
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCloseTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   555
      TabIndex        =   1
      Top             =   495
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
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
      Height          =   675
      Left            =   720
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmCloseSession.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1470
      Width           =   1110
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Must be date and time e.g. 1/3/2009 14:27"
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   120
      TabIndex        =   3
      Top             =   870
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00D3D3CB&
      BackStyle       =   0  'Transparent
      Caption         =   "End date and time"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   255
      TabIndex        =   2
      Top             =   240
      Width           =   2100
   End
End
Attribute VB_Name = "frmCloseSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCancelled As Boolean

Public Sub component(pQty As Long)
    On Error GoTo errHandler
    txtCloseTime = pQty
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCloseSession.component(pQty)", pQty
End Sub

Public Property Get CloseTime() As String
    CloseTime = txtCloseTime
End Property
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancelled = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCloseSession.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCloseSession.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtCloseTime_Change()
    On Error GoTo errHandler
    lblMsg.Visible = Not (IsDate(txtCloseTime))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCloseSession.txtCloseTime_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCloseTime_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not (IsDate(txtCloseTime))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCloseSession.txtCloseTime_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
