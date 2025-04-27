VERSION 5.00
Begin VB.Form frmSecurity 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Security"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSecCode 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   1950
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   900
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Cancel"
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
      Left            =   1500
      Picture         =   "frmSecurity.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1530
      Width           =   1000
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
      Height          =   615
      Left            =   2520
      Picture         =   "frmSecurity.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1530
      Width           =   1000
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your security code"
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
      Height          =   630
      Left            =   120
      TabIndex        =   1
      Top             =   165
      Width           =   4995
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim strSignature As String

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oPC.CurrentSecurityCode = ""
    bCancelled = True
    Me.Hide
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSecurity.cmdCancel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSecurity.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub
Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property
Private Sub cmdOK_Click()
    On Error GoTo errHandler
    Me.Hide
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSecurity.cmdOK_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSecurity.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub
Public Sub component(pMsg As String)
    On Error GoTo errHandler
    lblMsg.Caption = pMsg
    Me.txtSecCode.PasswordChar = "*"
    bCancelled = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSecurity.component(pMsg)", pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSecurity.component(pMsg)", pMsg
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = Chr(13) Then
        Me.Hide
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSecurity.Form_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSecurity.Form_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        TOP = (Screen.Height - Me.Height) / 2
        Left = (Screen.Width - Me.Width) / 2
    End If
    txtSecCode = "aaaaa"
    oPC.CurrentSecurityCode = ""
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSecurity.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSecurity.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtSecCode_Change()
    On Error GoTo errHandler
    oPC.CurrentSecurityCode = txtSecCode  'I think this is now redundant
    strSignature = txtSecCode
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSecurity.txtSecCode_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSecurity.txtSecCode_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtSecCode_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtSecCode")
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSecurity.txtSecCode_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSecurity.txtSecCode_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Public Function GetSignature() As String
    On Error GoTo errHandler
    GetSignature = Trim(strSignature)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSecurity.GetSignature"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSecurity.GetSignature"
End Function
