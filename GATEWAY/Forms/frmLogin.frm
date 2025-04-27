VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00DBFAFB&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2415
   ClientLeft      =   2835
   ClientTop       =   3195
   ClientWidth     =   6300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3615
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1260
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "OK"
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
      Height          =   480
      Left            =   2730
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1245
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtUserName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Left            =   2940
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2940
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   870
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading data extractor . . . "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1035
      TabIndex        =   4
      Top             =   1020
      Width           =   4800
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gPassword As String
Dim gUserName As String
Dim arCommandline() As String

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    If MsgBox("Application will not load without database connection!", _
              vbOKCancel, "WARNING!") = vbCancel Then
      Exit Sub
    End If
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLogin.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long
Dim strPos As String

    strPos = "Pos 1s"
    Set oPC = New a_connection
    strPos = "Pos 2"
    Screen.MousePointer = vbHourglass
    strPos = "Pos 3"
    oPC.InitializeSettings
    strPos = "Pos 4"
    oPC.OpenDBSHort
    strPos = "Pos 5"
    Screen.MousePointer = vbDefault
 
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLogin.cmdOK_Click", , EA_NORERAISE, , "strPOS", Array(strPos)
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
  Left = (Screen.Width - Width) / 2
  Top = (Screen.Height - Height) / 2 + 2000
  flgDBConnected = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLogin.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Public Property Get Password() As String
  Password = gPassword
End Property

Public Property Get Username() As String
  Username = gUserName
End Property

Private Sub txtPassword_GotFocus()
    On Error GoTo errHandler
    With txtPassword
    If Len(.Text) > 0 Then
      .SelStart = 0
      .SelLength = Len(.Text)
    End If
  End With

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLogin.txtPassword_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtUserName_GotFocus()
    On Error GoTo errHandler
  With txtUserName
    If Len(.Text) > 0 Then
      .SelStart = 0
      .SelLength = Len(.Text)
    End If
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLogin.txtUserName_GotFocus", , EA_NORERAISE
    HandleError
End Sub
