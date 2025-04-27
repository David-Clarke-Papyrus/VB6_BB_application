VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2415
   ClientLeft      =   2835
   ClientTop       =   3195
   ClientWidth     =   6300
   ControlBox      =   0   'False
   Icon            =   "frmLoginManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLoginManager.frx":000C
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
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   90
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
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
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
      Left            =   2370
      TabIndex        =   1
      Top             =   45
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
      Left            =   3870
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   1290
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gPassword As String
Dim gUserName As String
Dim bSuccessfulConnection As Boolean
Dim bCancelled As Boolean

Public Property Get Cancelled() As Boolean
50590     Cancelled = bCancelled
End Property
Public Property Get SuccessfulConnection() As Boolean
50600     SuccessfulConnection = bSuccessfulConnection
End Property
Private Sub cmdCancel_Click()
50610     bCancelled = True
50620     Me.Hide
End Sub

Public Sub cmdOK_Click()
50630     On Error GoTo errHandler
      Dim lngResult As Long

50640     errSysHandlerSet
          
50650     bCancelled = False
          
50660     Set oPC = New PapyConn
50670     Screen.MousePointer = vbHourglass
          
50680     If UBound(arCommandLine) > 0 Then
50690         oPC.DatabaseName = arCommandLine(0)
50700     Else
50710         oPC.DatabaseName = ""
50720     End If
          
50730     If UBound(arCommandLine) > 1 Then
50740         oPC.InitializeSettings True
50750     Else
50760         oPC.InitializeSettings
50770     End If
          
          'Just checking
50780     bSuccessfulConnection = (oPC.OpenDB() = 0)
50790     If bSuccessfulConnection Then
50800         oPC.Disconnect
50810         SaveSetting "PBKS", "Users", "Username", Me.txtUserName
50820     Else
50830         MsgBox "Invalid login.", vbOKOnly, "Login status"
50840     End If
          
50850     Screen.MousePointer = vbDefault
          
EXIT_Handler:
50860     Exit Sub
errHandler:
50870     If ErrMustStop Then Debug.Assert False: Resume
50880     ErrPreserve
50890     ErrorIn "Login.cmdOK_Click", , EA_NORERAISE
50900     bCancelled = True
50910     HandleError
End Sub


Private Sub Form_Load()
50920     On Error GoTo errHandler
        'Place under splash form
50930     If Me.WindowState <> 2 Then
50940         Left = (Screen.Width - Width) / 2
50950         TOP = (Screen.Height - Height) / 2 + 2000
50960     End If
50970     Exit Sub
errHandler:
50980     If ErrMustStop Then Debug.Assert False: Resume
50990     ErrorIn "Login.Form_Load"
End Sub

Public Property Get Password() As String
51000   Password = gPassword
End Property

Public Property Get UserName() As String
51010   UserName = gUserName
End Property

Private Sub txtPassword_GotFocus()
51020     With txtPassword
51030     If Len(.Text) > 0 Then
51040       .SelStart = 0
51050       .SelLength = Len(.Text)
51060     End If
51070   End With

End Sub

Private Sub txtUserName_GotFocus()
51080   With txtUserName
51090     If Len(.Text) > 0 Then
51100       .SelStart = 0
51110       .SelLength = Len(.Text)
51120     End If
51130   End With
End Sub
