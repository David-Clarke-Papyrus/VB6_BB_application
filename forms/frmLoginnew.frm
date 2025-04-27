VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2415
   ClientLeft      =   2835
   ClientTop       =   3195
   ClientWidth     =   6300
   ControlBox      =   0   'False
   Icon            =   "frmLoginnew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLoginnew.frx":000C
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
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gPassword As String
Dim gUserName As String

Private Sub cmdCancel_Click()
    If MsgBox("Application will not load without database connection!", _
              vbOKCancel, "WARNING!") = vbCancel Then
      Exit Sub
    End If
    Me.Hide
End Sub

Public Sub cmdOK_Click()
On Error GoTo ERR_Handler
Dim lngResult As Long
'=========
    Set oPC = New PapyConn
    If UBound(arCommandLine) > 0 Then
        oPC.DatabaseName = arCommandLine(0)
    Else
        oPC.DatabaseName = ""
    End If
    
    oPC.InitializeSettings
    Screen.MousePointer = vbHourglass
    
    lngResult = oPC.OpenDB()
    
    flgDBConnected = True
    Screen.MousePointer = vbDefault
    Me.Hide
    SaveSetting "PBKS", "Users", "Username", Me.txtUserName
 
EXIT_Handler:
    Exit Sub
ERR_Handler:
    If Err = vbObjectError + 333 Then
        MsgBox "Problem with Bookfind"
        Resume
    ElseIf Err = vbObjectError + 800 Then
        MsgBox "You are trying to Connect twice"
        flgDBConnected = False 'to force closing of program
        Me.Hide
    End If
    MsgBox Error & " In Login form :cmdOK"
    GoTo EXIT_Handler
    Resume
Resume
End Sub

Private Sub Form_Activate()
'  Me.txtPassword.SetFocus
End Sub

Private Sub Form_Load()
  'txtUserName = "admin"
  'txtPassword = "sru"
  'Place under splash form
  Left = (Screen.Width - Width) / 2
  Top = (Screen.Height - Height) / 2 + 2000
  flgDBConnected = False
  Me.txtUserName = GetSetting("PBKS", "Users", "Username", "sa")
 ' Me.txtPassword.SetFocus
End Sub

Public Property Get Password() As String
  Password = gPassword
End Property

Public Property Get UserName() As String
  UserName = gUserName
End Property

Private Sub txtPassword_GotFocus()
    With txtPassword
    If Len(.Text) > 0 Then
      .SelStart = 0
      .SelLength = Len(.Text)
    End If
  End With

End Sub

Private Sub txtUserName_GotFocus()
  With txtUserName
    If Len(.Text) > 0 Then
      .SelStart = 0
      .SelLength = Len(.Text)
    End If
  End With
End Sub
