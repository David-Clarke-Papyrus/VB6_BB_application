VERSION 5.00
Begin VB.Form Loginold 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4845
   ClientLeft      =   2835
   ClientTop       =   3195
   ClientWidth     =   5100
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":000C
   ScaleHeight     =   323
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   360
      Left            =   1710
      TabIndex        =   0
      Top             =   2730
      Width           =   1725
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3510
      Width           =   855
   End
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
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3525
      Width           =   855
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1710
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3120
      Width           =   1725
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F2FFF2&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
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
      Height          =   270
      Index           =   1
      Left            =   1935
      TabIndex        =   1
      Top             =   2655
      Width           =   1080
   End
End
Attribute VB_Name = "Loginold"
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

Private Sub cmdOK_Click()
On Error GoTo ERR_Handler
Dim lngResult As Long
    Set oPC = New PapyConn
    Screen.MousePointer = vbHourglass
    oError.SetError "Diagnostic", CStr(Err) & Error, Now(), "UI:Login:Before connect", "", ""
    
    oPC.UserName = Trim$(txtUserName)
    oPC.Password = Trim$(txtPassword)
 '  oPC.Servername = Trim$(txtServer)
  '  oPC.Database = "Papyrus"
    
    lngResult = oPC.Connect()
    oError.SetError "Diagnostic", CStr(Err) & Error, Now(), "UI:Login:after connect", "", ""
    Select Case lngResult
    Case 1, 3
        Screen.MousePointer = vbDefault
        MsgBox "Failed to open Database - Papyrus cannot start!" & vbCrLf & "Check the following:" & vbCrLf & "1. The network path to the database exists." & vbCrLf & "2. The ODBC settings (Set via the Windows control panel) point to the correct database." & vbCrLf & "3. The computer that stores the database is switched on.", vbOKOnly, "WARNING!"
   '     Unload frmS
        txtUserName.SetFocus
        GoTo EXIT_Handler
    Case 2
        Screen.MousePointer = vbDefault
        If MsgBox("Connection to Bookfind failed" & vbCrLf & "Probably the disk is not in the drive or possibly it is incorrectly installed." & vbCrLf & "Choose OK to to continue without Bookfind and CANCEL to exit.", vbCritical + vbOKCancel, "Bookfind problem") = vbCancel Then
            flgDBConnected = False 'to force closing of program
  '          Unload frmS
            Me.Hide
            GoTo EXIT_Handler
        End If
   '     frmS.Refresh
    Case 98
        MsgBox "Your drive mappings are not correctly set. Refer to your support person or use the Papyrus notes under 'Troubleshooting' for help on correcting the problem."
        flgDBConnected = False 'to force closing of program
     '   Unload frmS
        Me.Hide
        GoTo EXIT_Handler
    Case 99
        MsgBox "Invalid username or password"
    '    Me.Show
    '    Unload frmS
        txtPassword.SetFocus
        GoTo EXIT_Handler
    End Select
    flgDBConnected = True
 '   Unload frmS
    Screen.MousePointer = vbDefault
    Me.Hide
    SaveSetting App.Title, "Users", "Username", Me.txtUserName
 
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
  Me.txtPassword.SetFocus
End Sub

Private Sub Form_Load()
  'txtUserName = "admin"
  'txtPassword = "sru"
  'Place under splash form
  Left = (Screen.Width - Width) / 2
  Top = (Screen.Height - Height) / 2 + 2000
  flgDBConnected = False
  Me.txtUserName = GetSetting(App.Title, "Users", "Username", "myadmin")
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
