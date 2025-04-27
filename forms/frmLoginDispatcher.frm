VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2415
   ClientLeft      =   2835
   ClientTop       =   3195
   ClientWidth     =   6300
   ControlBox      =   0   'False
   Icon            =   "frmLoginDispatcher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLoginDispatcher.frx":000C
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
      Left            =   2925
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
Attribute VB_Name = "frmLogin"
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
    On Error GoTo errHandler
Dim lngResult As Long
    Set oPC = New PapyConn
    Screen.MousePointer = vbHourglass
    
  '  oPC.InitializeSettings
    lngResult = oPC.Connect("")
    Select Case lngResult
    Case 1, 3
        Screen.MousePointer = vbDefault
        MsgBox "Failed to open Database - Papyrus cannot start!" & vbCrLf & "Check the following:" & vbCrLf & "1. The network path to the database exists." & vbCrLf & "2. The ODBC settings (Set via the Windows control panel) point to the correct database." & vbCrLf & "3. The computer that stores the database is switched on.", vbOKOnly, "WARNING!"
        txtUserName.SetFocus
        GoTo EXIT_Handler
    Case 2
        Screen.MousePointer = vbDefault
        If MsgBox("Connection to Bookfind failed" & vbCrLf & "Probably the disk is not in the drive or possibly it is incorrectly installed." & vbCrLf & "Choose OK to to continue without Bookfind and CANCEL to exit.", vbCritical + vbOKCancel, "Bookfind problem") = vbCancel Then
            flgDBConnected = False 'to force closing of program
            Me.Hide
            GoTo EXIT_Handler
        End If
    Case 98
        MsgBox "Your drive mappings are not correctly set. Refer to your support person or use the Papyrus notes under 'Troubleshooting' for help on correcting the problem."
        flgDBConnected = False 'to force closing of program
        Me.Hide
        GoTo EXIT_Handler
    Case 99
        MsgBox "Invalid username or password"
        txtPassword.SetFocus
        GoTo EXIT_Handler
    End Select
    If oPC.Configuration.UnallocatedPT = 0 Then MsgBox "WARNING, no product type has been set as the default, you connot add product records until this is set. " & vbCrLf & "Use the Product types option under Master files on the menu to select a product type as the default.", vbExclamation, "WARNING"
    flgDBConnected = True
 '   Unload frmS
    Screen.MousePointer = vbDefault
    Me.Hide
    SaveSetting "PBKS", "Users", "Username", Me.txtUserName
 
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Login.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
  'Place under splash form
  Left = (Screen.Width - Width) / 2
  Top = (Screen.Height - Height) / 2 + 2000
  flgDBConnected = False
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
