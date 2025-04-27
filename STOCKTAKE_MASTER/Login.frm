VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2340
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   156
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   225
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
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
      Height          =   450
      Left            =   870
      TabIndex        =   4
      Top             =   1590
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   2490
      TabIndex        =   5
      Top             =   1590
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   615
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Index           =   1
      Left            =   135
      TabIndex        =   2
      Top             =   630
      Width           =   1080
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
    If MsgBox("Application will NOT load without Database connection!", _
              vbOKCancel, "WARNING!") = vbCancel Then
      Exit Sub
    End If
    Me.Hide
End Sub


Private Sub cmdOK_Click()
On Error GoTo ERR_Handler
Dim lngResult As Long
    Set gPapyConn = New PapyConn
    Screen.MousePointer = vbHourglass
    
    gPapyConn.Username = Trim$(txtUserName)
    gPapyConn.Password = Trim$(txtPassword)
    gPapyConn.Database = "Papyrus"
    lngResult = gPapyConn.Connect()
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
    Case 99
        MsgBox "Invalid username or password"
        txtUserName.SetFocus
        GoTo EXIT_Handler
    End Select
    flgDBConnected = True
    Set gError = New a_Error
    
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
        GoTo EXIT_Handler
    End If
    MsgBox Error
   ' Resume Next
End Sub

Private Sub Form_Load()
  'txtUserName = "admin"
  'txtPassword = "sru"
  'Place under splash form
  Left = (Screen.Width - Width) / 2
  Top = (Screen.Height - Height) / 2 + 2000
  flgDBConnected = False
  Me.txtUserName = GetSetting(App.Title, "Users", "Username", "rwuser")
End Sub

Public Property Get Password() As String
  Password = gPassword
End Property

Public Property Get Username() As String
  Username = gUserName
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
