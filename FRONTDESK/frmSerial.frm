VERSION 5.00
Begin VB.Form frmSerial 
   Caption         =   "Serial port"
   ClientHeight    =   6030
   ClientLeft      =   4860
   ClientTop       =   2910
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   4260
   Begin VB.TextBox txtTimerInterval 
      Height          =   300
      Left            =   2595
      TabIndex        =   9
      Top             =   4980
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Show scanned value"
      Height          =   495
      Left            =   660
      TabIndex        =   8
      Top             =   2700
      Width           =   2265
   End
   Begin VB.CheckBox chkCRLF 
      Caption         =   "Scanner sends CR/LF"
      Height          =   495
      Left            =   660
      TabIndex        =   7
      Top             =   2220
      Width           =   2265
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   570
      Left            =   225
      TabIndex        =   5
      Top             =   4920
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select COM port"
      Height          =   1230
      Left            =   225
      TabIndex        =   0
      Top             =   390
      Width           =   3240
      Begin VB.OptionButton opt4 
         Caption         =   "COM 4"
         Height          =   195
         Left            =   1575
         TabIndex        =   4
         Top             =   645
         Width           =   885
      End
      Begin VB.OptionButton opt3 
         Caption         =   "COM 3"
         Height          =   255
         Left            =   1575
         TabIndex        =   3
         Top             =   300
         Width           =   930
      End
      Begin VB.OptionButton opt2 
         Caption         =   "COM 2"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   645
         Width           =   885
      End
      Begin VB.OptionButton opt1 
         Caption         =   "COM 1"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Timer interval is set in hte PBKS.INI file"
      Height          =   240
      Left            =   150
      TabIndex        =   10
      Top             =   3285
      Width           =   3795
   End
   Begin VB.Label Label1 
      Caption         =   "baud=9600  parity=N  data=7  stop=1"
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
      Left            =   210
      TabIndex        =   6
      Top             =   4080
      Width           =   3435
   End
End
Attribute VB_Name = "frmSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkCRLF_Click()
    bSendsCRLF = (chkCRLF = 1)
End Sub

Private Sub chkShow_Click()
    DebugMode = (chkShow = 1)
End Sub

Private Sub cmdOK_Click()
    SaveSetting App.Title, "Settings", "CRLF", IIf(bSendsCRLF, "TRUE", "FALSE")
    SaveSetting App.Title, "Settings", "Debugmode", IIf(DebugMode, "TRUE", "FALSE")
    SaveSetting App.Title, "Settings", "COMPort", strCom
    SaveSetting App.Title, "Settings", "Show", DebugMode
    If IsNumeric(txtTimerInterval) Then
        iTimer1Interval = CInt(txtTimerInterval)
        SaveSetting App.Title, "Settings", "Timer1Interval", iTimer1Interval
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    strCom = GetSetting(App.Title, "Settings", "COMPort", "COM2")
    Select Case strCom
    Case "COM1:"
        Me.opt1.Value = True
    Case "COM2:"
        Me.opt2.Value = True
    Case "COM3:"
        Me.opt3.Value = True
    Case "COM4:"
        Me.opt4.Value = True
    End Select
    bSendsCRLF = GetSetting(App.Title, "Settings", "CRLF", False)
    chkCRLF = IIf(bSendsCRLF = "TRUE", 1, 0)
    DebugMode = GetSetting(App.Title, "Settings", "Debugmode", False)
    chkShow = IIf(DebugMode = "TRUE", 1, 0)
End Sub

Private Sub opt1_Click()
    strCom = "COM1:"
End Sub

Private Sub opt2_Click()
    strCom = "COM2:"
End Sub

Private Sub opt3_Click()
    strCom = "COM3:"
End Sub

Private Sub opt4_Click()
    strCom = "COM4:"
End Sub
