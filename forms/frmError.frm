VERSION 5.00
Begin VB.Form frmError 
   BackColor       =   &H00D1D1D1&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExpand 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2895
      Width           =   270
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C8B9B3&
      Caption         =   "OK"
      Height          =   480
      Left            =   1590
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2685
      Width           =   1200
   End
   Begin VB.TextBox txtMsg 
      Height          =   3435
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmError.frx":0000
      Top             =   3945
      Width           =   9975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Click OK to continue. The application may be forced to close."
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
      Height          =   630
      Left            =   330
      TabIndex        =   2
      Top             =   1905
      Width           =   3675
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmError.frx":0006
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
      Height          =   1575
      Left            =   345
      TabIndex        =   1
      Top             =   315
      Width           =   3675
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bExpanded As Boolean
Private Sub cmdExpand_Click()
    bExpanded = Not bExpanded
    If bExpanded Then
        txtMsg.Visible = True
        Me.Width = 10485
        Me.Height = 8175
    Else
        txtMsg.Visible = False
        Me.Width = 4515
        Me.Height = 3915
    End If
    Me.Refresh
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Public Sub SettxtMsg(pMsg As String)
    Me.txtMsg = pMsg
End Sub

Private Sub Form_Load()
        Me.Width = 4515
        Me.Height = 3915
End Sub
