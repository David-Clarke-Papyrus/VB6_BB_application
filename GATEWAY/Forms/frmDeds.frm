VERSION 5.00
Begin VB.Form frmConfiguration 
   Caption         =   "Configuration"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   7020
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00CCC8BB&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6180
      Width           =   915
   End
   Begin VB.TextBox txtQQ 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1260
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   4845
      Width           =   6360
   End
   Begin VB.TextBox txtQ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2715
      Left            =   330
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1755
      Width           =   6360
   End
   Begin VB.CommandButton cmdCancek 
      BackColor       =   &H00CCC8BB&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5805
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6180
      Width           =   915
   End
   Begin VB.TextBox txtDSN 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   660
      Left            =   270
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   555
      Width           =   6360
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Connection string to AS400"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   255
      TabIndex        =   6
      Top             =   4590
      Width           =   2490
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Connection string to AS400"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   285
      TabIndex        =   4
      Top             =   1500
      Width           =   2490
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Connection string to AS400"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   225
      TabIndex        =   1
      Top             =   300
      Width           =   2490
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cmdCancek_Click()
    oPC.Configuration.CancelEdit
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim strError As String
    oPC.Configuration.ApplyEdit strError
    Unload Me
End Sub

Private Sub Form_Load()
    txtDSN = oPC.Configuration.AS400COnnectionString
    txtQ = oPC.Configuration.Q
    txtQQ = oPC.Configuration.QQ
End Sub

Private Sub txtQ_Validate(Cancel As Boolean)
    oPC.Configuration.Q = FNS(txtQ)
End Sub

Private Sub txtQQQ_Validate(Cancel As Boolean)
    oPC.Configuration.QQ = FNS(txtQQ)
End Sub

Private Sub txtDSN_Validate(Cancel As Boolean)
    oPC.Configuration.AS400COnnectionString = FNS(txtDSN)
End Sub

