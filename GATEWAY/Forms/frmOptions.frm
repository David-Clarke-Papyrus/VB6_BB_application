VERSION 5.00
Begin VB.Form frmConfiguration 
   Caption         =   "Configuration "
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00CCC8BB&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   6330
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6090
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00CCC8BB&
      Caption         =   "&OK"
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
      Left            =   5325
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6090
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
      Height          =   1725
      Left            =   195
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4155
      Width           =   7020
   End
   Begin VB.TextBox txtQ 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1740
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1770
      Width           =   7020
   End
   Begin VB.TextBox txtAS400Conn 
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
      Height          =   585
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   510
      Width           =   7020
   End
   Begin VB.Label Label3 
      Caption         =   "AS400 query 1"
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
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   3885
      Width           =   3105
   End
   Begin VB.Label Label2 
      Caption         =   "AS400 query 1"
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
      Height          =   270
      Left            =   225
      TabIndex        =   3
      Top             =   1500
      Width           =   3105
   End
   Begin VB.Label Label1 
      Caption         =   "AS400 connection string"
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
      Height          =   270
      Left            =   225
      TabIndex        =   1
      Top             =   240
      Width           =   3105
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Component()

End Sub

Private Sub cmdCancel_Click()
    oPC.Configuration.CancelEdit
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim strErrors As String
    oPC.Configuration.ApplyEdit strErrors
    Unload Me
End Sub



Private Sub Form_Load()
    txtAS400Conn = oPC.Configuration.AS400COnnectionString
    txtQ = oPC.Configuration.Q
    txtQQ = oPC.Configuration.QQ
End Sub

Private Sub txtAS400Conn_Validate(Cancel As Boolean)
    oPC.Configuration.AS400COnnectionString = txtAS400Conn
End Sub

Private Sub txtQ_Validate(Cancel As Boolean)
    oPC.Configuration.Q = txtQ
End Sub
Private Sub txtQQ_Validate(Cancel As Boolean)
    oPC.Configuration.QQ = txtQQ
End Sub

