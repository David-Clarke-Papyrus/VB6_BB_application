VERSION 5.00
Begin VB.Form frmStartup 
   Caption         =   "Startup"
   ClientHeight    =   3150
   ClientLeft      =   5115
   ClientTop       =   3180
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start scanner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2790
      TabIndex        =   4
      Top             =   1950
      Width           =   1935
   End
   Begin VB.ComboBox cboSP 
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
      Height          =   360
      Left            =   2250
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   690
      Width           =   3075
   End
   Begin VB.TextBox txtNote 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   2250
      TabIndex        =   0
      Top             =   1125
      Width           =   3090
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Salesperson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   930
      TabIndex        =   3
      Top             =   735
      Width           =   1290
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   915
      TabIndex        =   2
      Top             =   1155
      Width           =   1290
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objCS As a_CS

Private Sub cmdStart_Click()
 '   frmMain.SetOperator (cboSP.Text)
 '   frmMain.SetNote (txtNote)
    Me.Hide
End Sub

Private Sub Form_Load()
    LoadCombo cboSP, objCS.operators
End Sub
Public Sub Component(oCS As a_CS)
    Set objCS = oCS
End Sub
