VERSION 5.00
Begin VB.Form frmConfirmStrip 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Confirm"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboOp 
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
      Left            =   630
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1350
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&No"
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
      Left            =   2655
      Picture         =   "frmConfirmStrip.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2190
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Yes"
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
      Height          =   615
      Left            =   1545
      Picture         =   "frmConfirmStrip.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2190
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   630
      TabIndex        =   4
      Top             =   1050
      Width           =   1725
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "You want to permanently remove 'A ', 'The ', 'n ' and 'An ' from the start of all titles?"""
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
      Height          =   945
      Left            =   270
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmConfirmStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iResult As Integer
Dim lngOperatorID As Long
Dim tlOperators As New z_TextList


Private Sub Command1_Click()
    iResult = 1
    Me.Hide
End Sub

Private Sub Command2_Click()
    iResult = 0
    Me.Hide
End Sub
Public Function GetResult() As Integer
    GetResult = iResult
End Function
Friend Property Get OperatorID() As Long
    OperatorID = lngOperatorID
End Property

Private Sub cboOp_Click()
    lngOperatorID = tlOperators.Key(cboOp.Text)
End Sub

Private Sub Form_Load()
    tlOperators.Load ltStaff
    LoadCombo Me.cboOp, tlOperators
    lngOperatorID = tlOperators.Key(cboOp.Text)
End Sub

