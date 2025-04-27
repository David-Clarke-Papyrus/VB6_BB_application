VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmPOSInvoice 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Payment details"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tCash 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000005&
      Height          =   480
      Left            =   2210
      TabIndex        =   12
      Top             =   5010
      Width           =   2115
   End
   Begin VB.TextBox tQ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000005&
      Height          =   480
      Left            =   2210
      TabIndex        =   9
      Top             =   2040
      Width           =   2115
   End
   Begin VB.TextBox tCN 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000005&
      Height          =   480
      Left            =   2210
      TabIndex        =   6
      Top             =   4365
      Width           =   2115
   End
   Begin VB.TextBox tCC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000005&
      Height          =   480
      Left            =   2210
      TabIndex        =   4
      Top             =   1305
      Width           =   2115
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00D7D1BF&
      Caption         =   "&OK"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   4740
      Picture         =   "frmPOSInvoice.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4350
      Width           =   2025
   End
   Begin TrueOleDBGrid60.TDBGrid GV 
      Height          =   1380
      Left            =   2220
      OleObjectBlob   =   "frmPOSInvoice.frx":038A
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2760
      Width           =   4515
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00D3D3CB&
      Caption         =   "R0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   4155
      TabIndex        =   15
      Top             =   735
      Width           =   2670
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   375
      TabIndex        =   14
      Top             =   690
      Width           =   3195
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00D3D3CB&
      Caption         =   "R0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   4155
      TabIndex        =   13
      Top             =   240
      Width           =   2670
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Cash"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   180
      TabIndex        =   11
      Top             =   5040
      Width           =   1605
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Credit note"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   180
      TabIndex        =   8
      Top             =   4395
      Width           =   1605
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Voucher"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   180
      TabIndex        =   7
      Top             =   2670
      Width           =   1605
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Cheque"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   180
      TabIndex        =   5
      Top             =   1995
      Width           =   1605
   End
   Begin VB.Label lblError 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1260
      Left            =   4515
      TabIndex        =   3
      Top             =   1275
      Width           =   3225
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Credit card"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   180
      TabIndex        =   1
      Top             =   1335
      Width           =   1605
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Amount to pay"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   915
      TabIndex        =   0
      Top             =   255
      Width           =   2670
   End
End
Attribute VB_Name = "frmPOSInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub Component(pTRID As Long)

End Sub

Private Sub cmdOK_Click()
    'ID the operator
End Sub

Private Sub Form_Load()
    InitForm
End Sub
