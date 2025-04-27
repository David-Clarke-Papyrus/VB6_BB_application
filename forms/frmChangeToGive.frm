VERSION 5.00
Begin VB.Form frmChangeToGive 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Voucher too far exceeds value of sale."
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCash 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Issue Cash and Gift Vouchers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1680
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      UseMaskColor    =   -1  'True
      Width           =   3105
   End
   Begin VB.CommandButton cmdCN 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&YES: Issue credit note"
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
      Height          =   585
      Left            =   2085
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2475
      UseMaskColor    =   -1  'True
      Width           =   2145
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
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
      Height          =   1110
      Left            =   330
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   375
      Width           =   5460
   End
End
Attribute VB_Name = "frmChangeToGive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bIssueCreditNote As Boolean

Private Sub cmdCash_Click()
    bIssueCreditNote = False
    Me.Hide
End Sub

Private Sub cmdCN_Click()
    bIssueCreditNote = True
    Me.Hide
End Sub
Public Property Get IssueChangeAsCreditNote() As Boolean
    IssueChangeAsCreditNote = bIssueCreditNote
End Property

Public Sub component(pstr As String)
    Me.txtMessage = pstr
      If oPC.UsageContext = "BB" Then
        cmdCN.Visible = False
      End If
End Sub
