VERSION 5.00
Begin VB.Form frmRefunds 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Refunds management"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   570
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmRefunds.frx":0000
      Top             =   1395
      Width           =   3120
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
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
      Height          =   420
      Left            =   5235
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3210
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton cmd 
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
      Top             =   2115
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
      Height          =   570
      Left            =   330
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   375
      Width           =   5460
   End
End
Attribute VB_Name = "frmRefunds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean

Private Sub cmd_Click()
    Me.Hide
End Sub

Public Sub Component(pstr As String, pbuttonmsg As String, pValue As String)
    Me.txtMessage = pstr
    Me.cmd.Caption = pbuttonmsg
    Me.txtValue = pValue
    bCancelled = False
End Sub

Private Sub cmdCancel_Click()
    bCancelled = True
End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property
