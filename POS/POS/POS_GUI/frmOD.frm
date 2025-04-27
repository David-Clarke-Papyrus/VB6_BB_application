VERSION 5.00
Begin VB.Form frmOD 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Opening drawer"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   4425
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00DACDCD&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2310
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00DACDCD&
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1770
      Width           =   1410
   End
   Begin VB.TextBox txtReason 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   990
      IMEMode         =   3  'DISABLE
      Left            =   135
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   645
      Width           =   3750
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Esc to quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   210
      Left            =   840
      TabIndex        =   4
      Top             =   2490
      Width           =   1095
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason (min 5 characters)"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   330
      Left            =   435
      TabIndex        =   2
      Top             =   270
      Width           =   3195
   End
End
Attribute VB_Name = "frmOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strReason As String
Dim bCancel As Boolean

Private Sub cmdClose_Click()
    bCancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    bCancel = False
    Me.Hide
End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = bCancel
End Property
Public Property Get Reason() As String
    Reason = Replace(strReason, vbTab, "")
End Property

Private Sub txtReason_Change()
    cmdOK.Enabled = Len(Trim(txtReason)) > 5
    strReason = Trim(txtReason)
End Sub

Public Sub Component(Optional pTitle As String)
    Me.Caption = "Opening drawer"
    If pTitle > "" Then
        Me.Caption = pTitle
    End If
End Sub
