VERSION 5.00
Begin VB.Form frmCancelPOLs 
   Caption         =   "Cancel P.O.L.s"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3930
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Reason"
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
      Height          =   1320
      Left            =   780
      TabIndex        =   2
      Top             =   420
      Width           =   2325
      Begin VB.PictureBox Picture 
         Height          =   900
         Left            =   90
         ScaleHeight     =   840
         ScaleWidth      =   2115
         TabIndex        =   3
         Top             =   285
         Width           =   2175
         Begin VB.OptionButton optOOP 
            Caption         =   "Out of Print"
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
            Height          =   300
            Left            =   225
            TabIndex        =   5
            Top             =   120
            Value           =   -1  'True
            Width           =   1605
         End
         Begin VB.OptionButton optREPrint 
            Caption         =   "Reprinting"
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
            Height          =   300
            Left            =   225
            TabIndex        =   4
            Top             =   480
            Width           =   1605
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&No action"
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
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1500
   End
   Begin VB.CommandButton cmdRediarize 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Cancel P.O.L.s"
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
      Left            =   465
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   1500
   End
End
Attribute VB_Name = "frmCancelPOLs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iNumber As Integer
Dim strUnit As String
Dim bCancel As Boolean

Private Sub cmdPrint_Click()
Me.Hide
End Sub

Private Sub cmdCancel_Click()
    bCancel = True
    Me.Hide
End Sub

Private Sub cmdRediarize_Click()
    bCancel = False
    Me.Hide
End Sub


Public Property Get Cancelled() As Boolean
    Cancelled = bCancel
End Property

Public Property Get Reason() As String
    If Me.optOOP Then
        Reason = "O"
    Else
        Reason = "R"
    End If
    
End Property
