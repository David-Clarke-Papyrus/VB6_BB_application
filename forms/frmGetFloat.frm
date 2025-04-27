VERSION 5.00
Begin VB.Form frmGetFloat 
   Caption         =   "Float"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   4020
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cancel"
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
      Height          =   465
      Left            =   705
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7785
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3135
      Left            =   240
      TabIndex        =   30
      Top             =   120
      Width           =   3585
      Begin VB.TextBox txtN200 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   1245
         TabIndex        =   0
         Text            =   "0"
         Top             =   510
         Width           =   885
      End
      Begin VB.TextBox txtN100 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   1245
         TabIndex        =   1
         Text            =   "0"
         Top             =   900
         Width           =   885
      End
      Begin VB.TextBox txtN50 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   1245
         TabIndex        =   2
         Text            =   "0"
         Top             =   1290
         Width           =   885
      End
      Begin VB.TextBox txtN20 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   1245
         TabIndex        =   3
         Text            =   "0"
         Top             =   1680
         Width           =   885
      End
      Begin VB.TextBox txtN10 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   1245
         TabIndex        =   4
         Text            =   "0"
         Top             =   2070
         Width           =   885
      End
      Begin VB.CommandButton cmdRefreshNotes 
         Height          =   315
         Left            =   1470
         Picture         =   "frmGetFloat.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2565
         Width           =   330
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "R200"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   375
         TabIndex        =   42
         Top             =   510
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "R100"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   375
         TabIndex        =   41
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "R50"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   375
         TabIndex        =   40
         Top             =   1290
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "R20"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   375
         TabIndex        =   39
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "R10"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   375
         TabIndex        =   38
         Top             =   2070
         Width           =   690
      End
      Begin VB.Line Line2 
         X1              =   105
         X2              =   3270
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Line Line3 
         X1              =   105
         X2              =   3270
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Line Line4 
         X1              =   105
         X2              =   3270
         Y1              =   2025
         Y2              =   2025
      End
      Begin VB.Line Line1 
         X1              =   105
         X2              =   3270
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Label lblN200 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   2295
         TabIndex        =   37
         Top             =   480
         Width           =   960
      End
      Begin VB.Label lblN100 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   2295
         TabIndex        =   36
         Top             =   877
         Width           =   960
      End
      Begin VB.Label lblN50 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   2295
         TabIndex        =   35
         Top             =   1274
         Width           =   960
      End
      Begin VB.Label lblN20 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   2295
         TabIndex        =   34
         Top             =   1671
         Width           =   960
      End
      Begin VB.Label lblN10 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   2295
         TabIndex        =   33
         Top             =   2070
         Width           =   960
      End
      Begin VB.Label lblNotesTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1860
         TabIndex        =   32
         Top             =   2550
         Width           =   1395
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Coins"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3870
      Left            =   240
      TabIndex        =   12
      Top             =   3360
      Width           =   3585
      Begin VB.TextBox txtC20 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   1245
         TabIndex        =   9
         Text            =   "0"
         Top             =   2070
         Width           =   885
      End
      Begin VB.TextBox txtC50 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   1245
         TabIndex        =   8
         Text            =   "0"
         Top             =   1680
         Width           =   885
      End
      Begin VB.TextBox txtC100 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   1245
         TabIndex        =   7
         Text            =   "0"
         Top             =   1275
         Width           =   885
      End
      Begin VB.TextBox txtC200 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   1245
         TabIndex        =   6
         Text            =   "0"
         Top             =   900
         Width           =   885
      End
      Begin VB.TextBox txtC500 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   1245
         TabIndex        =   5
         Text            =   "0"
         Top             =   510
         Width           =   885
      End
      Begin VB.TextBox txtC10 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   1245
         TabIndex        =   10
         Text            =   "0"
         Top             =   2460
         Width           =   885
      End
      Begin VB.TextBox txtC5 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   1245
         TabIndex        =   11
         Text            =   "0"
         Top             =   2850
         Width           =   885
      End
      Begin VB.CommandButton cmdRefreshCoins 
         Height          =   315
         Left            =   1530
         Picture         =   "frmGetFloat.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3360
         Width           =   330
      End
      Begin VB.Label lblC20 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   2295
         TabIndex        =   29
         Top             =   2070
         Width           =   960
      End
      Begin VB.Label lblC50 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   2295
         TabIndex        =   28
         Top             =   1671
         Width           =   960
      End
      Begin VB.Label lblC100 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   2295
         TabIndex        =   27
         Top             =   1274
         Width           =   960
      End
      Begin VB.Label lblC200 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   2295
         TabIndex        =   26
         Top             =   877
         Width           =   960
      End
      Begin VB.Label lblC500 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   2295
         TabIndex        =   25
         Top             =   480
         Width           =   960
      End
      Begin VB.Line Line5 
         X1              =   105
         X2              =   3270
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Line Line6 
         X1              =   105
         X2              =   3270
         Y1              =   2025
         Y2              =   2025
      End
      Begin VB.Line Line7 
         X1              =   105
         X2              =   3270
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Line Line8 
         X1              =   105
         X2              =   3270
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "20c"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   375
         TabIndex        =   24
         Top             =   2070
         Width           =   690
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "50c"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   375
         TabIndex        =   23
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "R1"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   375
         TabIndex        =   22
         Top             =   1290
         Width           =   690
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "R2"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   375
         TabIndex        =   21
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "R5"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   375
         TabIndex        =   20
         Top             =   510
         Width           =   690
      End
      Begin VB.Label lblC10 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   2295
         TabIndex        =   19
         Top             =   2460
         Width           =   960
      End
      Begin VB.Line Line9 
         X1              =   105
         X2              =   3270
         Y1              =   2415
         Y2              =   2415
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "10c"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   375
         TabIndex        =   18
         Top             =   2460
         Width           =   690
      End
      Begin VB.Label lblC5 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   2295
         TabIndex        =   17
         Top             =   2850
         Width           =   960
      End
      Begin VB.Line Line10 
         X1              =   105
         X2              =   3270
         Y1              =   2805
         Y2              =   2805
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "5c"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   375
         TabIndex        =   16
         Top             =   2835
         Width           =   690
      End
      Begin VB.Label lblCoinsTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1905
         TabIndex        =   15
         Top             =   3345
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2085
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7770
      Width           =   1260
   End
   Begin VB.Label lblFinalTotal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1305
      TabIndex        =   43
      Top             =   7335
      Width           =   1395
   End
End
Attribute VB_Name = "frmGetFloat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dblFLoat As Double
Dim qtyN200 As Long
Dim qtyN100 As Long
Dim qtyN50 As Long
Dim qtyN20 As Long
Dim qtyN10 As Long
Dim qtyC500 As Long
Dim qtyC200 As Long
Dim qtyC100 As Long
Dim qtyC50 As Long
Dim qtyC20 As Long
Dim qtyC10 As Long
Dim qtyC5 As Long

Dim dblTotalNotes As Double
Dim dblTotalCoins As Double
Dim dblTotalOverall As Double
Dim bCancelled As Boolean

Public Property Get FloatValue() As Double
    Recalculate
    FloatValue = dblTotalOverall
End Property

Private Sub cmdCancel_Click()
    bCancelled = True
    Me.Hide
End Sub
Public Property Get IsCancelled() As Boolean
    IsCancelled = bCancelled
End Property
Private Sub cmdOK_Click()
    Me.Hide
End Sub



Public Function GetFloatBreakdown() As String
Dim s As String
    s = CStr(qtyN200) & "," & CStr(qtyN100) & "," & CStr(qtyN50) & "," & CStr(qtyN20) & "," & CStr(qtyN10) _
     & "," & CStr(qtyC500) & "," & CStr(qtyC200) & "," & CStr(qtyC100) & "," & CStr(qtyC50) _
      & "," & CStr(qtyC20) & "," & CStr(qtyC10) & "," & CStr(qtyC5)
      GetFloatBreakdown = s
End Function

Private Sub Form_Load()
    bCancelled = False
    Screen.MousePointer = vbDefault
End Sub

'=================
Private Sub txtN200_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtN200)
    If Not Cancel Then
        qtyN200 = Absolute_Lng(CLng(txtN200))
        txtN200 = Format(qtyN200, "0")
    End If
    Recalculate
End Sub
Private Sub txtN100_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtN100)
    If Not Cancel Then
        qtyN100 = Absolute_Lng(CLng(txtN100))
        txtN100 = Format(qtyN100, "0")
    End If
    Recalculate
End Sub
Private Sub txtN50_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtN50)
    If Not Cancel Then
        qtyN50 = Absolute_Lng(CLng(txtN50))
        txtN50 = Format(qtyN50, "0")
    End If
    Recalculate
End Sub
Private Sub txtN20_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtN20)
    If Not Cancel Then
        qtyN20 = Absolute_Lng(CLng(txtN20))
        txtN20 = Format(qtyN20, "0")
    End If
    Recalculate
End Sub
Private Sub txtN10_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtN10)
    If Not Cancel Then
        qtyN10 = Absolute_Lng(CLng(txtN10))
        txtN10 = Format(qtyN10, "0")
    End If
    Recalculate
End Sub
Private Sub txtC500_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtC500)
    If Not Cancel Then
        qtyC500 = Absolute_Lng(CLng(txtC500))
        txtC500 = Format(qtyC500, "0")
    End If
    Recalculate
End Sub
Private Sub txtC200_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtC200)
    If Not Cancel Then
        qtyC200 = Absolute_Lng(CLng(txtC200))
        txtC200 = Format(qtyC200, "0")
    End If
    Recalculate
End Sub
Private Sub txtC100_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtC100)
    If Not Cancel Then
        qtyC100 = Absolute_Lng(CLng(txtC100))
        txtC100 = Format(qtyC100, "0")
    End If
    Recalculate
End Sub
Private Sub txtC50_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtC50)
    If Not Cancel Then
        qtyC50 = Absolute_Lng(CLng(txtC50))
        txtC50 = Format(qtyC50, "0")
    End If
    Recalculate
End Sub
Private Sub txtC20_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtC20)
    If Not Cancel Then
        qtyC20 = Absolute_Lng(CLng(txtC20))
        txtC20 = Format(qtyC20, "0")
    End If
    Recalculate
End Sub
Private Sub txtC10_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtC10)
    If Not Cancel Then
        qtyC10 = Absolute_Lng(CLng(txtC10))
        txtC10 = Format(qtyC10, "0")
    End If
    Recalculate
End Sub
Private Sub txtC5_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtC5)
    If Not Cancel Then
        qtyC5 = Absolute_Lng(CLng(txtC5))
        txtC5 = Format(qtyC5, "0")
    End If
    Recalculate
End Sub
''
Public Sub AutoSelect(ctl As Control)
    ctl.SelStart = 0
    ctl.SelLength = Len(ctl.Text)
End Sub

Private Sub txtN200_GotFocus()
    AutoSelect txtN200
End Sub
Private Sub txtN100_GotFocus()
    AutoSelect txtN100
End Sub
Private Sub txtN50_GotFocus()
    AutoSelect txtN50
End Sub
Private Sub txtN20_GotFocus()
    AutoSelect txtN20
End Sub
Private Sub txtN10_GotFocus()
    AutoSelect txtN10
End Sub
Private Sub txtC500_GotFocus()
    AutoSelect txtC500
End Sub
Private Sub txtC200_GotFocus()
    AutoSelect txtC200
End Sub
Private Sub txtC100_GotFocus()
    AutoSelect txtC100
End Sub
Private Sub txtC50_GotFocus()
    AutoSelect txtC50
End Sub
Private Sub txtC20_GotFocus()
    AutoSelect txtC20
End Sub
Private Sub txtC10_GotFocus()
    AutoSelect txtC10
End Sub
Private Sub txtC5_GotFocus()
    AutoSelect txtC5
End Sub

Private Sub Recalculate()
On Error Resume Next
    dblTotalNotes = CDbl((qtyN200 * 20000) + _
                    (qtyN100 * 10000) + _
                    (qtyN50 * 5000) + _
                    (qtyN20 * 2000) + _
                    (qtyN10 * 1000)) / 100

    dblTotalCoins = CDbl((qtyC500 * 500) + _
                    (qtyC200 * 200) + _
                    (qtyC100 * 100) + _
                    (qtyC50 * 50) + _
                    (qtyC20 * 20) + _
                    (qtyC10 * 10) + _
                    (qtyC5 * 5)) / 100

    dblTotalOverall = dblTotalNotes + dblTotalCoins '+ dblTotalCC
    lblNotesTotal.Caption = Format(dblTotalNotes, "###,##0.00")
    lblCoinsTotal.Caption = Format(dblTotalCoins, "###,##0.00")

    lblN200.Caption = Format(qtyN200 * 200, oPC.CurrencyFormat)
    
    lblN100.Caption = Format(qtyN100 * 100, oPC.CurrencyFormat)
    
    lblN50.Caption = Format(qtyN50 * 50, oPC.CurrencyFormat)
    
    lblN20.Caption = Format(qtyN20 * 20, oPC.CurrencyFormat)
    
    lblN10.Caption = Format(qtyN10 * 10, oPC.CurrencyFormat)
    
    lblC500.Caption = Format(qtyC500 * 5, oPC.CurrencyFormat)
    
    lblC200.Caption = Format(qtyC200 * 2, oPC.CurrencyFormat)
    
    lblC100.Caption = Format(qtyC100 * 1, oPC.CurrencyFormat)
    
    lblC50.Caption = Format(qtyC50 * 0.5, oPC.CurrencyFormat)
    
    lblC20.Caption = Format(qtyC20 * 0.2, oPC.CurrencyFormat)
    
    lblC10.Caption = Format(qtyC10 * 0.1, oPC.CurrencyFormat)
    
    lblC5.Caption = Format(qtyC5 * 0.05, oPC.CurrencyFormat)

    Me.lblFinalTotal.Caption = Format(dblTotalOverall, "###,##0.00")
End Sub

