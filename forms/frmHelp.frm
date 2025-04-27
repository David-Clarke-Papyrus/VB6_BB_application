VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help and Tips"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080C0FF&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2745
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7215
      Width           =   795
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Hot Keys"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   7110
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   6105
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":0000
         ForeColor       =   &H008080FF&
         Height          =   585
         Index           =   16
         Left            =   570
         TabIndex        =   18
         Top             =   1500
         Width           =   5250
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":00A2
         ForeColor       =   &H0080C0FF&
         Height          =   1185
         Index           =   15
         Left            =   570
         TabIndex        =   17
         Top             =   6210
         Width           =   5475
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "How to change Operator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Index           =   14
         Left            =   135
         TabIndex        =   16
         Top             =   5940
         Width           =   2520
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opens this window..."
         ForeColor       =   &H0080C0FF&
         Height          =   195
         Index           =   13
         Left            =   1215
         TabIndex        =   15
         Top             =   375
         Width           =   1605
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":019C
         ForeColor       =   &H0080C0FF&
         Height          =   690
         Index           =   12
         Left            =   570
         TabIndex        =   14
         Top             =   885
         Width           =   5475
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Only available after 'Process Sale ' has been activated."
         ForeColor       =   &H0080C0FF&
         Height          =   195
         Index           =   11
         Left            =   570
         TabIndex        =   13
         Top             =   5640
         Width           =   4005
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F12  Open Till"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Index           =   10
         Left            =   135
         TabIndex        =   12
         Top             =   5370
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "All Sale data will be cleared without saving and the application is put into standby mode. To start a new Sale, hit F2."
         ForeColor       =   &H0080C0FF&
         Height          =   450
         Index           =   9
         Left            =   570
         TabIndex        =   11
         Top             =   4845
         Width           =   5385
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F10  Clear Sale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Index           =   8
         Left            =   135
         TabIndex        =   10
         Top             =   4590
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "This will add the same discount to all items. To allocate individual discount to items, select 'Edit Sale Line' mode."
         ForeColor       =   &H0080C0FF&
         Height          =   420
         Index           =   7
         Left            =   570
         TabIndex        =   9
         Top             =   4095
         Width           =   5145
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F9   Add or remove general Discount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Index           =   6
         Left            =   135
         TabIndex        =   8
         Top             =   3855
         Width           =   3795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHelp.frx":0262
         ForeColor       =   &H0080C0FF&
         Height          =   600
         Index           =   5
         Left            =   570
         TabIndex        =   7
         Top             =   3180
         Width           =   5430
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F8   Edit Sale Line"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   6
         Top             =   2940
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F7   Edit Customer Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   2565
         Width           =   2685
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F5   Process Sale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   4
         Top             =   2160
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F2   Start new Sale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   675
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F1   Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   330
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then Unload Me
End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
End Sub
