VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmBlindCashup 
   BackColor       =   &H00D3D3CB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash-up data entry"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   165
   ClientWidth     =   11730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   11730
   Begin VB.Frame Frame6 
      BackColor       =   &H00D3D3CB&
      Caption         =   "1. Count new float"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   60
      TabIndex        =   130
      Top             =   75
      Width           =   3585
      Begin VB.CommandButton cmdCountNewFloat 
         Height          =   330
         Left            =   2835
         Picture         =   "frmBlindCashup.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.TextBox txtExplanation 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   855
      Left            =   8220
      MultiLine       =   -1  'True
      TabIndex        =   128
      Top             =   7275
      Width           =   3360
   End
   Begin VB.Frame frBankedAmount 
      BackColor       =   &H00D3D3CB&
      Caption         =   "7. Banked amount"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1500
      Left            =   3690
      TabIndex        =   112
      Top             =   6645
      Visible         =   0   'False
      Width           =   4485
      Begin VB.TextBox txtReturned 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   2775
         TabIndex        =   136
         Text            =   "0"
         Top             =   810
         Width           =   1230
      End
      Begin VB.TextBox txtBanked 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   127
         Text            =   "0"
         Top             =   1155
         Width           =   1185
      End
      Begin VB.TextBox txtRetained 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   2775
         TabIndex        =   124
         Text            =   "0"
         Top             =   525
         Width           =   1230
      End
      Begin VB.TextBox txtActualCashInDrawer2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   270
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   122
         Text            =   "0"
         Top             =   255
         Width           =   1185
      End
      Begin VB.Label Label10 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Returned"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   195
         TabIndex        =   137
         Top             =   810
         Width           =   2130
      End
      Begin VB.Line Line17 
         BorderColor     =   &H80000010&
         X1              =   255
         X2              =   4005
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label Label35 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "Banked"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   1965
         TabIndex        =   126
         Top             =   1185
         Width           =   1050
      End
      Begin VB.Label Label34 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Retained"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   195
         TabIndex        =   125
         Top             =   540
         Width           =   2130
      End
      Begin VB.Label Label32 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash available"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   195
         TabIndex        =   123
         Top             =   270
         Width           =   2430
      End
   End
   Begin VB.CommandButton cmdStage3 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Sign for H.O."
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
      Height          =   615
      Left            =   6390
      Style           =   1  'Graphical
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   8235
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.Frame frmResults 
      BackColor       =   &H00D3D3CB&
      Caption         =   "6. Comparison"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3105
      Left            =   3705
      TabIndex        =   75
      Top             =   3525
      Width           =   7860
      Begin VB.CommandButton cmdExplainCheque 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Explain"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6570
         Style           =   1  'Graphical
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   1245
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin VB.CommandButton cmdExplainVoucherRedeemed 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Explain"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6555
         Style           =   1  'Graphical
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   2670
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin VB.CommandButton cmdExplainDeposit 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Explain"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6570
         Style           =   1  'Graphical
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   2220
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin VB.CommandButton cmdExplainCard 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Explain"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6555
         Style           =   1  'Graphical
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   1725
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin VB.CommandButton cmdExplainCash 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Explain"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6555
         Style           =   1  'Graphical
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   750
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin VB.TextBox txtVoucherExplanation 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   195
         Left            =   1290
         TabIndex        =   106
         Top             =   2760
         Width           =   5205
      End
      Begin VB.TextBox txtDepositExplanation 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   195
         Left            =   1305
         TabIndex        =   105
         Top             =   2310
         Width           =   5205
      End
      Begin VB.TextBox txtCardExplanation 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   195
         Left            =   1305
         TabIndex        =   104
         Top             =   1800
         Width           =   5205
      End
      Begin VB.TextBox txtChequeExplanation 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   195
         Left            =   1305
         TabIndex        =   103
         Top             =   1320
         Width           =   5205
      End
      Begin VB.TextBox txtCashExplanation 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   195
         Left            =   1290
         TabIndex        =   98
         Top             =   825
         Width           =   5205
      End
      Begin VB.Label lblTotalChequesS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   3600
         TabIndex        =   102
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1110
         TabIndex        =   101
         Top             =   1065
         Width           =   1230
      End
      Begin VB.Label lblDiffCheques 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   4770
         TabIndex        =   100
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label lblTotalCheques 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   2430
         TabIndex        =   99
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Line Line15 
         BorderColor     =   &H80000010&
         X1              =   780
         X2              =   7280
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Diff."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   270
         Left            =   4845
         TabIndex        =   96
         Top             =   255
         Width           =   1185
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Calc."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   270
         Left            =   3660
         TabIndex        =   95
         Top             =   255
         Width           =   1065
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "In drawer"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   270
         Left            =   2265
         TabIndex        =   94
         Top             =   255
         Width           =   1290
      End
      Begin VB.Label lblTotalVouchersRedeemed 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   330
         Left            =   2430
         TabIndex        =   93
         Top             =   2535
         Width           =   1125
      End
      Begin VB.Label lblTotalCash 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   2430
         TabIndex        =   92
         Top             =   540
         Width           =   1125
      End
      Begin VB.Label lblTotalCreditCards 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   300
         Left            =   2430
         TabIndex        =   91
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label lblTotalDeposits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   2430
         TabIndex        =   90
         Top             =   2040
         Width           =   1125
      End
      Begin VB.Label lblDiffVouchersRedeemed 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   330
         Left            =   4770
         TabIndex        =   87
         Top             =   2535
         Width           =   1260
      End
      Begin VB.Label lblDiffDeposits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   4770
         TabIndex        =   86
         Top             =   2040
         Width           =   1260
      End
      Begin VB.Label lblDiffCreditCards 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   300
         Left            =   4770
         TabIndex        =   85
         Top             =   1560
         Width           =   1260
      End
      Begin VB.Label lblDiffCash 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   4770
         TabIndex        =   84
         Top             =   540
         Width           =   1260
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vouchers redeemed"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   225
         TabIndex        =   83
         Top             =   2520
         Width           =   2115
      End
      Begin VB.Line Line16 
         BorderColor     =   &H80000010&
         X1              =   780
         X2              =   7280
         Y1              =   2505
         Y2              =   2505
      End
      Begin VB.Label lblTotalVouchersRedeemedS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   360
         Left            =   3600
         TabIndex        =   82
         Top             =   2535
         Width           =   1140
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   1290
         TabIndex        =   81
         Top             =   540
         Width           =   1050
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cards"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   375
         TabIndex        =   80
         Top             =   1560
         Width           =   1965
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Direct deposits"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   225
         TabIndex        =   79
         Top             =   2040
         Width           =   2115
      End
      Begin VB.Line Line13 
         BorderColor     =   &H80000010&
         X1              =   780
         X2              =   7280
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line12 
         BorderColor     =   &H80000010&
         X1              =   765
         X2              =   7265
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Label lblTOtalCashS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   3600
         TabIndex        =   78
         Top             =   540
         Width           =   1140
      End
      Begin VB.Label lblTotalCreditCardsS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   330
         Left            =   3600
         TabIndex        =   77
         Top             =   1560
         Width           =   1140
      End
      Begin VB.Label lblTotalDepositsS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   3600
         TabIndex        =   76
         Top             =   2040
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdStage2 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Sign (stage 2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.Frame frCards 
      BackColor       =   &H00D3D3CB&
      Caption         =   "4. Cards"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1260
      Left            =   3795
      TabIndex        =   70
      Top             =   105
      Width           =   2940
      Begin VB.TextBox txtCCards 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   1710
         TabIndex        =   16
         Text            =   "0"
         Top             =   405
         Width           =   1125
      End
      Begin VB.TextBox txtDCards 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   1710
         TabIndex        =   17
         Text            =   "0"
         Top             =   810
         Width           =   1125
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit cards"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   150
         TabIndex        =   72
         Top             =   405
         Width           =   1470
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Debit cards"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   150
         TabIndex        =   71
         Top             =   795
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Sa&ve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   4950
      MaskColor       =   &H00C4BCA4&
      Style           =   1  'Graphical
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   8220
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.CommandButton cmdStage1 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Sign (stage 1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   8235
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3690
      Picture         =   "frmBlindCashup.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   8220
      Width           =   1110
   End
   Begin VB.Frame frOther 
      BackColor       =   &H00D3D3CB&
      Caption         =   "5. Other"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3360
      Left            =   7080
      TabIndex        =   63
      Top             =   105
      Width           =   4470
      Begin VB.CommandButton cmdEndingFloat 
         Height          =   330
         Left            =   3900
         Picture         =   "frmBlindCashup.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   1305
         Width           =   375
      End
      Begin VB.TextBox txtFloatAtEnd 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   270
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   134
         Text            =   "0"
         Top             =   1260
         Width           =   1230
      End
      Begin VB.CommandButton cmdStartingFloat 
         Height          =   330
         Left            =   3900
         Picture         =   "frmBlindCashup.frx":0A9E
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   900
         Width           =   375
      End
      Begin VB.TextBox txtCalcCashInDrawer 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   270
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   120
         Text            =   "0"
         Top             =   1800
         Width           =   1185
      End
      Begin VB.TextBox txtCashFromSales 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   270
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   116
         Text            =   "0"
         Top             =   690
         Width           =   1185
      End
      Begin VB.TextBox txtActualCashInDrawer 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   270
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   114
         Text            =   "0"
         Top             =   285
         Width           =   1185
      End
      Begin VB.TextBox txtPettyCashNett 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   270
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "0"
         Top             =   1515
         Width           =   1185
      End
      Begin VB.TextBox txtFloatAtStart 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   270
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   88
         Text            =   "0"
         Top             =   975
         Width           =   1185
      End
      Begin VB.TextBox txtVouchersRedeemed 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   2610
         TabIndex        =   15
         Text            =   "0"
         Top             =   2910
         Width           =   1230
      End
      Begin VB.TextBox txtDeposits 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   2610
         TabIndex        =   14
         Text            =   "0"
         Top             =   2625
         Width           =   1230
      End
      Begin VB.TextBox txtCheques 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   2610
         TabIndex        =   13
         Text            =   "0"
         Top             =   2340
         Width           =   1230
      End
      Begin VB.Label Label36 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Closing float"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   135
         TabIndex        =   133
         Top             =   1230
         Width           =   2430
      End
      Begin VB.Line Line14 
         BorderColor     =   &H80000010&
         X1              =   165
         X2              =   3915
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Line Line11 
         BorderColor     =   &H80000010&
         X1              =   150
         X2              =   3900
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label29 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "System"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   120
         TabIndex        =   121
         Top             =   1800
         Width           =   2490
      End
      Begin VB.Label Label26 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "    Cash from sales"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   120
         TabIndex        =   119
         Top             =   690
         Width           =   2430
      End
      Begin VB.Label Label21 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Opening float"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   135
         TabIndex        =   118
         Top             =   960
         Width           =   2430
      End
      Begin VB.Label Label24 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Petty cash nett"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   135
         TabIndex        =   117
         Top             =   1500
         Width           =   2430
      End
      Begin VB.Label Label25 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "Counted"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   135
         TabIndex        =   115
         Top             =   300
         Width           =   2430
      End
      Begin VB.Label Label9 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher/CN redeemed"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   150
         TabIndex        =   66
         Top             =   2925
         Width           =   2325
      End
      Begin VB.Label Label8 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "Direct deposits"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   150
         TabIndex        =   65
         Top             =   2640
         Width           =   2130
      End
      Begin VB.Label Label7 
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   150
         TabIndex        =   64
         Top             =   2355
         Width           =   930
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Vouchers"
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
      Height          =   7110
      Left            =   20685
      TabIndex        =   57
      Top             =   225
      Width           =   3315
      Begin VB.CommandButton Command2 
         Height          =   300
         Left            =   300
         Picture         =   "frmBlindCashup.frx":0E28
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   6480
         Width           =   180
      End
      Begin VB.CommandButton Command1 
         Height          =   300
         Left            =   45
         Picture         =   "frmBlindCashup.frx":11B2
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   6480
         Width           =   180
      End
      Begin TrueOleDBGrid60.TDBGrid GV 
         Height          =   6030
         Left            =   105
         OleObjectBlob   =   "frmBlindCashup.frx":153C
         TabIndex        =   60
         Top             =   360
         Width           =   1620
      End
      Begin TrueOleDBGrid60.TDBGrid TDBGrid2 
         Height          =   6030
         Left            =   1740
         OleObjectBlob   =   "frmBlindCashup.frx":387B
         TabIndex        =   61
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label Label6 
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
         Left            =   585
         TabIndex        =   62
         Top             =   6465
         Width           =   1395
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cheques"
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
      Height          =   7110
      Left            =   17400
      TabIndex        =   51
      Top             =   225
      Width           =   3315
      Begin VB.CommandButton cmd1Up 
         Height          =   300
         Left            =   45
         Picture         =   "frmBlindCashup.frx":5BC0
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   6480
         Width           =   180
      End
      Begin VB.CommandButton cmd1Down 
         Height          =   300
         Left            =   300
         Picture         =   "frmBlindCashup.frx":5F4A
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   6480
         Width           =   180
      End
      Begin TrueOleDBGrid60.TDBGrid GCH 
         Height          =   6030
         Left            =   105
         OleObjectBlob   =   "frmBlindCashup.frx":62D4
         TabIndex        =   54
         Top             =   360
         Width           =   1620
      End
      Begin TrueOleDBGrid60.TDBGrid G1C 
         Height          =   6030
         Left            =   1740
         OleObjectBlob   =   "frmBlindCashup.frx":8618
         TabIndex        =   55
         Top             =   375
         Width           =   1485
      End
      Begin VB.Label lblTotalCH 
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
         Left            =   585
         TabIndex        =   56
         Top             =   6465
         Width           =   1395
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Credit cards"
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
      Height          =   7110
      Left            =   13980
      TabIndex        =   45
      Top             =   225
      Width           =   3450
      Begin VB.CommandButton cmdDown 
         Height          =   300
         Left            =   315
         Picture         =   "frmBlindCashup.frx":A95C
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   6480
         Width           =   180
      End
      Begin VB.CommandButton cmdUp 
         Height          =   300
         Left            =   60
         Picture         =   "frmBlindCashup.frx":ACE6
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   6480
         Width           =   180
      End
      Begin TrueOleDBGrid60.TDBGrid GCC 
         Height          =   6030
         Left            =   75
         OleObjectBlob   =   "frmBlindCashup.frx":B070
         TabIndex        =   46
         Top             =   360
         Width           =   1695
      End
      Begin TrueOleDBGrid60.TDBGrid GC 
         Height          =   6030
         Left            =   1785
         OleObjectBlob   =   "frmBlindCashup.frx":D3BC
         TabIndex        =   48
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label lblTotalCC 
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
         Left            =   585
         TabIndex        =   47
         Top             =   6465
         Width           =   1395
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "3. Coins"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3720
      Left            =   60
      TabIndex        =   28
      Top             =   4140
      Width           =   3585
      Begin VB.TextBox txtC5 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
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
         TabIndex        =   12
         Text            =   "0"
         Top             =   2850
         Width           =   705
      End
      Begin VB.TextBox txtC10 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
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
         Top             =   2460
         Width           =   705
      End
      Begin VB.TextBox txtC500 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
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
         Top             =   510
         Width           =   705
      End
      Begin VB.TextBox txtC200 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
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
         Top             =   900
         Width           =   705
      End
      Begin VB.TextBox txtC100 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
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
         Top             =   1275
         Width           =   705
      End
      Begin VB.TextBox txtC50 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
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
         Top             =   1680
         Width           =   705
      End
      Begin VB.TextBox txtC20 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
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
         Top             =   2070
         Width           =   705
      End
      Begin VB.Label lblCoinsTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   1995
         TabIndex        =   44
         Top             =   3255
         Width           =   1260
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "5c"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   375
         TabIndex        =   42
         Top             =   2835
         Width           =   690
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000010&
         X1              =   105
         X2              =   3270
         Y1              =   2805
         Y2              =   2805
      End
      Begin VB.Label lblC5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   1995
         TabIndex        =   41
         Top             =   2850
         Width           =   1260
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "10c"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   375
         TabIndex        =   40
         Top             =   2460
         Width           =   690
      End
      Begin VB.Line Line9 
         BorderColor     =   &H80000010&
         X1              =   105
         X2              =   3270
         Y1              =   2415
         Y2              =   2415
      End
      Begin VB.Label lblC10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   1995
         TabIndex        =   39
         Top             =   2460
         Width           =   1260
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R5"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   375
         TabIndex        =   38
         Top             =   510
         Width           =   690
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R2"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   375
         TabIndex        =   37
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   375
         TabIndex        =   36
         Top             =   1290
         Width           =   690
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "50c"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   375
         TabIndex        =   35
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "20c"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   375
         TabIndex        =   34
         Top             =   2070
         Width           =   690
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000010&
         X1              =   105
         X2              =   3270
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000010&
         X1              =   105
         X2              =   3270
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000010&
         X1              =   105
         X2              =   3270
         Y1              =   2025
         Y2              =   2025
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000010&
         X1              =   105
         X2              =   3270
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Label lblC500 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   1995
         TabIndex        =   33
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label lblC200 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   1995
         TabIndex        =   32
         Top             =   870
         Width           =   1260
      End
      Begin VB.Label lblC100 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   1995
         TabIndex        =   31
         Top             =   1275
         Width           =   1260
      End
      Begin VB.Label lblC50 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   1995
         TabIndex        =   30
         Top             =   1665
         Width           =   1260
      End
      Begin VB.Label lblC20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   1995
         TabIndex        =   29
         Top             =   2070
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "2. Notes"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2910
      Left            =   60
      TabIndex        =   0
      Top             =   900
      Width           =   3585
      Begin VB.TextBox txtN10 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
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
         Top             =   2070
         Width           =   690
      End
      Begin VB.TextBox txtN20 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
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
         Top             =   1680
         Width           =   690
      End
      Begin VB.TextBox txtN50 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
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
         Top             =   1290
         Width           =   690
      End
      Begin VB.TextBox txtN100 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
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
         Top             =   900
         Width           =   690
      End
      Begin VB.TextBox txtN200 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
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
         Top             =   510
         Width           =   690
      End
      Begin VB.Label lblNotesTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   1950
         TabIndex        =   43
         Top             =   2460
         Width           =   1305
      End
      Begin VB.Label lblN10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   1950
         TabIndex        =   27
         Top             =   2070
         Width           =   1305
      End
      Begin VB.Label lblN20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   1950
         TabIndex        =   26
         Top             =   1665
         Width           =   1305
      End
      Begin VB.Label lblN50 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   1950
         TabIndex        =   25
         Top             =   1275
         Width           =   1305
      End
      Begin VB.Label lblN100 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   1950
         TabIndex        =   24
         Top             =   870
         Width           =   1305
      End
      Begin VB.Label lblN200 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   345
         Left            =   1950
         TabIndex        =   23
         Top             =   480
         Width           =   1305
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   105
         X2              =   3270
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   105
         X2              =   3270
         Y1              =   2025
         Y2              =   2025
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   105
         X2              =   3270
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   105
         X2              =   3270
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R10"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   375
         TabIndex        =   22
         Top             =   2070
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R20"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   375
         TabIndex        =   21
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R50"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   375
         TabIndex        =   20
         Top             =   1290
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R100"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   375
         TabIndex        =   19
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R200"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   375
         TabIndex        =   18
         Top             =   510
         Width           =   690
      End
   End
   Begin VB.Label lblHelp1 
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   6510
      TabIndex        =   138
      Top             =   8430
      Width           =   195
   End
   Begin VB.Label lblReport 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8220
      TabIndex        =   129
      Top             =   6705
      Width           =   3360
   End
   Begin VB.Label lblStaff 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000011&
      Height          =   600
      Left            =   8655
      TabIndex        =   89
      Top             =   8220
      Width           =   2865
   End
   Begin VB.Label lblStageNumber 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stage 0"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   1770
      TabIndex        =   74
      Top             =   8400
      Visible         =   0   'False
      Width           =   1620
   End
End
Attribute VB_Name = "frmBlindCashup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCU As z_Cashup
Dim clr
Dim rsBlindCashup As ADODB.Recordset
Dim sFloatBreakdownatStart As String
Dim sFloatBreakdownatEnd As String

Dim mStageNumber As Integer
Dim mWorkstationID As Long
Dim mWorkstationName As String

Dim mCapturedByStaffName As String
Dim mIssuedByStaffName As String
Dim mExplainedByStaffName As String

Dim mCaptureDate As Date
Dim mIssueDate As Date

Dim mXID As String

Dim mCapturedByStaffID As Long
Dim mIssuedByStaffID As Long
Dim mExplainedByStaffID As Long
Dim mChequesVal As Double
Dim mVouchersVal As Double
Dim mDepositsVal As Double
Dim mCCardsVal As Double
Dim mDCardsVal As Double

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
Dim dblTotalCC As Double
Dim dblTotalCH As Double
Dim dblTotalV As Double

Dim dblCashDiff As Double
Dim dblChequesDiff As Double
Dim dblCardsDiff As Double
Dim dblCDepositsDiff As Double
Dim dblVouchersDiff As Double
Dim dblFloatAtEnd As Double
Dim dblFloatAtStart As Double
Dim dblShortage As Double
Dim dblBanked As Double
Dim dblPettyCashNett As Double

Dim dblCalcCashInDrawer As Double

Dim sReconciliationMessage As String

Dim dblActualCashTotal As Double
Dim dblRetained As Double
Dim dblReturned As Double
Dim bCashExplanationNeeded As Boolean
Dim bCardExplanationNeeded As Boolean
Dim bChequeExplanationNeeded As Boolean
Dim bDepositExplanationNeeded As Boolean
Dim bVoucherRedeemedExplanationNeeded As Boolean

Dim strCashExplanation As String
Dim strCardExplanation As String
Dim strChequeExplanation As String
Dim strDepositExplanation As String
Dim strVoucherRedeemedExplanation As String
Dim strExplanation As String
Dim strGroupExplanation As String

Dim XCC As New XArrayDB
Dim XCH As New XArrayDB
Dim XV As New XArrayDB
Dim oSQL As New z_SQL
Public Sub component(XID As String)
    On Error GoTo errHandler
  '  Me.lblStageNumber.Visible = True
    If XID > "" Then
        mXID = XID
        Set rsBlindCashup = New ADODB.Recordset
        rsBlindCashup.CursorLocation = adUseClient
        oSQL.LoadBlindCashup rsBlindCashup, mXID
        LoadCashup
    End If
    Me.cmdStartingFloat.Enabled = (mStageNumber > 0)
    cmdStage1.Visible = (mStageNumber = 0)
    cmdStage2.Visible = (Not oPC.IsFrontDeskWorkstation) And mStageNumber = 1
    cmdStage3.Visible = (Not oPC.IsFrontDeskWorkstation) And mStageNumber = 2
    ValidateExplanations
    cmdSave.Visible = oPC.IsFrontDeskWorkstation And mStageNumber = 0

    If oPC.IsFrontDeskWorkstation Then LockFormControlsforFront
    If oPC.IsFrontDeskWorkstation = False Then LockFormControlsForBack

    frCards.Visible = (Not oPC.IsFrontDeskWorkstation)
    frmResults.Visible = Not oPC.IsFrontDeskWorkstation And mStageNumber > 1
    Me.frBankedAmount.Visible = mStageNumber >= 1
    lblStageNumber.Caption = CStr(mStageNumber)
    Me.TOP = 1000
    Me.Left = 400
    If mStageNumber < 2 Then
        Me.Width = 12000
        Me.Height = 9500
    ElseIf mStageNumber > 1 Then
        Me.Width = 12000
        Me.Height = 9500
    End If
    If Not oCU Is Nothing Then
        dblShortage = (mChequesVal + mVouchersVal + mDepositsVal + mCCardsVal + mDCardsVal + dblActualCashTotal) - _
                        (oCU.TotalChequesDec + oCU.TotalVouchersRedeemedDec + oCU.TotalDirectDepositsDec + oCU.TotalCreditCardsDec + oCU.TotalCreditCardsRefundsDec + oCU.TotalCashInDrawerDec + dblFloatAtStart - dblFloatAtEnd)
        If dblShortage < 0 Then
            sReconciliationMessage = "You are SHORT by " & Format(dblShortage * -1, oPC.Configuration.DefaultCurrency.FormatString) & ". Supply a reason and pay in."
        Else
            sReconciliationMessage = "You are OVER by " & Format(dblShortage, oPC.Configuration.DefaultCurrency.FormatString) & ". Supply a reason."
        End If
        If mStageNumber > 1 Then Me.lblReport.Caption = sReconciliationMessage
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.component(XID)", XID
End Sub
''
Private Sub Recalculate()
    On Error GoTo errHandler
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

    dblActualCashTotal = dblTotalNotes + dblTotalCoins
    lblNotesTotal.Caption = Format(dblTotalNotes, "###,##0.00")
    lblCoinsTotal.Caption = Format(dblTotalCoins, "###,##0.00")
    CalcBankedAmount

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.Recalculate"
End Sub
''
''
''
''
Private Sub cmdClose_Click()
    On Error GoTo errHandler
    If MsgBox("Close form?", vbOKCancel + vbQuestion, "Confirm") = vbOK Then
        Unload Me
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCountNewFloat_Click()
    On Error GoTo errHandler
    EndingFloat
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.cmdCountNewFloat_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub EndingFloat()
    On Error GoTo errHandler
Dim f As New frmGetFloat
    f.component (mStageNumber > 1), sFloatBreakdownatEnd
    f.Show vbModal
    sFloatBreakdownatEnd = f.GetFloatBreakdown
    CalcFloatValue sFloatBreakdownatEnd
    txtFloatAtEnd = Format(dblFloatAtEnd, oPC.Configuration.DefaultCurrency.FormatString)
    CalcBankedAmount
    ValidateExplanations
    Unload f
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.EndingFloat"
End Sub
Private Sub CalcFloatValue(s As String)
    On Error GoTo errHandler
Dim ar() As String
    ar() = Split(s, ",")
    dblFloatAtEnd = CDbl((FNN(ar(0)) * 20000) + _
                    (FNN(ar(1)) * 10000) + _
                    (FNN(ar(2)) * 5000) + _
                    (FNN(ar(3)) * 2000) + _
                    (FNN(ar(4)) * 1000)) / 100 + _
                     CDbl((FNN(ar(5)) * 500) + _
                    (FNN(ar(6)) * 200) + _
                    (FNN(ar(7)) * 100) + _
                    (FNN(ar(8)) * 50) + _
                    (FNN(ar(9)) * 20) + _
                    (FNN(ar(10)) * 10) + _
                    (FNN(ar(11)) * 5)) / 100


    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    txtActualCashInDrawer2 = Format(dblActualCashTotal, "###,##0.00")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.CalcFloatValue(s)", s
End Sub
Private Sub CalcBankedAmount()
    On Error GoTo errHandler
    dblBanked = dblActualCashTotal - dblRetained + dblReturned
    txtBanked = Format(dblBanked, "###,##0.00")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.CalcBankedAmount"
End Sub
Public Property Get FloatBreakdownatEnd() As String
    FloatBreakdownatEnd = sFloatBreakdownatEnd
End Property

Private Sub cmdEndingFloat_Click()
    On Error GoTo errHandler
    EndingFloat
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.cmdEndingFloat_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdExplainCard_Click()
    On Error GoTo errHandler
    ReceiveExplanation "Card", strCardExplanation
   ' bCardExplanationNeeded = Len(strCardExplanation) < 5
    cmdExplainCard.Enabled = bCardExplanationNeeded
    txtCardExplanation.text = strCardExplanation
    ValidateExplanations
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.cmdExplainCard_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdExplainCash_Click()
    On Error GoTo errHandler
    ReceiveExplanation "Cash", strCashExplanation
  '  bCashExplanationNeeded = Len(strCashExplanation) < 5
    cmdExplainCash.Enabled = bCashExplanationNeeded
    txtCashExplanation.text = strCashExplanation
    
    ValidateExplanations
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.cmdExplainCash_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdExplainCheque_Click()
    On Error GoTo errHandler
    ReceiveExplanation "Cheque", strChequeExplanation
  '  bChequeExplanationNeeded = Len(strChequeExplanation) < 5
    cmdExplainCheque.Enabled = bChequeExplanationNeeded
    txtChequeExplanation.text = strChequeExplanation
    ValidateExplanations
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.cmdExplainCheque_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdExplainDeposit_Click()
    On Error GoTo errHandler
    ReceiveExplanation "Deposit", strDepositExplanation
  '  bDepositExplanationNeeded = Len(strDepositExplanation) < 5
    cmdExplainDeposit.Enabled = bDepositExplanationNeeded
    txtDepositExplanation.text = strDepositExplanation
    ValidateExplanations
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.cmdExplainDeposit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdExplainVoucherRedeemed_Click()
    On Error GoTo errHandler
    ReceiveExplanation "VoucherRedeemed", strVoucherRedeemedExplanation
  '  bVoucherRedeemedExplanationNeeded = Len(strVoucherRedeemedExplanation) < 5
    cmdExplainVoucherRedeemed.Enabled = bVoucherRedeemedExplanationNeeded
    Me.txtVoucherExplanation.text = strVoucherRedeemedExplanation
    ValidateExplanations
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.cmdExplainVoucherRedeemed_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub ReceiveExplanation(txt As String, sExplanation As String)
    On Error GoTo errHandler
Dim f As New frmOverUnderExplanation
    
    f.component sExplanation
    f.Show vbModal
    sExplanation = f.Explanation
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.ReceiveExplanation(txt,sExplanation)", Array(txt, sExplanation)
End Sub
''
Private Sub cmdSave_Click()
    On Error GoTo errHandler

    SaveCashup

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdStage1_Click()
    On Error GoTo errHandler
    If SecurityControl(enSECURITY_ISOPERATOR, , "Sign this cash up", DOCAPPROVAL, , , gSTAFFID) = False Then
           Exit Sub
    End If

    mCapturedByStaffID = gSTAFFID

    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    If dblActualCashTotal = 0 Then
        If MsgBox("The cashup as a value of zero. Are you sure you want to save this?", vbOKCancel + vbQuestion, "Warning") = vbCancel Then
            Exit Sub
        End If
    End If
    dblBanked = dblActualCashTotal + (dblFloatAtStart - dblFloatAtEnd)

    SaveCashup

    Unload Me
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.cmdStage1_Click", , EA_NORERAISE
    HandleError
End Sub
''
Private Sub cmdStage2_Click()
    On Error GoTo errHandler

    If dblBanked = 0 Then
        If MsgBox("You are not banking anything. Continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If

    If SecurityControl(enSECURITY_CASHUP_SIGN, , "Sign stage 2 (Confirm till contents and Credit and debit cards.)", DOCAPPROVAL, , , gSTAFFID) = False Then
           Exit Sub
    End If

    mIssuedByStaffID = gSTAFFID
    
    dblPettyCashNett = oCU.TotalPettyCashNettDec

    ReportDifferences
    SaveCashup

    Unload Me

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.cmdStage2_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdStage3_Click()
    On Error GoTo errHandler

    If dblFloatAtEnd = 0 Then
        MsgBox "You have not supplied a closing float. You cannot continue.", vbOKOnly, "Warning"
    End If
    If SecurityControl(enSECURITY_CASHUP_SIGN, , "Sign final stage (for transmission to Head Office.)", DOCAPPROVAL, , , gSTAFFID) = False Then
           Exit Sub
    End If

    mExplainedByStaffID = gSTAFFID
    
    ReportDifferences
    SaveCashup

    Unload Me

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.cmdStage3_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdStartingFloat_Click()
    On Error GoTo errHandler
Dim f As New frmGetFloat
'Unload Me
    f.component True, sFloatBreakdownatStart
    f.Show vbModal
    Unload f
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.cmdStartingFloat_Click", , EA_NORERAISE
    HandleError
End Sub

''
Private Sub Form_Initialize()
    On Error GoTo errHandler

    XCC.ReDim 1, 10, 1, 1
    GCC.Array = XCC
    GCC.ReBind
    XCH.ReDim 1, 10, 1, 1
    GCH.Array = XCH
    GCH.ReBind
    XV.ReDim 1, 10, 1, 1
    GV.Array = XV
    GV.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub


'==================================================
Private Sub txtCheques_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtCheques)
    If Not Cancel Then
        mChequesVal = CDbl(txtCheques)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtCheques_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtDeposits_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtDeposits)
    If Not Cancel Then
        mDepositsVal = CDbl(txtDeposits)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtDeposits_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtExplanation_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    strExplanation = stripCRLF(txtExplanation)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtExplanation_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtVouchersRedeemed_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtVouchersRedeemed)
    If Not Cancel Then
        mVouchersVal = CDbl(txtVouchersRedeemed)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtVouchersRedeemed_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtFloatAtEnd_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtFloatAtEnd)
    If Not Cancel Then
        dblFloatAtEnd = CDbl(txtFloatAtEnd)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtFloatAtEnd_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtCCards_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtCCards)
    If Not Cancel Then
        mCCardsVal = CDbl(txtCCards)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtCCards_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtDCards_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtDCards)
    If Not Cancel Then
        mDCardsVal = CDbl(txtDCards)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtDCards_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
''
''
Private Sub txtN200_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtN200
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtN200_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtN100_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtN100
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtN100_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtN50_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtN50
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtN50_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtN20_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtN20
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtN20_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtN10_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtN10
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtN10_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtC500_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtC500
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtC500_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtC200_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtC200
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtC200_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtC100_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtC100
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtC100_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtC50_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtC50
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtC50_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtC20_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtC20
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtC20_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtC10_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtC10
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtC10_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtC5_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtC5
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtC5_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDCards_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtDCards
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtDCards_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDeposits_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtDeposits
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtDeposits_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFloatAtEnd_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtFloatAtEnd
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtFloatAtEnd_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtVouchersRedeemed_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtVouchersRedeemed
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtVouchersRedeemed_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtCheques_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtCheques
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtCheques_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtCCards_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtCCards
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtCCards_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtRetained_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtRetained)
    If Not Cancel Then
        dblRetained = CDbl(txtRetained)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtRetained_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtReturned_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtReturned)
    If Not Cancel Then
        dblReturned = CDbl(txtReturned)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtReturned_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtN200_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtN200)
    If Not Cancel Then
        qtyN200 = CLng(txtN200)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    lblN200.Caption = Format(qtyN200 * 200, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtN200_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtN100_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtN100)
    If Not Cancel Then
        qtyN100 = CLng(txtN100)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    lblN100.Caption = Format(qtyN100 * 100, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtN100_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtN50_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtN50)
    If Not Cancel Then
        qtyN50 = CLng(txtN50)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    lblN50.Caption = Format(qtyN50 * 50, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtN50_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtN20_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtN20)
    If Not Cancel Then
        qtyN20 = CLng(txtN20)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    lblN20.Caption = Format(qtyN20 * 20, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtN20_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtN10_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtN10)
    If Not Cancel Then
        qtyN10 = CLng(txtN10)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    lblN10.Caption = Format(qtyN10 * 10, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtN10_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtC500_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtC500)
    If Not Cancel Then
        qtyC500 = CLng(txtC500)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    lblC500.Caption = Format(qtyC500 * 5, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtC500_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtC200_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtC200)
    If Not Cancel Then
        qtyC200 = CLng(txtC200)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    lblC200.Caption = Format(qtyC200 * 2, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtC200_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtC100_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtC100)
    If Not Cancel Then
        qtyC100 = CLng(txtC100)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    lblC100.Caption = Format(qtyC100, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtC100_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtC50_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtC50)
    If Not Cancel Then
        qtyC50 = CLng(txtC50)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    lblC50.Caption = Format(qtyC50 * 0.5, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtC50_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtC20_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtC20)
    If Not Cancel Then
        qtyC20 = CLng(txtC20)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    lblC20.Caption = Format(qtyC20 * 0.2, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtC20_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtC10_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtC10)
    If Not Cancel Then
        qtyC10 = CLng(txtC10)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    lblC10.Caption = Format(qtyC10 * 0.1, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtC10_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtC5_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsNumeric(txtC5)
    If Not Cancel Then
        qtyC5 = CLng(txtC5)
    End If
    Recalculate
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    
    lblC5.Caption = Format(qtyC5 * 0.05, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.txtC5_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
''
Public Sub AutoSelect(ctl As Control)
    On Error GoTo errHandler
    ctl.SelStart = 0
    ctl.SelLength = Len(ctl.text)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.AutoSelect(ctl)", ctl
End Sub

Private Sub LoadCashup()
    On Error GoTo errHandler
Dim oSQL As New z_SQL
Dim ar() As String

    If mXID = "" Then Exit Sub
    If rsBlindCashup.eof Then Exit Sub
    
    mWorkstationID = FNN(rsBlindCashup.Fields("WorkstationID"))
    mWorkstationName = FNS(rsBlindCashup.Fields("WorkstationName"))
    Me.Caption = Me.Caption & "     " & FNS(rsBlindCashup.Fields("Tillpoint")) ' & "     " & mCaptureDate
    mCapturedByStaffName = FNS(rsBlindCashup.Fields("CapturedByStaffName"))
    mIssuedByStaffName = FNS(rsBlindCashup.Fields("IssuedByStaffName"))
    mExplainedByStaffName = FNS(rsBlindCashup.Fields("ExplainedByStaffName"))
    mCapturedByStaffID = FNN(rsBlindCashup.Fields("CapturedByStaffID"))
    mIssuedByStaffID = FNN(rsBlindCashup.Fields("IssuedByStaffID"))
    mExplainedByStaffID = FNN(rsBlindCashup.Fields("ExplainedByStaffID"))
    sFloatBreakdownatStart = FNS(rsBlindCashup.Fields("FloatBreakdownatStart"))
    sFloatBreakdownatEnd = FNS(rsBlindCashup.Fields("FloatBreakdownatEnd"))
    lblStaff.Caption = "Captured by: " & mCapturedByStaffName & vbCrLf & "Approved by: " & mIssuedByStaffName & vbCrLf & "Finalized by: " & mExplainedByStaffName
    
    txtFloatAtStart = Format(FNDBL(rsBlindCashup.Fields("FloatValueatStart")), oPC.Configuration.LocalCurrency.FormatString)
    dblFloatAtStart = FNDBL(rsBlindCashup.Fields("FloatValueatStart"))
    
    dblFloatAtEnd = FNDBL(rsBlindCashup.Fields("Float"))
    txtFloatAtEnd = Format(dblFloatAtEnd, oPC.Configuration.LocalCurrency.FormatString)
    
    mChequesVal = FNDBL(rsBlindCashup.Fields("Cheques"))
    txtCheques = Format(mChequesVal, oPC.Configuration.LocalCurrency.FormatString)
    
    mVouchersVal = FNDBL(rsBlindCashup.Fields("Vouchers"))
    txtVouchersRedeemed = Format(mVouchersVal, oPC.Configuration.LocalCurrency.FormatString)
    
    mDepositsVal = FNDBL(rsBlindCashup.Fields("DirectDeposits"))
    txtDeposits = Format(mDepositsVal, oPC.Configuration.LocalCurrency.FormatString)
    
    mCCardsVal = FNDBL(rsBlindCashup.Fields("CreditCards"))
    txtCCards = Format(mCCardsVal, oPC.Configuration.LocalCurrency.FormatString)
    
    mDCardsVal = FNDBL(rsBlindCashup.Fields("DebitCards"))
    txtDCards = Format(mDCardsVal, oPC.Configuration.LocalCurrency.FormatString)

    dblBanked = FNDBL(rsBlindCashup.Fields("Banked"))
    txtBanked = Format(dblBanked, oPC.Configuration.LocalCurrency.FormatString)

    dblPettyCashNett = FNDBL(rsBlindCashup.Fields("PettyCash"))
    txtPettyCashNett = Format(dblPettyCashNett, oPC.Configuration.LocalCurrency.FormatString)

    dblRetained = FNDBL(rsBlindCashup.Fields("CashRetained"))
    txtRetained = Format(dblRetained, oPC.Configuration.LocalCurrency.FormatString)

    dblReturned = FNDBL(rsBlindCashup.Fields("CashReturned"))
    txtReturned = Format(dblReturned, oPC.Configuration.LocalCurrency.FormatString)

    strGroupExplanation = FNS(rsBlindCashup.Fields("CU_Explanation"))
    If strGroupExplanation > "" Then
        ar = Split(strGroupExplanation, vbTab)
        strExplanation = ar(0)
        If UBound(ar) > 0 Then strCashExplanation = FNS(ar(1))
        If UBound(ar) > 1 Then strCardExplanation = FNS(ar(2))
        If UBound(ar) > 2 Then strChequeExplanation = FNS(ar(3))
        If UBound(ar) > 3 Then strDepositExplanation = FNS(ar(4))
        If UBound(ar) > 4 Then strVoucherRedeemedExplanation = FNS(ar(5))
    End If
    txtExplanation = strExplanation
    txtCashExplanation = strCashExplanation
    txtChequeExplanation = strChequeExplanation
    txtCardExplanation = strCardExplanation
    txtDepositExplanation = strDepositExplanation
    txtVoucherExplanation = strVoucherRedeemedExplanation
    
    qtyN200 = FNN(rsBlindCashup.Fields("qtyN200"))
    txtN200 = Format(qtyN200, "###,##0")
    lblN200.Caption = Format(qtyN200 * 200, oPC.Configuration.DefaultCurrency.FormatString)
    
    qtyN100 = FNN(rsBlindCashup.Fields("qtyN100"))
    txtN100 = Format(qtyN100, "###,##0")
    lblN100.Caption = Format(qtyN100 * 100, oPC.Configuration.DefaultCurrency.FormatString)
    
    qtyN50 = FNN(rsBlindCashup.Fields("qtyN50"))
    txtN50 = Format(qtyN50, "###,##0")
    lblN50.Caption = Format(qtyN50 * 50, oPC.Configuration.DefaultCurrency.FormatString)
    
    qtyN20 = FNN(rsBlindCashup.Fields("qtyN20"))
    txtN20 = Format(qtyN20, "###,##0")
    lblN20.Caption = Format(qtyN20 * 20, oPC.Configuration.DefaultCurrency.FormatString)
    
    qtyN10 = FNN(rsBlindCashup.Fields("qtyN10"))
    txtN10 = Format(qtyN10, "###,##0")
    lblN10.Caption = Format(qtyN10 * 10, oPC.Configuration.DefaultCurrency.FormatString)
    
    qtyC500 = FNN(rsBlindCashup.Fields("qtyC500"))
    txtC500 = Format(qtyC500, "###,##0")
    lblC500.Caption = Format(qtyC500 * 5, oPC.Configuration.DefaultCurrency.FormatString)
    
    qtyC200 = FNN(rsBlindCashup.Fields("qtyC200"))
    txtC200 = Format(qtyC200, "###,##0")
    lblC200.Caption = Format(qtyC200 * 2, oPC.Configuration.DefaultCurrency.FormatString)
    
    qtyC100 = FNN(rsBlindCashup.Fields("qtyC100"))
    txtC100 = Format(qtyC100, "###,##0")
    lblC100.Caption = Format(qtyC100 * 1, oPC.Configuration.DefaultCurrency.FormatString)
    
    qtyC50 = FNN(rsBlindCashup.Fields("qtyC50"))
    txtC50 = Format(qtyC50, "###,##0")
    lblC50.Caption = Format(qtyC50 * 0.5, oPC.Configuration.DefaultCurrency.FormatString)
    
    qtyC20 = FNN(rsBlindCashup.Fields("qtyC20"))
    txtC20 = Format(qtyC20, "###,##0")
    lblC20.Caption = Format(qtyC20 * 0.2, oPC.Configuration.DefaultCurrency.FormatString)
    
    qtyC10 = FNN(rsBlindCashup.Fields("qtyC10"))
    txtC10 = Format(qtyC10, "###,##0")
    lblC10.Caption = Format(qtyC10 * 0.1, oPC.Configuration.DefaultCurrency.FormatString)
    
    qtyC5 = FNN(rsBlindCashup.Fields("qtyC5"))
    txtC5 = Format(qtyC5, "###,##0")
    lblC5.Caption = Format(qtyC5 * 0.05, oPC.Configuration.DefaultCurrency.FormatString)

    If mCapturedByStaffName > "" Then
        mStageNumber = 1
    End If
    If mIssuedByStaffName > "" Then
        mStageNumber = 2
    End If
    If mExplainedByStaffName > "" Then
        mStageNumber = 3
    End If
    
    Set oCU = New z_Cashup
    oCU.SelectSession "X", mXID
    oCU.Calculate

    Recalculate
    dblCalcCashInDrawer = oCU.TotalCashInDrawerDec + dblFloatAtStart - dblFloatAtEnd
    
    txtActualCashInDrawer = Format(dblActualCashTotal, "###,##0.00")
    txtActualCashInDrawer2 = Format(dblActualCashTotal, "###,##0.00")
    If mStageNumber > 1 Then
        txtCashFromSales = Format((oCU.TotalCashDec - oCU.TotalChangeGivenDec), "###,##0.00")
        txtCalcCashInDrawer = Format((dblCalcCashInDrawer), "###,##0.00")
    End If
    
    

    If mStageNumber > 1 Then
        ReportDifferences
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.LoadCashup", , , , "Line number", Array(Erl())
End Sub
Private Sub ReportDifferences()
    On Error GoTo errHandler
    lblTotalCash.Caption = Format(dblActualCashTotal, oPC.Configuration.LocalCurrency.FormatString)
    lblTotalCheques.Caption = Format(mChequesVal, oPC.Configuration.LocalCurrency.FormatString)
    lblTotalCreditCards.Caption = Format(mDCardsVal + mCCardsVal, oPC.Configuration.LocalCurrency.FormatString)
    lblTotalDeposits.Caption = Format(mDepositsVal, oPC.Configuration.LocalCurrency.FormatString)
    lblTotalVouchersRedeemed.Caption = Format(mVouchersVal, oPC.Configuration.LocalCurrency.FormatString)
    
    lblTOtalCashS.Caption = Format((dblCalcCashInDrawer), "###,##0.00")
    lblTotalChequesS.Caption = oCU.TotalChequesF
    lblTotalCreditCardsS.Caption = oCU.TotalCreditCardsNettF
    lblTotalDepositsS.Caption = oCU.TotalDirectDepositsF
    lblTotalVouchersRedeemedS.Caption = oCU.TotalVouchersRedeemedDecF
    
'---
    dblCashDiff = Round(dblActualCashTotal - dblCalcCashInDrawer, 2)
    bCashExplanationNeeded = (dblCashDiff <> 0)
    cmdExplainCash.Enabled = bCashExplanationNeeded And mStageNumber < 3
    
    dblChequesDiff = Round(mChequesVal - oCU.TotalChequesDec, 2)
    bChequeExplanationNeeded = (dblChequesDiff <> 0)
    cmdExplainCheque.Enabled = bChequeExplanationNeeded And mStageNumber < 3
    
    dblCardsDiff = Round(mDCardsVal + mCCardsVal - oCU.TotalCreditCardsNettDec, 2)
    bCardExplanationNeeded = (dblCardsDiff <> 0)
    cmdExplainCard.Enabled = bCardExplanationNeeded And mStageNumber < 3
    
    dblCDepositsDiff = Round(mDepositsVal - oCU.TotalDirectDepositsDec, 2)
    bDepositExplanationNeeded = (dblCDepositsDiff <> 0)
    cmdExplainDeposit.Enabled = bDepositExplanationNeeded And mStageNumber < 3
    
    dblVouchersDiff = Round(mVouchersVal - oCU.TotalVouchersRedeemedDec, 2)
    bVoucherRedeemedExplanationNeeded = (dblVouchersDiff <> 0)
    cmdExplainVoucherRedeemed.Enabled = bVoucherRedeemedExplanationNeeded And mStageNumber < 3
    
    lblDiffCash.Caption = Format(dblCashDiff, oPC.Configuration.LocalCurrency.FormatString)
    lblDiffCheques.Caption = Format(dblChequesDiff, oPC.Configuration.LocalCurrency.FormatString)
    lblDiffCreditCards.Caption = Format(dblCardsDiff, oPC.Configuration.LocalCurrency.FormatString)
    lblDiffDeposits.Caption = Format(dblCDepositsDiff, oPC.Configuration.LocalCurrency.FormatString)
    lblDiffVouchersRedeemed.Caption = Format(dblVouchersDiff, oPC.Configuration.LocalCurrency.FormatString)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.ReportDifferences"
End Sub
Private Sub ValidateExplanations()
    On Error GoTo errHandler
Dim bFormOK As Boolean
    bFormOK = (bCashExplanationNeeded = False Or Me.txtCashExplanation > "") And _
                (bChequeExplanationNeeded = False Or Me.txtChequeExplanation > "") And _
                (bCardExplanationNeeded = False Or Me.txtCardExplanation > "") And _
                (bDepositExplanationNeeded = False Or Me.txtDepositExplanation > "") And _
                (bVoucherRedeemedExplanationNeeded = False Or Me.txtVoucherExplanation > "") And _
                (dblFloatAtEnd >= 0) And _
                (dblReturned >= 0) And _
                (dblRetained >= 0) And _
                (dblBanked >= 0)
    cmdStage3.Enabled = bFormOK
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.ValidateExplanations"
End Sub
Private Sub lblHelp1_Click()
Dim s1, s2, s3, s4, s5, s6, s7, s8, s9 As String

    If Not (bCashExplanationNeeded = False Or Me.txtCashExplanation > "") Then s1 = "Cash explanation problem"
    If Not (bChequeExplanationNeeded = False Or Me.txtChequeExplanation > "") Then s2 = "Cheque explanation problem"
    If Not (bCardExplanationNeeded = False Or Me.txtCardExplanation > "") Then s3 = "Card explanation problem"
    If Not (bDepositExplanationNeeded = False Or Me.txtDepositExplanation > "") Then s4 = "Deposit explanation problem"
    If Not (bVoucherRedeemedExplanationNeeded = False Or Me.txtVoucherExplanation > "") Then s5 = "Voucher explanation problem"
    If Not (dblFloatAtEnd >= 0) Then s6 = "End float problem"
    If Not (dblReturned >= 0) Then s7 = "Returned problem"
    If Not (dblRetained >= 0) Then s8 = "Retained problem"
    If Not (dblBanked >= 0) Then s9 = "Banked problem"
    MsgBox s1 & IIf(s1 > "", vbCrLf, "") _
        & s2 & IIf(s2 > "", vbCrLf, "") _
        & s3 & IIf(s3 > "", vbCrLf, "") _
        & s4 & IIf(s4 > "", vbCrLf, "") _
        & s5 & IIf(s5 > "", vbCrLf, "") _
        & s6 & IIf(s6 > "", vbCrLf, "") _
        & s9 & IIf(s7 > "", vbCrLf, "") _
        & s8 & IIf(s8 > "", vbCrLf, "") _
        & s1 & IIf(s9 > "", vbCrLf, ""), vbInformation + vbOKOnly, "Problems"
End Sub

Private Sub SaveCashup()
    On Error GoTo errHandler
Dim oSQL As New z_SQL


    strGroupExplanation = FNS(strExplanation) & vbTab & FNS(strCashExplanation) & vbTab & FNS(strCardExplanation) & vbTab & _
                        FNS(strChequeExplanation) & vbTab & FNS(strDepositExplanation) & vbTab & _
                        FNS(strVoucherRedeemedExplanation)
    
    oSQL.SaveCashup mStageNumber, _
                    Trim(mXID), _
                    mCapturedByStaffID, _
                    mIssuedByStaffID, _
                    mExplainedByStaffID, dblFloatAtEnd, _
                    mChequesVal, mCCardsVal, _
                    mDCardsVal, mVouchersVal, _
                    mDepositsVal, qtyN200, _
                    qtyN100, qtyN50, _
                    qtyN20, qtyN10, _
                    qtyC500, qtyC200, _
                    qtyC100, qtyC50, _
                    qtyC20, qtyC10, _
                    qtyC5, sFloatBreakdownatEnd, _
                    strGroupExplanation, dblCashDiff, _
                    dblChequesDiff, dblCardsDiff, _
                    dblCDepositsDiff, dblVouchersDiff, _
                    (dblFloatAtEnd - dblFloatAtStart), dblBanked, _
                    dblPettyCashNett, dblRetained, dblReturned, oCU.SalesDec, oCU.TotalCOGSDec, oCU.TotalVouchersSoldDec, _
                    oCU.TotalDepositsReceived, oCU.TotalDepositsRedeemed, oCU.TotalDepositsrefunded, _
                    oCU.TotalCNIssuedDec, oCU.TotalWagesDec, oCU.TotalSickLeaveDec, oCU.TotalLeavePayDec
                    
 
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.SaveCashup", , , , "line number", Array(Erl())
End Sub
'
Private Sub LockFormControlsforFront()
    On Error GoTo errHandler
Dim bON As Boolean
    bON = (mStageNumber <> 0)
    If bON Then
        clr = COLOR_Lightblue
    Else
        clr = COLOR_White
    End If
    txtFloatAtEnd.Locked = bON
    txtCheques.Locked = bON
    txtVouchersRedeemed.Locked = bON
    txtDeposits.Locked = bON
    txtPettyCashNett.Locked = bON
    txtN200.Locked = bON
    txtN100.Locked = bON
    txtN50.Locked = bON
    txtN20.Locked = bON
    txtN10.Locked = bON
    txtC500.Locked = bON
    txtC200.Locked = bON
    txtC100.Locked = bON
    txtC50.Locked = bON
    txtC20.Locked = bON
    txtC10.Locked = bON
    txtC5.Locked = bON
    cmdCountNewFloat.Enabled = Not bON
    
    txtFloatAtEnd.BackColor = clr
    txtCheques.BackColor = clr
    txtVouchersRedeemed.BackColor = clr
    txtDeposits.BackColor = clr
    txtN200.BackColor = clr
    txtN100.BackColor = clr
    txtN50.BackColor = clr
    txtN20.BackColor = clr
    txtN10.BackColor = clr
    txtC500.BackColor = clr
    txtC200.BackColor = clr
    txtC100.BackColor = clr
    txtC50.BackColor = clr
    txtC20.BackColor = clr
    txtC10.BackColor = clr
    txtC5.BackColor = clr
    If mStageNumber > 2 Then
        txtRetained.Locked = True
        txtRetained.BackColor = COLOR_Lightblue
        txtReturned.Locked = True
        txtReturned.BackColor = COLOR_Lightblue
    Else
        txtRetained.Locked = False
        txtRetained.BackColor = COLOR_White
        txtReturned.Locked = False
        txtReturned.BackColor = COLOR_White
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.LockFormControlsforFront"
End Sub
'

Private Sub LockFormControlsForBack()
    On Error GoTo errHandler
Dim bON As Boolean
    bON = (mStageNumber > 1)
    If bON Then
        clr = COLOR_Lightblue
    Else
        clr = COLOR_White
    End If
    txtCCards.Locked = bON
    txtDCards.Locked = bON
    txtCCards.BackColor = clr
    txtDCards.BackColor = clr
    
    txtFloatAtEnd.Locked = bON
    txtCheques.Locked = bON
    txtVouchersRedeemed.Locked = bON
    txtDeposits.Locked = bON
    txtPettyCashNett.Locked = bON
    txtN200.Locked = bON
    txtN100.Locked = bON
    txtN50.Locked = bON
    txtN20.Locked = bON
    txtN10.Locked = bON
    txtC500.Locked = bON
    txtC200.Locked = bON
    txtC100.Locked = bON
    txtC50.Locked = bON
    txtC20.Locked = bON
    txtC10.Locked = bON
    txtC5.Locked = bON
    cmdCountNewFloat.Enabled = Not bON
    
    txtFloatAtEnd.BackColor = clr
    txtCheques.BackColor = clr
    txtVouchersRedeemed.BackColor = clr
    txtDeposits.BackColor = clr
    txtN200.BackColor = clr
    txtN100.BackColor = clr
    txtN50.BackColor = clr
    txtN20.BackColor = clr
    txtN10.BackColor = clr
    txtC500.BackColor = clr
    txtC200.BackColor = clr
    txtC100.BackColor = clr
    txtC50.BackColor = clr
    txtC20.BackColor = clr
    txtC10.BackColor = clr
    txtC5.BackColor = clr
    txtRetained.BackColor = clr
    If mStageNumber > 2 Then
        txtRetained.Locked = True
        txtRetained.BackColor = COLOR_Lightblue
        txtReturned.Locked = True
        txtReturned.BackColor = COLOR_Lightblue
    Else
        txtRetained.Locked = False
        txtRetained.BackColor = COLOR_White
        txtReturned.Locked = False
        txtReturned.BackColor = COLOR_White
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBlindCashup.LockFormControlsForBack"
End Sub
'
'
'
'
'
