VERSION 5.00
Begin VB.Form frmBudgetPreview 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Budget"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRefresh 
      Height          =   285
      Left            =   330
      Picture         =   "frmBudgetChart2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   30
      Width           =   315
   End
   Begin VB.CommandButton cmdHide 
      Height          =   285
      Left            =   975
      Picture         =   "frmBudgetChart2.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   131
      Top             =   30
      Width           =   315
   End
   Begin VB.CommandButton cmdExport 
      Height          =   285
      Left            =   660
      Picture         =   "frmBudgetChart2.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   130
      Top             =   30
      Width           =   315
   End
   Begin VB.Label lblMemo 
      Height          =   165
      Left            =   15
      TabIndex        =   172
      Top             =   3345
      Width           =   10950
   End
   Begin VB.Shape Shape4 
      Height          =   2985
      Left            =   1620
      Top             =   285
      Width           =   9345
   End
   Begin VB.Label L1_5b 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1560
      TabIndex        =   171
      Top             =   1545
      Width           =   750
   End
   Begin VB.Label L2_5b 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2340
      TabIndex        =   170
      Top             =   1545
      Width           =   750
   End
   Begin VB.Label L3_5b 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3105
      TabIndex        =   169
      Top             =   1545
      Width           =   750
   End
   Begin VB.Label L4_5b 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3870
      TabIndex        =   168
      Top             =   1545
      Width           =   750
   End
   Begin VB.Label L5_5b 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4650
      TabIndex        =   167
      Top             =   1545
      Width           =   750
   End
   Begin VB.Label L6_5b 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5430
      TabIndex        =   166
      Top             =   1545
      Width           =   750
   End
   Begin VB.Label L7_5b 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6195
      TabIndex        =   165
      Top             =   1545
      Width           =   750
   End
   Begin VB.Label L8_5b 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6975
      TabIndex        =   164
      Top             =   1545
      Width           =   750
   End
   Begin VB.Label L9_5b 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7725
      TabIndex        =   163
      Top             =   1560
      Width           =   750
   End
   Begin VB.Label L10_5b 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8520
      TabIndex        =   162
      Top             =   1545
      Width           =   750
   End
   Begin VB.Label L11_5b 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   9285
      TabIndex        =   161
      Top             =   1545
      Width           =   750
   End
   Begin VB.Label L12_5b 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   10065
      TabIndex        =   160
      Top             =   1545
      Width           =   750
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nett suppl.inv."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      TabIndex        =   159
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label L1_10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   158
      Top             =   3030
      Width           =   750
   End
   Begin VB.Label L2_10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2340
      TabIndex        =   157
      Top             =   3030
      Width           =   750
   End
   Begin VB.Label L3_10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3105
      TabIndex        =   156
      Top             =   3030
      Width           =   750
   End
   Begin VB.Label L4_10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3885
      TabIndex        =   155
      Top             =   3030
      Width           =   750
   End
   Begin VB.Label L5_10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4650
      TabIndex        =   154
      Top             =   3030
      Width           =   750
   End
   Begin VB.Label L6_10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5430
      TabIndex        =   153
      Top             =   3030
      Width           =   750
   End
   Begin VB.Label L7_10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6195
      TabIndex        =   152
      Top             =   3030
      Width           =   750
   End
   Begin VB.Label L8_10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6975
      TabIndex        =   151
      Top             =   3030
      Width           =   750
   End
   Begin VB.Label L9_10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7740
      TabIndex        =   150
      Top             =   3030
      Width           =   750
   End
   Begin VB.Label L10_10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   149
      Top             =   3030
      Width           =   750
   End
   Begin VB.Label L11_10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9285
      TabIndex        =   148
      Top             =   3030
      Width           =   750
   End
   Begin VB.Label L12_10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10065
      TabIndex        =   147
      Top             =   3030
      Width           =   750
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Subscr/Replen"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   146
      Top             =   3030
      Width           =   1515
   End
   Begin VB.Label L1_0 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   145
      Top             =   345
      Width           =   750
   End
   Begin VB.Label L2_0 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2340
      TabIndex        =   144
      Top             =   345
      Width           =   750
   End
   Begin VB.Label L3_0 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3105
      TabIndex        =   143
      Top             =   345
      Width           =   750
   End
   Begin VB.Label L4_0 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3885
      TabIndex        =   142
      Top             =   345
      Width           =   750
   End
   Begin VB.Label L5_0 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4650
      TabIndex        =   141
      Top             =   345
      Width           =   750
   End
   Begin VB.Label L6_0 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5430
      TabIndex        =   140
      Top             =   345
      Width           =   750
   End
   Begin VB.Label L7_0 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6195
      TabIndex        =   139
      Top             =   345
      Width           =   750
   End
   Begin VB.Label L8_0 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6975
      TabIndex        =   138
      Top             =   345
      Width           =   750
   End
   Begin VB.Label L9_0 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7740
      TabIndex        =   137
      Top             =   345
      Width           =   750
   End
   Begin VB.Label L10_0 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   136
      Top             =   345
      Width           =   750
   End
   Begin VB.Label L11_0 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9285
      TabIndex        =   135
      Top             =   345
      Width           =   750
   End
   Begin VB.Label L12_0 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10065
      TabIndex        =   134
      Top             =   345
      Width           =   750
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   0
      TabIndex        =   133
      Top             =   360
      Width           =   1515
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   315
      Left            =   -15
      TabIndex        =   129
      Top             =   -30
      Width           =   360
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Suppl.inv.budget"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   128
      Top             =   600
      Width           =   1515
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ret. budget"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      TabIndex        =   127
      Top             =   825
      Width           =   1515
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Act. suppl.inv."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   126
      Top             =   1080
      Width           =   1515
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Act. ret."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      TabIndex        =   125
      Top             =   1320
      Width           =   1515
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Act. orders"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   124
      Top             =   1845
      Width           =   1515
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Un-iss. orders"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   123
      Top             =   2085
      Width           =   1515
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "orders/budget"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   122
      Top             =   2325
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nett/budget"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   121
      Top             =   2565
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "4-mth avg. D/B"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   120
      Top             =   2805
      Width           =   1515
   End
   Begin VB.Label H_12 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10095
      TabIndex        =   119
      Top             =   45
      Width           =   750
   End
   Begin VB.Label H_11 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9315
      TabIndex        =   118
      Top             =   45
      Width           =   750
   End
   Begin VB.Label H_10 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8550
      TabIndex        =   117
      Top             =   45
      Width           =   750
   End
   Begin VB.Label H_9 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7770
      TabIndex        =   116
      Top             =   45
      Width           =   750
   End
   Begin VB.Label H_8 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7005
      TabIndex        =   115
      Top             =   45
      Width           =   750
   End
   Begin VB.Label H_7 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6225
      TabIndex        =   114
      Top             =   45
      Width           =   750
   End
   Begin VB.Label H_6 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5460
      TabIndex        =   113
      Top             =   45
      Width           =   750
   End
   Begin VB.Label H_5 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4680
      TabIndex        =   112
      Top             =   45
      Width           =   750
   End
   Begin VB.Label H_4 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3915
      TabIndex        =   111
      Top             =   45
      Width           =   750
   End
   Begin VB.Label H_3 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3135
      TabIndex        =   110
      Top             =   45
      Width           =   750
   End
   Begin VB.Label H_2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2370
      TabIndex        =   109
      Top             =   45
      Width           =   750
   End
   Begin VB.Label H_1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1605
      TabIndex        =   108
      Top             =   45
      Width           =   750
   End
   Begin VB.Label L12_9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10065
      TabIndex        =   107
      Top             =   2790
      Width           =   750
   End
   Begin VB.Label L12_8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10065
      TabIndex        =   106
      Top             =   2550
      Width           =   750
   End
   Begin VB.Label L12_7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10065
      TabIndex        =   105
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label L12_6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10065
      TabIndex        =   104
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label L12_5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10065
      TabIndex        =   103
      Top             =   1830
      Width           =   750
   End
   Begin VB.Label L12_4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   10065
      TabIndex        =   102
      Top             =   1305
      Width           =   750
   End
   Begin VB.Label L12_3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10065
      TabIndex        =   101
      Top             =   1065
      Width           =   750
   End
   Begin VB.Label L12_2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   10065
      TabIndex        =   100
      Top             =   825
      Width           =   750
   End
   Begin VB.Label L12_1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10065
      TabIndex        =   99
      Top             =   585
      Width           =   750
   End
   Begin VB.Label L11_9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9285
      TabIndex        =   98
      Top             =   2790
      Width           =   750
   End
   Begin VB.Label L11_8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9285
      TabIndex        =   97
      Top             =   2550
      Width           =   750
   End
   Begin VB.Label L11_7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9285
      TabIndex        =   96
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label L11_6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9285
      TabIndex        =   95
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label L11_5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9285
      TabIndex        =   94
      Top             =   1830
      Width           =   750
   End
   Begin VB.Label L11_4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   9285
      TabIndex        =   93
      Top             =   1305
      Width           =   750
   End
   Begin VB.Label L11_3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9285
      TabIndex        =   92
      Top             =   1065
      Width           =   750
   End
   Begin VB.Label L11_2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   9285
      TabIndex        =   91
      Top             =   825
      Width           =   750
   End
   Begin VB.Label L11_1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9285
      TabIndex        =   90
      Top             =   585
      Width           =   750
   End
   Begin VB.Label L10_9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   89
      Top             =   2790
      Width           =   750
   End
   Begin VB.Label L10_8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   88
      Top             =   2550
      Width           =   750
   End
   Begin VB.Label L10_7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   87
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label L10_6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   86
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label L10_5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   85
      Top             =   1830
      Width           =   750
   End
   Begin VB.Label L10_4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8520
      TabIndex        =   84
      Top             =   1305
      Width           =   750
   End
   Begin VB.Label L10_3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   83
      Top             =   1065
      Width           =   750
   End
   Begin VB.Label L10_2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8520
      TabIndex        =   82
      Top             =   825
      Width           =   750
   End
   Begin VB.Label L10_1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   81
      Top             =   585
      Width           =   750
   End
   Begin VB.Label L9_9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7740
      TabIndex        =   80
      Top             =   2790
      Width           =   750
   End
   Begin VB.Label L9_8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7740
      TabIndex        =   79
      Top             =   2550
      Width           =   750
   End
   Begin VB.Label L9_7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7740
      TabIndex        =   78
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label L9_6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7740
      TabIndex        =   77
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label L9_5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7740
      TabIndex        =   76
      Top             =   1830
      Width           =   750
   End
   Begin VB.Label L9_4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7740
      TabIndex        =   75
      Top             =   1305
      Width           =   750
   End
   Begin VB.Label L9_3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7740
      TabIndex        =   74
      Top             =   1065
      Width           =   750
   End
   Begin VB.Label L9_2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7740
      TabIndex        =   73
      Top             =   825
      Width           =   750
   End
   Begin VB.Label L9_1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7740
      TabIndex        =   72
      Top             =   585
      Width           =   750
   End
   Begin VB.Label L8_9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6975
      TabIndex        =   71
      Top             =   2790
      Width           =   750
   End
   Begin VB.Label L8_8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6975
      TabIndex        =   70
      Top             =   2550
      Width           =   750
   End
   Begin VB.Label L8_7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6975
      TabIndex        =   69
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label L8_6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6975
      TabIndex        =   68
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label L8_5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6975
      TabIndex        =   67
      Top             =   1830
      Width           =   750
   End
   Begin VB.Label L8_4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6975
      TabIndex        =   66
      Top             =   1305
      Width           =   750
   End
   Begin VB.Label L8_3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6975
      TabIndex        =   65
      Top             =   1065
      Width           =   750
   End
   Begin VB.Label L8_2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6975
      TabIndex        =   64
      Top             =   825
      Width           =   750
   End
   Begin VB.Label L8_1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6975
      TabIndex        =   63
      Top             =   585
      Width           =   750
   End
   Begin VB.Label L7_9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6195
      TabIndex        =   62
      Top             =   2790
      Width           =   750
   End
   Begin VB.Label L7_8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6195
      TabIndex        =   61
      Top             =   2550
      Width           =   750
   End
   Begin VB.Label L7_7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6195
      TabIndex        =   60
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label L7_6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6195
      TabIndex        =   59
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label L7_5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6195
      TabIndex        =   58
      Top             =   1830
      Width           =   750
   End
   Begin VB.Label L7_4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6195
      TabIndex        =   57
      Top             =   1305
      Width           =   750
   End
   Begin VB.Label L7_3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6195
      TabIndex        =   56
      Top             =   1065
      Width           =   750
   End
   Begin VB.Label L7_2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6195
      TabIndex        =   55
      Top             =   825
      Width           =   750
   End
   Begin VB.Label L7_1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6195
      TabIndex        =   54
      Top             =   585
      Width           =   750
   End
   Begin VB.Label L6_9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5430
      TabIndex        =   53
      Top             =   2790
      Width           =   750
   End
   Begin VB.Label L6_8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5430
      TabIndex        =   52
      Top             =   2550
      Width           =   750
   End
   Begin VB.Label L6_7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5430
      TabIndex        =   51
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label L6_6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5430
      TabIndex        =   50
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label L6_5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5430
      TabIndex        =   49
      Top             =   1830
      Width           =   750
   End
   Begin VB.Label L6_4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5430
      TabIndex        =   48
      Top             =   1305
      Width           =   750
   End
   Begin VB.Label L6_3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5430
      TabIndex        =   47
      Top             =   1065
      Width           =   750
   End
   Begin VB.Label L6_2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5430
      TabIndex        =   46
      Top             =   825
      Width           =   750
   End
   Begin VB.Label L6_1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5430
      TabIndex        =   45
      Top             =   585
      Width           =   750
   End
   Begin VB.Label L5_9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4650
      TabIndex        =   44
      Top             =   2790
      Width           =   750
   End
   Begin VB.Label L5_8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4650
      TabIndex        =   43
      Top             =   2550
      Width           =   750
   End
   Begin VB.Label L5_7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4650
      TabIndex        =   42
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label L5_6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4650
      TabIndex        =   41
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label L5_5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4650
      TabIndex        =   40
      Top             =   1830
      Width           =   750
   End
   Begin VB.Label L5_4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4650
      TabIndex        =   39
      Top             =   1305
      Width           =   750
   End
   Begin VB.Label L5_3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4650
      TabIndex        =   38
      Top             =   1065
      Width           =   750
   End
   Begin VB.Label L5_2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4650
      TabIndex        =   37
      Top             =   825
      Width           =   750
   End
   Begin VB.Label L5_1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4650
      TabIndex        =   36
      Top             =   585
      Width           =   750
   End
   Begin VB.Label L4_9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3885
      TabIndex        =   35
      Top             =   2790
      Width           =   750
   End
   Begin VB.Label L4_8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3885
      TabIndex        =   34
      Top             =   2550
      Width           =   750
   End
   Begin VB.Label L4_7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3885
      TabIndex        =   33
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label L4_6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3885
      TabIndex        =   32
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label L4_5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3885
      TabIndex        =   31
      Top             =   1830
      Width           =   750
   End
   Begin VB.Label L4_4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3885
      TabIndex        =   30
      Top             =   1305
      Width           =   750
   End
   Begin VB.Label L4_3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3885
      TabIndex        =   29
      Top             =   1065
      Width           =   750
   End
   Begin VB.Label L4_2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3885
      TabIndex        =   28
      Top             =   825
      Width           =   750
   End
   Begin VB.Label L4_1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3885
      TabIndex        =   27
      Top             =   585
      Width           =   750
   End
   Begin VB.Label L3_9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3105
      TabIndex        =   26
      Top             =   2790
      Width           =   750
   End
   Begin VB.Label L3_8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3105
      TabIndex        =   25
      Top             =   2550
      Width           =   750
   End
   Begin VB.Label L3_7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3105
      TabIndex        =   24
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label L3_6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3105
      TabIndex        =   23
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label L3_5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3105
      TabIndex        =   22
      Top             =   1830
      Width           =   750
   End
   Begin VB.Label L3_4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3105
      TabIndex        =   21
      Top             =   1305
      Width           =   750
   End
   Begin VB.Label L3_3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3105
      TabIndex        =   20
      Top             =   1065
      Width           =   750
   End
   Begin VB.Label L3_2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3105
      TabIndex        =   19
      Top             =   825
      Width           =   750
   End
   Begin VB.Label L3_1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3105
      TabIndex        =   18
      Top             =   585
      Width           =   750
   End
   Begin VB.Label L2_9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2340
      TabIndex        =   17
      Top             =   2790
      Width           =   750
   End
   Begin VB.Label L2_8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2340
      TabIndex        =   16
      Top             =   2550
      Width           =   750
   End
   Begin VB.Label L2_7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2340
      TabIndex        =   15
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label L2_6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2340
      TabIndex        =   14
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label L2_5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2340
      TabIndex        =   13
      Top             =   1830
      Width           =   750
   End
   Begin VB.Label L2_4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2340
      TabIndex        =   12
      Top             =   1305
      Width           =   750
   End
   Begin VB.Label L2_3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2340
      TabIndex        =   11
      Top             =   1065
      Width           =   750
   End
   Begin VB.Label L2_2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2340
      TabIndex        =   10
      Top             =   825
      Width           =   750
   End
   Begin VB.Label L2_1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2340
      TabIndex        =   9
      Top             =   585
      Width           =   750
   End
   Begin VB.Label L1_9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   8
      Top             =   2790
      Width           =   750
   End
   Begin VB.Label L1_8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   7
      Top             =   2550
      Width           =   750
   End
   Begin VB.Label L1_7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   6
      Top             =   2310
      Width           =   750
   End
   Begin VB.Label L1_6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   5
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label L1_5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   4
      Top             =   1830
      Width           =   750
   End
   Begin VB.Label L1_4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1545
      TabIndex        =   3
      Top             =   1305
      Width           =   750
   End
   Begin VB.Label L1_3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   2
      Top             =   1065
      Width           =   750
   End
   Begin VB.Label L1_2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1560
      TabIndex        =   1
      Top             =   825
      Width           =   750
   End
   Begin VB.Label L1_1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "120,000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   0
      Top             =   585
      Width           =   750
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   -45
      Top             =   1500
      Width           =   11010
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   510
      Left            =   -15
      Top             =   555
      Width           =   10980
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      Height          =   285
      Left            =   -30
      Top             =   2520
      Width           =   10995
   End
End
Attribute VB_Name = "frmBudgetPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
    Dim excApp As Object
    Dim excWb As Object
    Dim excWs As Object
Dim oSheet As Object
Public Sub LoadData()
    On Error GoTo errHandler


    Set rs = oPC.rsBudgetData
    H_1 = Format(FND(rs.fields("H01")), "MM-YY")
    H_2 = Format(FND(rs.fields("H02")), "MM-YY")
    H_3 = Format(FND(rs.fields("H03")), "MM-YY")
    H_4 = Format(FND(rs.fields("H04")), "MM-YY")
    H_5 = Format(FND(rs.fields("H05")), "MM-YY")
    H_6 = Format(FND(rs.fields("H06")), "MM-YY")
    H_7 = Format(FND(rs.fields("H07")), "MM-YY")
    H_8 = Format(FND(rs.fields("H08")), "MM-YY")
    H_9 = Format(FND(rs.fields("H09")), "MM-YY")
    H_10 = Format(FND(rs.fields("H10")), "MM-YY")
    H_11 = Format(FND(rs.fields("H11")), "MM-YY")
    H_12 = Format(FND(rs.fields("H12")), "MM-YY")
    
    L1_0 = Format(FNN(rs.fields("M01_0")), "###,##0")
    L2_0 = Format(FNN(rs.fields("M02_0")), "###,##0")
    L3_0 = Format(FNN(rs.fields("M03_0")), "###,##0")
    L4_0 = Format(FNN(rs.fields("M04_0")), "###,##0")
    L5_0 = Format(FNN(rs.fields("M05_0")), "###,##0")
    L6_0 = Format(FNN(rs.fields("M06_0")), "###,##0")
    L7_0 = Format(FNN(rs.fields("M07_0")), "###,##0")
    L8_0 = Format(FNN(rs.fields("M08_0")), "###,##0")
    L9_0 = Format(FNN(rs.fields("M09_0")), "###,##0")
    L10_0 = Format(FNN(rs.fields("M10_0")), "###,##0")
    L11_0 = Format(FNN(rs.fields("M11_0")), "###,##0")
    L12_0 = Format(FNN(rs.fields("M12_0")), "###,##0")
    
    L1_1 = Format(FNN(rs.fields("M01_1")), "###,##0")
    L2_1 = Format(FNN(rs.fields("M02_1")), "###,##0")
    L3_1 = Format(FNN(rs.fields("M03_1")), "###,##0")
    L4_1 = Format(FNN(rs.fields("M04_1")), "###,##0")
    L5_1 = Format(FNN(rs.fields("M05_1")), "###,##0")
    L6_1 = Format(FNN(rs.fields("M06_1")), "###,##0")
    L7_1 = Format(FNN(rs.fields("M07_1")), "###,##0")
    L8_1 = Format(FNN(rs.fields("M08_1")), "###,##0")
    L9_1 = Format(FNN(rs.fields("M09_1")), "###,##0")
    L10_1 = Format(FNN(rs.fields("M10_1")), "###,##0")
    L11_1 = Format(FNN(rs.fields("M11_1")), "###,##0")
    L12_1 = Format(FNN(rs.fields("M12_1")), "###,##0")
    
    L1_2 = Format(FNN(rs.fields("M01_2")), "###,##0")
    L2_2 = Format(FNN(rs.fields("M02_2")), "###,##0")
    L3_2 = Format(FNN(rs.fields("M03_2")), "###,##0")
    L4_2 = Format(FNN(rs.fields("M04_2")), "###,##0")
    L5_2 = Format(FNN(rs.fields("M05_2")), "###,##0")
    L6_2 = Format(FNN(rs.fields("M06_2")), "###,##0")
    L7_2 = Format(FNN(rs.fields("M07_2")), "###,##0")
    L8_2 = Format(FNN(rs.fields("M08_2")), "###,##0")
    L9_2 = Format(FNN(rs.fields("M09_2")), "###,##0")
    L10_2 = Format(FNN(rs.fields("M10_2")), "###,##0")
    L11_2 = Format(FNN(rs.fields("M11_2")), "###,##0")
    L12_2 = Format(FNN(rs.fields("M12_2")), "###,##0")
    
    L1_3 = Format(FNN(rs.fields("M01_3")), "###,##0")
    L2_3 = Format(FNN(rs.fields("M02_3")), "###,##0")
    L3_3 = Format(FNN(rs.fields("M03_3")), "###,##0")
    L4_3 = Format(FNN(rs.fields("M04_3")), "###,##0")
    L5_3 = Format(FNN(rs.fields("M05_3")), "###,##0")
    L6_3 = Format(FNN(rs.fields("M06_3")), "###,##0")
    L7_3 = Format(FNN(rs.fields("M07_3")), "###,##0")
    L8_3 = Format(FNN(rs.fields("M08_3")), "###,##0")
    L9_3 = Format(FNN(rs.fields("M09_3")), "###,##0")
    L10_3 = Format(FNN(rs.fields("M10_3")), "###,##0")
    L11_3 = Format(FNN(rs.fields("M11_3")), "###,##0")
    L12_3 = Format(FNN(rs.fields("M12_3")), "###,##0")
    
    L1_4 = Format(FNN(rs.fields("M01_4")), "###,##0")
    L2_4 = Format(FNN(rs.fields("M02_4")), "###,##0")
    L3_4 = Format(FNN(rs.fields("M03_4")), "###,##0")
    L4_4 = Format(FNN(rs.fields("M04_4")), "###,##0")
    L5_4 = Format(FNN(rs.fields("M05_4")), "###,##0")
    L6_4 = Format(FNN(rs.fields("M06_4")), "###,##0")
    L7_4 = Format(FNN(rs.fields("M07_4")), "###,##0")
    L8_4 = Format(FNN(rs.fields("M08_4")), "###,##0")
    L9_4 = Format(FNN(rs.fields("M09_4")), "###,##0")
    L10_4 = Format(FNN(rs.fields("M10_4")), "###,##0")
    L11_4 = Format(FNN(rs.fields("M11_4")), "###,##0")
    L12_4 = Format(FNN(rs.fields("M12_4")), "###,##0")
    
    L1_5 = Format(FNN(rs.fields("M01_5")), "###,##0")
    L2_5 = Format(FNN(rs.fields("M02_5")), "###,##0")
    L3_5 = Format(FNN(rs.fields("M03_5")), "###,##0")
    L4_5 = Format(FNN(rs.fields("M04_5")), "###,##0")
    L5_5 = Format(FNN(rs.fields("M05_5")), "###,##0")
    L6_5 = Format(FNN(rs.fields("M06_5")), "###,##0")
    L7_5 = Format(FNN(rs.fields("M07_5")), "###,##0")
    L8_5 = Format(FNN(rs.fields("M08_5")), "###,##0")
    L9_5 = Format(FNN(rs.fields("M09_5")), "###,##0")
    L10_5 = Format(FNN(rs.fields("M10_5")), "###,##0")
    L11_5 = Format(FNN(rs.fields("M11_5")), "###,##0")
    L12_5 = Format(FNN(rs.fields("M12_5")), "###,##0")
    
    L1_5b = Format(FNN(rs.fields("M01_5b")), "###,##0")
    L2_5b = Format(FNN(rs.fields("M02_5b")), "###,##0")
    L3_5b = Format(FNN(rs.fields("M03_5b")), "###,##0")
    L4_5b = Format(FNN(rs.fields("M04_5b")), "###,##0")
    L5_5b = Format(FNN(rs.fields("M05_5b")), "###,##0")
    L6_5b = Format(FNN(rs.fields("M06_5b")), "###,##0")
    L7_5b = Format(FNN(rs.fields("M07_5b")), "###,##0")
    L8_5b = Format(FNN(rs.fields("M08_5b")), "###,##0")
    L9_5b = Format(FNN(rs.fields("M09_5b")), "###,##0")
    L10_5b = Format(FNN(rs.fields("M10_5b")), "###,##0")
    L11_5b = Format(FNN(rs.fields("M11_5b")), "###,##0")
    L12_5b = Format(FNN(rs.fields("M12_5b")), "###,##0")
    
    
    L1_6 = Format(FNN(rs.fields("M01_6")), "###,##0")
    L2_6 = Format(FNN(rs.fields("M02_6")), "###,##0")
    L3_6 = Format(FNN(rs.fields("M03_6")), "###,##0")
    L4_6 = Format(FNN(rs.fields("M04_6")), "###,##0")
    L5_6 = Format(FNN(rs.fields("M05_6")), "###,##0")
    L6_6 = Format(FNN(rs.fields("M06_6")), "###,##0")
    L7_6 = Format(FNN(rs.fields("M07_6")), "###,##0")
    L8_6 = Format(FNN(rs.fields("M08_6")), "###,##0")
    L9_6 = Format(FNN(rs.fields("M09_6")), "###,##0")
    L10_6 = Format(FNN(rs.fields("M10_6")), "###,##0")
    L11_6 = Format(FNN(rs.fields("M11_6")), "###,##0")
    L12_6 = Format(FNN(rs.fields("M12_6")), "###,##0")
    
    L1_7 = Format(FNDBL(rs.fields("M01_7")), "##0%")
    L2_7 = Format(FNDBL(rs.fields("M02_7")), "##0%")
    L3_7 = Format(FNDBL(rs.fields("M03_7")), "##0%")
    L4_7 = Format(FNDBL(rs.fields("M04_7")), "##0%")
    L5_7 = Format(FNDBL(rs.fields("M05_7")), "##0%")
    L6_7 = Format(FNDBL(rs.fields("M06_7")), "##0%")
    L7_7 = Format(FNDBL(rs.fields("M07_7")), "##0%")
    L8_7 = Format(FNDBL(rs.fields("M08_7")), "##0%")
    L9_7 = Format(FNDBL(rs.fields("M09_7")), "##0%")
    L10_7 = Format(FNDBL(rs.fields("M10_7")), "##0%")
    L11_7 = Format(FNDBL(rs.fields("M11_7")), "##0%")
    L12_7 = Format(FNDBL(rs.fields("M12_7")), "##0%")
    
    L1_8 = Format(FNDBL(rs.fields("M01_8")), "##0%")
    L2_8 = Format(FNDBL(rs.fields("M02_8")), "###0%")
    L3_8 = Format(FNDBL(rs.fields("M03_8")), "##0%")
    L4_8 = Format(FNDBL(rs.fields("M04_8")), "##0%")
    L5_8 = Format(FNDBL(rs.fields("M05_8")), "##0%")
    L6_8 = Format(FNDBL(rs.fields("M06_8")), "##0%")
    L7_8 = Format(FNDBL(rs.fields("M07_8")), "##0%")
    L8_8 = Format(FNDBL(rs.fields("M08_8")), "##0%")
    L9_8 = Format(FNDBL(rs.fields("M09_8")), "##0%")
    L10_8 = Format(FNDBL(rs.fields("M10_8")), "##0%")
    L11_8 = Format(FNDBL(rs.fields("M11_8")), "##0%")
    L12_8 = Format(FNDBL(rs.fields("M12_8")), "##0%")
    
    L1_9 = Format(FNDBL(rs.fields("M01_9")), "##0%")
    L2_9 = Format(FNDBL(rs.fields("M02_9")), "##0%")
    L3_9 = Format(FNDBL(rs.fields("M03_9")), "##0%")
    L4_9 = Format(FNDBL(rs.fields("M04_9")), "##0%")
    L5_9 = Format(FNDBL(rs.fields("M05_9")), "##0%")
    L6_9 = Format(FNDBL(rs.fields("M06_9")), "##0%")
    L7_9 = Format(FNDBL(rs.fields("M07_9")), "##0%")
    L8_9 = Format(FNDBL(rs.fields("M08_9")), "##0%")
    L9_9 = Format(FNDBL(rs.fields("M09_9")), "##0%")
    L10_9 = Format(FNDBL(rs.fields("M10_9")), "##0%")
    L11_9 = Format(FNDBL(rs.fields("M11_9")), "##0%")
    L12_9 = Format(FNDBL(rs.fields("M12_9")), "##0%")
    
    L1_10 = Format(FNDBL(rs.fields("M01_10")), "##0%")
    L2_10 = Format(FNDBL(rs.fields("M02_10")), "##0%")
    L3_10 = Format(FNDBL(rs.fields("M03_10")), "##0%")
    L4_10 = Format(FNDBL(rs.fields("M04_10")), "##0%")
    L5_10 = Format(FNDBL(rs.fields("M05_10")), "##0%")
    L6_10 = Format(FNDBL(rs.fields("M06_10")), "##0%")
    L7_10 = Format(FNDBL(rs.fields("M07_10")), "##0%")
    L8_10 = Format(FNDBL(rs.fields("M08_10")), "##0%")
    L9_10 = Format(FNDBL(rs.fields("M09_10")), "##0%")
    L10_10 = Format(FNDBL(rs.fields("M10_10")), "##0%")
    L11_10 = Format(FNDBL(rs.fields("M11_10")), "##0%")
    L12_10 = Format(FNDBL(rs.fields("M12_10")), "##0%")
    lblMemo.Caption = "Data was last calculated on " & Format(FND(rs.fields("B_LastCalculated")), "DD-MM-YY Hh:Nn") & ". Note suppliers invoices value include unissued documents."
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBudgetPreview.LoadData"
End Sub


Private Sub cmdHide_Click()
    Me.Height = 3920
    Me.Width = 11180
    Me.TOP = Forms(0).Height - Forms(0).TOP - Me.Height - 1600
    Me.Left = 0

    Me.Hide
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo errHandler
    If MsgBox("This refresh action may take a few minutes. Wait for it to complete before attempting further actions in Papyrus." & vbCrLf & "Click Cancel button to skip refresh action.", vbOKCancel, "Confirm refresh") = vbCancel Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    oPC.ReloadBudget
    LoadData
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBudgetPreview.cmdRefresh_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    Me.Height = 3920
    Me.Width = 11180
    Me.TOP = Forms(0).Height - Forms(0).TOP - Me.Height - 1600
    Me.Left = 0
    LoadData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBudgetPreview.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
  '  MsgBox "Budget 1"
    Me.Height = 3920
    Me.Width = 11180
    Me.TOP = Forms(0).Height - Forms(0).TOP - Me.Height - 1600
    Me.Left = 0
'MsgBox "Budget 2"
    LoadData
'MsgBox "Budget 3"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBudgetPreview.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub lblHelp_Click()
    On Error GoTo errHandler
Dim f As New frmHelpGen
Dim s As String

s = "Notes on understanding the budget report" & vbCrLf
s = s & "========================================" & vbCrLf & vbCrLf
s = s & "1.  The columns are headed by dates and they refer to the Expected delivery date (ETA) i.r.o. purchase orders" & vbCrLf
s = s & "    and to supplier invoice dates in respect of deliveries." & vbCrLf & vbCrLf
s = s & "2. The first band represents actual sales made in that month." & vbCrLf & vbCrLf
s = s & "3. The blue band represents the budgetted P.O.s receivable in that month and the returns to be effected in that month." & vbCrLf & vbCrLf
s = s & "4. The band below the blue represents suppliers invoices received with dates in that month and actual returns removed from stock in that month." & vbCrLf & vbCrLf
s = s & "5. The yellow band represents the difference (nett effect of the above). The cost incurred in that month." & vbCrLf & vbCrLf
s = s & "6. The plain band below the yellow band represents the actual orders placed with delivery dates(ETA) in that month. " & vbCrLf
s = s & "   The second row shows any unissued orders. The third row shows the actual orders value compared to the budgetted value as a percentage." & vbCrLf & vbCrLf
s = s & "7. The red band shows the (deliveries - returns) against the budget." & vbCrLf & vbCrLf
s = s & "8. The last band shows the average of the red band over the last 4 months." & vbCrLf
s = s & "   Stock is not always delivered in the actual month that it is expected so a smoothed average is necessary." & vbCrLf
s = s & "   The last line shows the spread between subscription and replenishment orders."
    f.component s, "Budget help", 13000, 6500
    f.Show
'MsgBox "Still coming"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBudgetPreview.lblHelp_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdExport_Click()
    On Error GoTo errHandler

    If oPC.UsesExcel Then
        ExportToExcel
    Else
        ExportToOO
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBudgetPreview.cmdExport_Click", , EA_NORERAISE
    HandleError
End Sub
Public Sub ExportToExcel()
    On Error GoTo errHandler
        Screen.MousePointer = vbHourglass

        Set excApp = CreateObject("excel.application")
        Set excWb = excApp.Workbooks.open(oPC.SharedFolderRoot & "\Templates\budget.xls")
       ' Set excWb = excApp.Workbooks.Open(oPC.LocalFolder & "\Templates\budget.xls")
        Set excWs = excWb.Sheets.Item(1)
        excWs.Application.Visible = True
        
        excWs.Cells(2, 1) = "Prepared: " & Format(Now(), "DD-MMM-YYYY HH:NN")
        excWs.Cells(5, 3) = Format(FND(rs.fields("H01")), "DD-MMM-YYYY")
        excWs.Cells(5, 4) = Format(FND(rs.fields("H02")), "DD-MMM-YYYY")
        excWs.Cells(5, 5) = Format(FND(rs.fields("H03")), "DD-MMM-YYYY")
        excWs.Cells(5, 6) = Format(FND(rs.fields("H04")), "DD-MMM-YYYY")
        excWs.Cells(5, 7) = Format(FND(rs.fields("H05")), "DD-MMM-YYYY")
        excWs.Cells(5, 8) = Format(FND(rs.fields("H06")), "DD-MMM-YYYY")
        excWs.Cells(5, 9) = Format(FND(rs.fields("H07")), "DD-MMM-YYYY")
        excWs.Cells(5, 10) = Format(FND(rs.fields("H08")), "DD-MMM-YYYY")
        excWs.Cells(5, 11) = Format(FND(rs.fields("H09")), "DD-MMM-YYYY")
        excWs.Cells(5, 12) = Format(FND(rs.fields("H10")), "DD-MMM-YYYY")
        excWs.Cells(5, 13) = Format(FND(rs.fields("H11")), "DD-MMM-YYYY")
        excWs.Cells(5, 14) = Format(FND(rs.fields("H12")), "DD-MMM-YYYY")

'Sales
        excWs.Cells(6, 3) = Format(FNN(rs.fields("M01_0")), "###,##0")
        excWs.Cells(6, 4) = Format(FNN(rs.fields("M02_0")), "###,##0")
        excWs.Cells(6, 5) = Format(FNN(rs.fields("M03_0")), "###,##0")
        excWs.Cells(6, 6) = Format(FNN(rs.fields("M04_0")), "###,##0")
        excWs.Cells(6, 7) = Format(FNN(rs.fields("M05_0")), "###,##0")
        excWs.Cells(6, 8) = Format(FNN(rs.fields("M06_0")), "###,##0")
        excWs.Cells(6, 9) = Format(FNN(rs.fields("M07_0")), "###,##0")
        excWs.Cells(6, 10) = Format(FNN(rs.fields("M08_0")), "###,##0")
        excWs.Cells(6, 11) = Format(FNN(rs.fields("M09_0")), "###,##0")
        excWs.Cells(6, 12) = Format(FNN(rs.fields("M10_0")), "###,##0")
        excWs.Cells(6, 13) = Format(FNN(rs.fields("M11_0")), "###,##0")
        excWs.Cells(6, 14) = Format(FNN(rs.fields("M12_0")), "###,##0")
        
        
        excWs.Cells(7, 3) = Format(FNN(rs.fields("M01_1")), "###,##0")
        excWs.Cells(7, 4) = Format(FNN(rs.fields("M02_1")), "###,##0")
        excWs.Cells(7, 5) = Format(FNN(rs.fields("M03_1")), "###,##0")
        excWs.Cells(7, 6) = Format(FNN(rs.fields("M04_1")), "###,##0")
        excWs.Cells(7, 7) = Format(FNN(rs.fields("M05_1")), "###,##0")
        excWs.Cells(7, 8) = Format(FNN(rs.fields("M06_1")), "###,##0")
        excWs.Cells(7, 9) = Format(FNN(rs.fields("M07_1")), "###,##0")
        excWs.Cells(7, 10) = Format(FNN(rs.fields("M08_1")), "###,##0")
        excWs.Cells(7, 11) = Format(FNN(rs.fields("M09_1")), "###,##0")
        excWs.Cells(7, 12) = Format(FNN(rs.fields("M10_1")), "###,##0")
        excWs.Cells(7, 13) = Format(FNN(rs.fields("M11_1")), "###,##0")
        excWs.Cells(7, 14) = Format(FNN(rs.fields("M12_1")), "###,##0")
    
        excWs.Cells(8, 3) = Format(FNN(rs.fields("M01_2")), "###,##0")
        excWs.Cells(8, 4) = Format(FNN(rs.fields("M02_2")), "###,##0")
        excWs.Cells(8, 5) = Format(FNN(rs.fields("M03_2")), "###,##0")
        excWs.Cells(8, 6) = Format(FNN(rs.fields("M04_2")), "###,##0")
        excWs.Cells(8, 7) = Format(FNN(rs.fields("M05_2")), "###,##0")
        excWs.Cells(8, 8) = Format(FNN(rs.fields("M06_2")), "###,##0")
        excWs.Cells(8, 9) = Format(FNN(rs.fields("M07_2")), "###,##0")
        excWs.Cells(8, 10) = Format(FNN(rs.fields("M08_2")), "###,##0")
        excWs.Cells(8, 11) = Format(FNN(rs.fields("M09_2")), "###,##0")
        excWs.Cells(8, 12) = Format(FNN(rs.fields("M10_2")), "###,##0")
        excWs.Cells(8, 13) = Format(FNN(rs.fields("M11_2")), "###,##0")
        excWs.Cells(8, 14) = Format(FNN(rs.fields("M12_2")), "###,##0")
    
        excWs.Cells(9, 3) = Format(FNN(rs.fields("M01_3")), "###,##0")
        excWs.Cells(9, 4) = Format(FNN(rs.fields("M02_3")), "###,##0")
        excWs.Cells(9, 5) = Format(FNN(rs.fields("M03_3")), "###,##0")
        excWs.Cells(9, 6) = Format(FNN(rs.fields("M04_3")), "###,##0")
        excWs.Cells(9, 7) = Format(FNN(rs.fields("M05_3")), "###,##0")
        excWs.Cells(9, 8) = Format(FNN(rs.fields("M06_3")), "###,##0")
        excWs.Cells(9, 9) = Format(FNN(rs.fields("M07_3")), "###,##0")
        excWs.Cells(9, 10) = Format(FNN(rs.fields("M08_3")), "###,##0")
        excWs.Cells(9, 11) = Format(FNN(rs.fields("M09_3")), "###,##0")
        excWs.Cells(9, 12) = Format(FNN(rs.fields("M10_3")), "###,##0")
        excWs.Cells(9, 13) = Format(FNN(rs.fields("M11_3")), "###,##0")
        excWs.Cells(9, 14) = Format(FNN(rs.fields("M12_3")), "###,##0")
    
        excWs.Cells(10, 3) = Format(FNN(rs.fields("M01_4")), "###,##0")
        excWs.Cells(10, 4) = Format(FNN(rs.fields("M02_4")), "###,##0")
        excWs.Cells(10, 5) = Format(FNN(rs.fields("M03_4")), "###,##0")
        excWs.Cells(10, 6) = Format(FNN(rs.fields("M04_4")), "###,##0")
        excWs.Cells(10, 7) = Format(FNN(rs.fields("M05_4")), "###,##0")
        excWs.Cells(10, 8) = Format(FNN(rs.fields("M06_4")), "###,##0")
        excWs.Cells(10, 9) = Format(FNN(rs.fields("M07_4")), "###,##0")
        excWs.Cells(10, 10) = Format(FNN(rs.fields("M08_4")), "###,##0")
        excWs.Cells(10, 11) = Format(FNN(rs.fields("M09_4")), "###,##0")
        excWs.Cells(10, 12) = Format(FNN(rs.fields("M10_4")), "###,##0")
        excWs.Cells(10, 13) = Format(FNN(rs.fields("M11_4")), "###,##0")
        excWs.Cells(10, 14) = Format(FNN(rs.fields("M12_4")), "###,##0")
    
'Insert nett suppliers invoices here
        excWs.Cells(11, 3) = Format(FNN(rs.fields("M01_5b")), "###,##0")
        excWs.Cells(11, 4) = Format(FNN(rs.fields("M02_5b")), "###,##0")
        excWs.Cells(11, 5) = Format(FNN(rs.fields("M03_5b")), "###,##0")
        excWs.Cells(11, 6) = Format(FNN(rs.fields("M04_5b")), "###,##0")
        excWs.Cells(11, 7) = Format(FNN(rs.fields("M05_5b")), "###,##0")
        excWs.Cells(11, 8) = Format(FNN(rs.fields("M06_5b")), "###,##0")
        excWs.Cells(11, 9) = Format(FNN(rs.fields("M07_5b")), "###,##0")
        excWs.Cells(11, 10) = Format(FNN(rs.fields("M08_5b")), "###,##0")
        excWs.Cells(11, 11) = Format(FNN(rs.fields("M09_5b")), "###,##0")
        excWs.Cells(11, 12) = Format(FNN(rs.fields("M10_5b")), "###,##0")
        excWs.Cells(11, 13) = Format(FNN(rs.fields("M11_5b")), "###,##0")
        excWs.Cells(11, 14) = Format(FNN(rs.fields("M12_5b")), "###,##0")
    
        excWs.Cells(12, 3) = Format(FNN(rs.fields("M01_5")), "###,##0")
        excWs.Cells(12, 4) = Format(FNN(rs.fields("M02_5")), "###,##0")
        excWs.Cells(12, 5) = Format(FNN(rs.fields("M03_5")), "###,##0")
        excWs.Cells(12, 6) = Format(FNN(rs.fields("M04_5")), "###,##0")
        excWs.Cells(12, 7) = Format(FNN(rs.fields("M05_5")), "###,##0")
        excWs.Cells(12, 8) = Format(FNN(rs.fields("M06_5")), "###,##0")
        excWs.Cells(12, 9) = Format(FNN(rs.fields("M07_5")), "###,##0")
        excWs.Cells(12, 10) = Format(FNN(rs.fields("M08_5")), "###,##0")
        excWs.Cells(12, 11) = Format(FNN(rs.fields("M09_5")), "###,##0")
        excWs.Cells(12, 12) = Format(FNN(rs.fields("M10_5")), "###,##0")
        excWs.Cells(12, 13) = Format(FNN(rs.fields("M11_5")), "###,##0")
        excWs.Cells(12, 14) = Format(FNN(rs.fields("M12_5")), "###,##0")
    
        excWs.Cells(13, 3) = Format(FNN(rs.fields("M01_6")), "###,##0")
        excWs.Cells(13, 4) = Format(FNN(rs.fields("M02_6")), "###,##0")
        excWs.Cells(13, 5) = Format(FNN(rs.fields("M03_6")), "###,##0")
        excWs.Cells(13, 6) = Format(FNN(rs.fields("M04_6")), "###,##0")
        excWs.Cells(13, 7) = Format(FNN(rs.fields("M05_6")), "###,##0")
        excWs.Cells(13, 8) = Format(FNN(rs.fields("M06_6")), "###,##0")
        excWs.Cells(13, 9) = Format(FNN(rs.fields("M07_6")), "###,##0")
        excWs.Cells(13, 10) = Format(FNN(rs.fields("M08_6")), "###,##0")
        excWs.Cells(13, 11) = Format(FNN(rs.fields("M09_6")), "###,##0")
        excWs.Cells(13, 12) = Format(FNN(rs.fields("M10_6")), "###,##0")
        excWs.Cells(13, 13) = Format(FNN(rs.fields("M11_6")), "###,##0")
        excWs.Cells(13, 14) = Format(FNN(rs.fields("M12_6")), "###,##0")
    
        excWs.Cells(14, 3) = Format(FNDBL(rs.fields("M01_7")), "##0%")
        excWs.Cells(14, 4) = Format(FNDBL(rs.fields("M02_7")), "##0%")
        excWs.Cells(14, 5) = Format(FNDBL(rs.fields("M03_7")), "##0%")
        excWs.Cells(14, 6) = Format(FNDBL(rs.fields("M04_7")), "##0%")
        excWs.Cells(14, 7) = Format(FNDBL(rs.fields("M05_7")), "##0%")
        excWs.Cells(14, 8) = Format(FNDBL(rs.fields("M06_7")), "##0%")
        excWs.Cells(14, 9) = Format(FNDBL(rs.fields("M07_7")), "##0%")
        excWs.Cells(14, 10) = Format(FNDBL(rs.fields("M08_7")), "##0%")
        excWs.Cells(14, 11) = Format(FNDBL(rs.fields("M09_7")), "##0%")
        excWs.Cells(14, 12) = Format(FNDBL(rs.fields("M10_7")), "##0%")
        excWs.Cells(14, 13) = Format(FNDBL(rs.fields("M11_7")), "##0%")
        excWs.Cells(14, 14) = Format(FNDBL(rs.fields("M12_7")), "##0%")
    
        excWs.Cells(15, 3) = Format(FNDBL(rs.fields("M01_8")), "##0%")
        excWs.Cells(15, 4) = Format(FNDBL(rs.fields("M02_8")), "###0%")
        excWs.Cells(15, 5) = Format(FNDBL(rs.fields("M03_8")), "##0%")
        excWs.Cells(15, 6) = Format(FNDBL(rs.fields("M04_8")), "##0%")
        excWs.Cells(15, 7) = Format(FNDBL(rs.fields("M05_8")), "##0%")
        excWs.Cells(15, 8) = Format(FNDBL(rs.fields("M06_8")), "##0%")
        excWs.Cells(15, 9) = Format(FNDBL(rs.fields("M07_8")), "##0%")
        excWs.Cells(15, 10) = Format(FNDBL(rs.fields("M08_8")), "##0%")
        excWs.Cells(15, 11) = Format(FNDBL(rs.fields("M09_8")), "##0%")
        excWs.Cells(15, 12) = Format(FNDBL(rs.fields("M10_8")), "##0%")
        excWs.Cells(15, 13) = Format(FNDBL(rs.fields("M11_8")), "##0%")
        excWs.Cells(15, 14) = Format(FNDBL(rs.fields("M12_8")), "##0%")
    
        excWs.Cells(16, 3) = Format(FNDBL(rs.fields("M01_9")), "##0%")
        excWs.Cells(16, 4) = Format(FNDBL(rs.fields("M02_9")), "##0%")
        excWs.Cells(16, 5) = Format(FNDBL(rs.fields("M03_9")), "##0%")
        excWs.Cells(16, 6) = Format(FNDBL(rs.fields("M04_9")), "##0%")
        excWs.Cells(16, 7) = Format(FNDBL(rs.fields("M05_9")), "##0%")
        excWs.Cells(16, 8) = Format(FNDBL(rs.fields("M06_9")), "##0%")
        excWs.Cells(16, 9) = Format(FNDBL(rs.fields("M07_9")), "##0%")
        excWs.Cells(16, 10) = Format(FNDBL(rs.fields("M08_9")), "##0%")
        excWs.Cells(16, 11) = Format(FNDBL(rs.fields("M09_9")), "##0%")
        excWs.Cells(16, 12) = Format(FNDBL(rs.fields("M10_9")), "##0%")
        excWs.Cells(16, 13) = Format(FNDBL(rs.fields("M11_9")), "##0%")
        excWs.Cells(16, 14) = Format(FNDBL(rs.fields("M12_9")), "##0%")
        
        excWs.Cells(17, 3) = Format(FNDBL(rs.fields("M01_10")), "##0%")
        excWs.Cells(17, 4) = Format(FNDBL(rs.fields("M02_10")), "##0%")
        excWs.Cells(17, 5) = Format(FNDBL(rs.fields("M03_10")), "##0%")
        excWs.Cells(17, 6) = Format(FNDBL(rs.fields("M04_10")), "##0%")
        excWs.Cells(17, 7) = Format(FNDBL(rs.fields("M05_10")), "##0%")
        excWs.Cells(17, 8) = Format(FNDBL(rs.fields("M06_10")), "##0%")
        excWs.Cells(17, 9) = Format(FNDBL(rs.fields("M07_10")), "##0%")
        excWs.Cells(17, 10) = Format(FNDBL(rs.fields("M08_10")), "##0%")
        excWs.Cells(17, 11) = Format(FNDBL(rs.fields("M09_10")), "##0%")
        excWs.Cells(17, 12) = Format(FNDBL(rs.fields("M10_10")), "##0%")
        excWs.Cells(17, 13) = Format(FNDBL(rs.fields("M11_10")), "##0%")
        excWs.Cells(17, 14) = Format(FNDBL(rs.fields("M12_10")), "##0%")
        
        
        excWb.Save
    
        '‘don’t forget to do this or you’ll not be able to open
        '‘book1.xls again, untill you restart you pc.
     '   excApp.ActiveWorkbook.Close False, "c:\budget.xls"
     '   excApp.Quit
        Set excWb = Nothing
        Set excApp = Nothing
        Screen.MousePointer = vbDefault
        
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBudgetPreview.ExportToExcel"
End Sub
Public Sub ExportToOO()
    On Error GoTo errHandler
Dim oSM                   'Root object for accessing OpenOffice from VB
Dim oDesk, oDOC As Object 'First objects from the API
Dim arg()                 'Ignore it for the moment !
Dim aNoArgs()
Dim strFilename As String
Dim fs As New FileSystemObject
 On Error Resume Next
        Screen.MousePointer = vbHourglass
'Instanciate OOo : this line is mandatory with VB for OOo API
    Set oSM = CreateObject("com.sun.star.ServiceManager")
'Create the first and most important service
    Set oDesk = oSM.CreateInstance("com.sun.star.frame.Desktop")
'Open an existing doc (pay attention to the syntax for first argument)
    strFilename = Replace(fs.GetDriveName(oPC.SharedFolderRoot), "\", "/") & "/Templates/OO_Budget.ods"
  ' strFilename = Replace(fs.GetDriveName(oPC.LocalFolder), "\", "/") & "/Templates/OO_Budget.ods"
  '  strFilename = "file:///PBKS_S/Templates/OO_Budget.ods"
    Set oDOC = oDesk.loadComponentFromURL("file:" & strFilename, "_blank", 0, arg())
'Save the doc
    Set oSheet = oDOC.getSheets().getByIndex(0)
        oSheet.getCellRangeByName("A2").setString ("Prepared: " & Format(Now(), "DD-MMM-YYYY HH:NN"))
        
        oSheet.getCellRangeByName("C5").Value = (FND(rs.fields("H01")))
        oSheet.getCellRangeByName("D5").Value = (FND(rs.fields("H02")))
        oSheet.getCellRangeByName("E5").Value = (FND(rs.fields("H03")))
        oSheet.getCellRangeByName("F5").Value = (FND(rs.fields("H04")))
        oSheet.getCellRangeByName("G5").Value = (FND(rs.fields("H05")))
        oSheet.getCellRangeByName("H5").Value = (FND(rs.fields("H06")))
        oSheet.getCellRangeByName("I5").Value = (FND(rs.fields("H07")))
        oSheet.getCellRangeByName("J5").Value = (FND(rs.fields("H08")))
        oSheet.getCellRangeByName("K5").Value = (FND(rs.fields("H09")))
        oSheet.getCellRangeByName("L5").Value = (FND(rs.fields("H10")))
        oSheet.getCellRangeByName("M5").Value = (FND(rs.fields("H11")))
        oSheet.getCellRangeByName("N5").Value = (FND(rs.fields("H12")))
    
        oSheet.getCellRangeByName("C6").Value = FNDBL(rs.fields("M01_0"))
        oSheet.getCellRangeByName("D6").Value = FNDBL(rs.fields("M02_0"))
        oSheet.getCellRangeByName("E6").Value = FNDBL(rs.fields("M03_0"))
        oSheet.getCellRangeByName("F6").Value = FNDBL(rs.fields("M04_0"))
        oSheet.getCellRangeByName("G6").Value = FNDBL(rs.fields("M05_0"))
        oSheet.getCellRangeByName("H6").Value = FNDBL(rs.fields("M06_0"))
        oSheet.getCellRangeByName("I6").Value = FNDBL(rs.fields("M07_0"))
        oSheet.getCellRangeByName("J6").Value = FNDBL(rs.fields("M08_0"))
        oSheet.getCellRangeByName("K6").Value = FNDBL(rs.fields("M09_0"))
        oSheet.getCellRangeByName("L6").Value = FNDBL(rs.fields("M10_0"))
        oSheet.getCellRangeByName("M6").Value = FNDBL(rs.fields("M11_0"))
        oSheet.getCellRangeByName("N6").Value = FNDBL(rs.fields("M12_0"))
        
        oSheet.getCellRangeByName("C7").Value = FNDBL(rs.fields("M01_1"))
        oSheet.getCellRangeByName("D7").Value = FNDBL(rs.fields("M02_1"))
        oSheet.getCellRangeByName("E7").Value = FNDBL(rs.fields("M03_1"))
        oSheet.getCellRangeByName("F7").Value = FNDBL(rs.fields("M04_1"))
        oSheet.getCellRangeByName("G7").Value = FNDBL(rs.fields("M05_1"))
        oSheet.getCellRangeByName("H7").Value = FNDBL(rs.fields("M06_1"))
        oSheet.getCellRangeByName("I7").Value = FNDBL(rs.fields("M07_1"))
        oSheet.getCellRangeByName("J7").Value = FNDBL(rs.fields("M08_1"))
        oSheet.getCellRangeByName("K7").Value = FNDBL(rs.fields("M09_1"))
        oSheet.getCellRangeByName("L7").Value = FNDBL(rs.fields("M10_1"))
        oSheet.getCellRangeByName("M7").Value = FNDBL(rs.fields("M11_1"))
        oSheet.getCellRangeByName("N7").Value = FNDBL(rs.fields("M12_1"))
    
        oSheet.getCellRangeByName("C8").Value = FNDBL(rs.fields("M01_2"))
        oSheet.getCellRangeByName("D8").Value = FNDBL(rs.fields("M02_2"))
        oSheet.getCellRangeByName("E8").Value = FNDBL(rs.fields("M03_2"))
        oSheet.getCellRangeByName("F8").Value = FNDBL(rs.fields("M04_2"))
        oSheet.getCellRangeByName("G8").Value = FNDBL(rs.fields("M05_2"))
        oSheet.getCellRangeByName("H8").Value = FNDBL(rs.fields("M06_2"))
        oSheet.getCellRangeByName("I8").Value = FNDBL(rs.fields("M07_2"))
        oSheet.getCellRangeByName("J8").Value = FNDBL(rs.fields("M08_2"))
        oSheet.getCellRangeByName("K8").Value = FNDBL(rs.fields("M09_2"))
        oSheet.getCellRangeByName("L8").Value = FNDBL(rs.fields("M10_2"))
        oSheet.getCellRangeByName("M8").Value = FNDBL(rs.fields("M11_2"))
        oSheet.getCellRangeByName("N8").Value = FNDBL(rs.fields("M12_2"))
    
        oSheet.getCellRangeByName("C9").Value = FNDBL(rs.fields("M01_3"))
        oSheet.getCellRangeByName("D9").Value = FNDBL(rs.fields("M02_3"))
        oSheet.getCellRangeByName("E9").Value = FNDBL(rs.fields("M03_3"))
        oSheet.getCellRangeByName("F9").Value = FNDBL(rs.fields("M04_3"))
        oSheet.getCellRangeByName("G9").Value = FNDBL(rs.fields("M05_3"))
        oSheet.getCellRangeByName("H9").Value = FNDBL(rs.fields("M06_3"))
        oSheet.getCellRangeByName("I9").Value = FNDBL(rs.fields("M07_3"))
        oSheet.getCellRangeByName("J9").Value = FNDBL(rs.fields("M08_3"))
        oSheet.getCellRangeByName("K9").Value = FNDBL(rs.fields("M09_3"))
        oSheet.getCellRangeByName("L9").Value = FNDBL(rs.fields("M10_3"))
        oSheet.getCellRangeByName("M9").Value = FNDBL(rs.fields("M11_3"))
        oSheet.getCellRangeByName("N9").Value = FNDBL(rs.fields("M12_3"))
    
        oSheet.getCellRangeByName("C10").Value = FNDBL(rs.fields("M01_4"))
        oSheet.getCellRangeByName("D10").Value = FNDBL(rs.fields("M02_4"))
        oSheet.getCellRangeByName("E10").Value = FNDBL(rs.fields("M03_4"))
        oSheet.getCellRangeByName("F10").Value = FNDBL(rs.fields("M04_4"))
        oSheet.getCellRangeByName("G10").Value = FNDBL(rs.fields("M05_4"))
        oSheet.getCellRangeByName("H10").Value = FNDBL(rs.fields("M06_4"))
        oSheet.getCellRangeByName("I10").Value = FNDBL(rs.fields("M07_4"))
        oSheet.getCellRangeByName("J10").Value = FNDBL(rs.fields("M08_4"))
        oSheet.getCellRangeByName("K10").Value = FNDBL(rs.fields("M09_4"))
        oSheet.getCellRangeByName("L10").Value = FNDBL(rs.fields("M10_4"))
        oSheet.getCellRangeByName("M10").Value = FNDBL(rs.fields("M11_4"))
        oSheet.getCellRangeByName("N10").Value = FNDBL(rs.fields("M12_4"))
   
        oSheet.getCellRangeByName("C11").Value = FNDBL(rs.fields("M01_5b"))
        oSheet.getCellRangeByName("D11").Value = FNDBL(rs.fields("M02_5b"))
        oSheet.getCellRangeByName("E11").Value = FNDBL(rs.fields("M03_5b"))
        oSheet.getCellRangeByName("F11").Value = FNDBL(rs.fields("M04_5b"))
        oSheet.getCellRangeByName("G11").Value = FNDBL(rs.fields("M05_5b"))
        oSheet.getCellRangeByName("H11").Value = FNDBL(rs.fields("M06_5b"))
        oSheet.getCellRangeByName("I11").Value = FNDBL(rs.fields("M07_5b"))
        oSheet.getCellRangeByName("J11").Value = FNDBL(rs.fields("M08_5b"))
        oSheet.getCellRangeByName("K11").Value = FNDBL(rs.fields("M09_5b"))
        oSheet.getCellRangeByName("L11").Value = FNDBL(rs.fields("M10_5b"))
        oSheet.getCellRangeByName("M11").Value = FNDBL(rs.fields("M11_5b"))
        oSheet.getCellRangeByName("N11").Value = FNDBL(rs.fields("M12_5b"))
       
        oSheet.getCellRangeByName("C12").Value = FNDBL(rs.fields("M01_5"))
        oSheet.getCellRangeByName("D12").Value = FNDBL(rs.fields("M02_5"))
        oSheet.getCellRangeByName("E12").Value = FNDBL(rs.fields("M03_5"))
        oSheet.getCellRangeByName("F12").Value = FNDBL(rs.fields("M04_5"))
        oSheet.getCellRangeByName("G12").Value = FNDBL(rs.fields("M05_5"))
        oSheet.getCellRangeByName("H12").Value = FNDBL(rs.fields("M06_5"))
        oSheet.getCellRangeByName("I12").Value = FNDBL(rs.fields("M07_5"))
        oSheet.getCellRangeByName("J12").Value = FNDBL(rs.fields("M08_5"))
        oSheet.getCellRangeByName("K12").Value = FNDBL(rs.fields("M09_5"))
        oSheet.getCellRangeByName("L12").Value = FNDBL(rs.fields("M10_5"))
        oSheet.getCellRangeByName("M12").Value = FNDBL(rs.fields("M11_5"))
        oSheet.getCellRangeByName("N12").Value = FNDBL(rs.fields("M12_5"))
    
        oSheet.getCellRangeByName("C13").Value = FNDBL(rs.fields("M01_6"))
        oSheet.getCellRangeByName("D13").Value = FNDBL(rs.fields("M02_6"))
        oSheet.getCellRangeByName("E13").Value = FNDBL(rs.fields("M03_6"))
        oSheet.getCellRangeByName("F13").Value = FNDBL(rs.fields("M04_6"))
        oSheet.getCellRangeByName("G13").Value = FNDBL(rs.fields("M05_6"))
        oSheet.getCellRangeByName("H13").Value = FNDBL(rs.fields("M06_6"))
        oSheet.getCellRangeByName("I13").Value = FNDBL(rs.fields("M07_6"))
        oSheet.getCellRangeByName("J13").Value = FNDBL(rs.fields("M08_6"))
        oSheet.getCellRangeByName("K13").Value = FNDBL(rs.fields("M09_6"))
        oSheet.getCellRangeByName("L13").Value = FNDBL(rs.fields("M10_6"))
        oSheet.getCellRangeByName("M13").Value = FNDBL(rs.fields("M11_6"))
        oSheet.getCellRangeByName("N13").Value = FNDBL(rs.fields("M12_6"))
    
        oSheet.getCellRangeByName("C14").Value = FNDBL(rs.fields("M01_7"))
        oSheet.getCellRangeByName("D14").Value = FNDBL(rs.fields("M02_7"))
        oSheet.getCellRangeByName("E14").Value = FNDBL(rs.fields("M03_7"))
        oSheet.getCellRangeByName("F14").Value = FNDBL(rs.fields("M04_7"))
        oSheet.getCellRangeByName("G14").Value = FNDBL(rs.fields("M05_7"))
        oSheet.getCellRangeByName("H14").Value = FNDBL(rs.fields("M06_7"))
        oSheet.getCellRangeByName("I14").Value = FNDBL(rs.fields("M07_7"))
        oSheet.getCellRangeByName("J14").Value = FNDBL(rs.fields("M08_7"))
        oSheet.getCellRangeByName("K14").Value = FNDBL(rs.fields("M09_7"))
        oSheet.getCellRangeByName("L14").Value = FNDBL(rs.fields("M10_7"))
        oSheet.getCellRangeByName("M14").Value = FNDBL(rs.fields("M11_7"))
        oSheet.getCellRangeByName("N14").Value = FNDBL(rs.fields("M12_7"))
    
        oSheet.getCellRangeByName("C15").Value = FNDBL(rs.fields("M01_8"))
        oSheet.getCellRangeByName("D15").Value = FNDBL(rs.fields("M02_8"))
        oSheet.getCellRangeByName("E15").Value = FNDBL(rs.fields("M03_8"))
        oSheet.getCellRangeByName("F15").Value = FNDBL(rs.fields("M04_8"))
        oSheet.getCellRangeByName("G15").Value = FNDBL(rs.fields("M05_8"))
        oSheet.getCellRangeByName("H15").Value = FNDBL(rs.fields("M06_8"))
        oSheet.getCellRangeByName("I15").Value = FNDBL(rs.fields("M07_8"))
        oSheet.getCellRangeByName("J15").Value = FNDBL(rs.fields("M08_8"))
        oSheet.getCellRangeByName("K15").Value = FNDBL(rs.fields("M09_8"))
        oSheet.getCellRangeByName("L15").Value = FNDBL(rs.fields("M10_8"))
        oSheet.getCellRangeByName("M15").Value = FNDBL(rs.fields("M11_8"))
        oSheet.getCellRangeByName("N15").Value = FNDBL(rs.fields("M12_8"))
    
        oSheet.getCellRangeByName("C16").Value = FNDBL(rs.fields("M01_9"))
        oSheet.getCellRangeByName("D16").Value = FNDBL(rs.fields("M02_9"))
        oSheet.getCellRangeByName("E16").Value = FNDBL(rs.fields("M03_9"))
        oSheet.getCellRangeByName("F16").Value = FNDBL(rs.fields("M04_9"))
        oSheet.getCellRangeByName("G16").Value = FNDBL(rs.fields("M05_9"))
        oSheet.getCellRangeByName("H16").Value = FNDBL(rs.fields("M06_9"))
        oSheet.getCellRangeByName("I16").Value = FNDBL(rs.fields("M07_9"))
        oSheet.getCellRangeByName("J16").Value = FNDBL(rs.fields("M08_9"))
        oSheet.getCellRangeByName("K16").Value = FNDBL(rs.fields("M09_9"))
        oSheet.getCellRangeByName("L16").Value = FNDBL(rs.fields("M10_9"))
        oSheet.getCellRangeByName("M16").Value = FNDBL(rs.fields("M11_9"))
        oSheet.getCellRangeByName("N16").Value = FNDBL(rs.fields("M12_9"))
        
        oSheet.getCellRangeByName("C17").Value = FNDBL(rs.fields("M01_10"))
        oSheet.getCellRangeByName("D17").Value = FNDBL(rs.fields("M02_10"))
        oSheet.getCellRangeByName("E17").Value = FNDBL(rs.fields("M03_10"))
        oSheet.getCellRangeByName("F17").Value = FNDBL(rs.fields("M04_10"))
        oSheet.getCellRangeByName("G17").Value = FNDBL(rs.fields("M05_10"))
        oSheet.getCellRangeByName("H17").Value = FNDBL(rs.fields("M06_10"))
        oSheet.getCellRangeByName("I17").Value = FNDBL(rs.fields("M07_10"))
        oSheet.getCellRangeByName("J17").Value = FNDBL(rs.fields("M08_10"))
        oSheet.getCellRangeByName("K17").Value = FNDBL(rs.fields("M09_10"))
        oSheet.getCellRangeByName("L17").Value = FNDBL(rs.fields("M10_10"))
        oSheet.getCellRangeByName("M17").Value = FNDBL(rs.fields("M11_10"))
        oSheet.getCellRangeByName("N17").Value = FNDBL(rs.fields("M12_10"))
    
    oDOC.Store
   ' oDOC.Close (True)
    
    Set oDOC = Nothing
    Set oDesk = Nothing
   ' oSM.Close
    Set oSM = Nothing
    Screen.MousePointer = vbDefault

'errHandler:
'    ErrorIn "frmBudgetPreview.ExportToOO"
'errHandler:
'    ErrorIn "frmBudgetPreview.ExportToOO"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBudgetPreview.ExportToOO"
End Sub
Public Function ConvertToUrl(strFile) As String
    On Error GoTo errHandler
'With thanks to didier.alain@kalitech.fr code used from http://www.kalitech.fr/clients/doc/VB_APIOOo_en.html#replacement_functions

  '  strFile = Replace(strFile, "\", "/")
   ' strFile = Replace(strFile, ":", "|")
    strFile = Replace(strFile, " ", "%20")
    strFile = "file:///" + strFile
    ConvertToUrl = strFile

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBudgetPreview.ConvertToUrl(strFile)", strFile
End Function
