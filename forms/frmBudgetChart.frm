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
      Picture         =   "frmBudgetChart.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   30
      Width           =   315
   End
   Begin VB.CommandButton cmdHide 
      Height          =   285
      Left            =   975
      Picture         =   "frmBudgetChart.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   131
      Top             =   30
      Width           =   315
   End
   Begin VB.CommandButton cmdExport 
      Height          =   285
      Left            =   660
      Picture         =   "frmBudgetChart.frx":0714
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
         Size            =   14.25
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
15480     On Error GoTo errHandler


15490     Set rs = oPC.rsBudgetData
15500     H_1 = Format(FND(rs.fields("H01")), "MM-YY")
15510     H_2 = Format(FND(rs.fields("H02")), "MM-YY")
15520     H_3 = Format(FND(rs.fields("H03")), "MM-YY")
15530     H_4 = Format(FND(rs.fields("H04")), "MM-YY")
15540     H_5 = Format(FND(rs.fields("H05")), "MM-YY")
15550     H_6 = Format(FND(rs.fields("H06")), "MM-YY")
15560     H_7 = Format(FND(rs.fields("H07")), "MM-YY")
15570     H_8 = Format(FND(rs.fields("H08")), "MM-YY")
15580     H_9 = Format(FND(rs.fields("H09")), "MM-YY")
15590     H_10 = Format(FND(rs.fields("H10")), "MM-YY")
15600     H_11 = Format(FND(rs.fields("H11")), "MM-YY")
15610     H_12 = Format(FND(rs.fields("H12")), "MM-YY")
          
15620     L1_0 = Format(FNN(rs.fields("M01_0")), "###,##0")
15630     L2_0 = Format(FNN(rs.fields("M02_0")), "###,##0")
15640     L3_0 = Format(FNN(rs.fields("M03_0")), "###,##0")
15650     L4_0 = Format(FNN(rs.fields("M04_0")), "###,##0")
15660     L5_0 = Format(FNN(rs.fields("M05_0")), "###,##0")
15670     L6_0 = Format(FNN(rs.fields("M06_0")), "###,##0")
15680     L7_0 = Format(FNN(rs.fields("M07_0")), "###,##0")
15690     L8_0 = Format(FNN(rs.fields("M08_0")), "###,##0")
15700     L9_0 = Format(FNN(rs.fields("M09_0")), "###,##0")
15710     L10_0 = Format(FNN(rs.fields("M10_0")), "###,##0")
15720     L11_0 = Format(FNN(rs.fields("M11_0")), "###,##0")
15730     L12_0 = Format(FNN(rs.fields("M12_0")), "###,##0")
          
15740     L1_1 = Format(FNN(rs.fields("M01_1")), "###,##0")
15750     L2_1 = Format(FNN(rs.fields("M02_1")), "###,##0")
15760     L3_1 = Format(FNN(rs.fields("M03_1")), "###,##0")
15770     L4_1 = Format(FNN(rs.fields("M04_1")), "###,##0")
15780     L5_1 = Format(FNN(rs.fields("M05_1")), "###,##0")
15790     L6_1 = Format(FNN(rs.fields("M06_1")), "###,##0")
15800     L7_1 = Format(FNN(rs.fields("M07_1")), "###,##0")
15810     L8_1 = Format(FNN(rs.fields("M08_1")), "###,##0")
15820     L9_1 = Format(FNN(rs.fields("M09_1")), "###,##0")
15830     L10_1 = Format(FNN(rs.fields("M10_1")), "###,##0")
15840     L11_1 = Format(FNN(rs.fields("M11_1")), "###,##0")
15850     L12_1 = Format(FNN(rs.fields("M12_1")), "###,##0")
          
15860     L1_2 = Format(FNN(rs.fields("M01_2")), "###,##0")
15870     L2_2 = Format(FNN(rs.fields("M02_2")), "###,##0")
15880     L3_2 = Format(FNN(rs.fields("M03_2")), "###,##0")
15890     L4_2 = Format(FNN(rs.fields("M04_2")), "###,##0")
15900     L5_2 = Format(FNN(rs.fields("M05_2")), "###,##0")
15910     L6_2 = Format(FNN(rs.fields("M06_2")), "###,##0")
15920     L7_2 = Format(FNN(rs.fields("M07_2")), "###,##0")
15930     L8_2 = Format(FNN(rs.fields("M08_2")), "###,##0")
15940     L9_2 = Format(FNN(rs.fields("M09_2")), "###,##0")
15950     L10_2 = Format(FNN(rs.fields("M10_2")), "###,##0")
15960     L11_2 = Format(FNN(rs.fields("M11_2")), "###,##0")
15970     L12_2 = Format(FNN(rs.fields("M12_2")), "###,##0")
          
15980     L1_3 = Format(FNN(rs.fields("M01_3")), "###,##0")
15990     L2_3 = Format(FNN(rs.fields("M02_3")), "###,##0")
16000     L3_3 = Format(FNN(rs.fields("M03_3")), "###,##0")
16010     L4_3 = Format(FNN(rs.fields("M04_3")), "###,##0")
16020     L5_3 = Format(FNN(rs.fields("M05_3")), "###,##0")
16030     L6_3 = Format(FNN(rs.fields("M06_3")), "###,##0")
16040     L7_3 = Format(FNN(rs.fields("M07_3")), "###,##0")
16050     L8_3 = Format(FNN(rs.fields("M08_3")), "###,##0")
16060     L9_3 = Format(FNN(rs.fields("M09_3")), "###,##0")
16070     L10_3 = Format(FNN(rs.fields("M10_3")), "###,##0")
16080     L11_3 = Format(FNN(rs.fields("M11_3")), "###,##0")
16090     L12_3 = Format(FNN(rs.fields("M12_3")), "###,##0")
          
16100     L1_4 = Format(FNN(rs.fields("M01_4")), "###,##0")
16110     L2_4 = Format(FNN(rs.fields("M02_4")), "###,##0")
16120     L3_4 = Format(FNN(rs.fields("M03_4")), "###,##0")
16130     L4_4 = Format(FNN(rs.fields("M04_4")), "###,##0")
16140     L5_4 = Format(FNN(rs.fields("M05_4")), "###,##0")
16150     L6_4 = Format(FNN(rs.fields("M06_4")), "###,##0")
16160     L7_4 = Format(FNN(rs.fields("M07_4")), "###,##0")
16170     L8_4 = Format(FNN(rs.fields("M08_4")), "###,##0")
16180     L9_4 = Format(FNN(rs.fields("M09_4")), "###,##0")
16190     L10_4 = Format(FNN(rs.fields("M10_4")), "###,##0")
16200     L11_4 = Format(FNN(rs.fields("M11_4")), "###,##0")
16210     L12_4 = Format(FNN(rs.fields("M12_4")), "###,##0")
          
16220     L1_5 = Format(FNN(rs.fields("M01_5")), "###,##0")
16230     L2_5 = Format(FNN(rs.fields("M02_5")), "###,##0")
16240     L3_5 = Format(FNN(rs.fields("M03_5")), "###,##0")
16250     L4_5 = Format(FNN(rs.fields("M04_5")), "###,##0")
16260     L5_5 = Format(FNN(rs.fields("M05_5")), "###,##0")
16270     L6_5 = Format(FNN(rs.fields("M06_5")), "###,##0")
16280     L7_5 = Format(FNN(rs.fields("M07_5")), "###,##0")
16290     L8_5 = Format(FNN(rs.fields("M08_5")), "###,##0")
16300     L9_5 = Format(FNN(rs.fields("M09_5")), "###,##0")
16310     L10_5 = Format(FNN(rs.fields("M10_5")), "###,##0")
16320     L11_5 = Format(FNN(rs.fields("M11_5")), "###,##0")
16330     L12_5 = Format(FNN(rs.fields("M12_5")), "###,##0")
          
16340     L1_5b = Format(FNN(rs.fields("M01_5b")), "###,##0")
16350     L2_5b = Format(FNN(rs.fields("M02_5b")), "###,##0")
16360     L3_5b = Format(FNN(rs.fields("M03_5b")), "###,##0")
16370     L4_5b = Format(FNN(rs.fields("M04_5b")), "###,##0")
16380     L5_5b = Format(FNN(rs.fields("M05_5b")), "###,##0")
16390     L6_5b = Format(FNN(rs.fields("M06_5b")), "###,##0")
16400     L7_5b = Format(FNN(rs.fields("M07_5b")), "###,##0")
16410     L8_5b = Format(FNN(rs.fields("M08_5b")), "###,##0")
16420     L9_5b = Format(FNN(rs.fields("M09_5b")), "###,##0")
16430     L10_5b = Format(FNN(rs.fields("M10_5b")), "###,##0")
16440     L11_5b = Format(FNN(rs.fields("M11_5b")), "###,##0")
16450     L12_5b = Format(FNN(rs.fields("M12_5b")), "###,##0")
          
          
16460     L1_6 = Format(FNN(rs.fields("M01_6")), "###,##0")
16470     L2_6 = Format(FNN(rs.fields("M02_6")), "###,##0")
16480     L3_6 = Format(FNN(rs.fields("M03_6")), "###,##0")
16490     L4_6 = Format(FNN(rs.fields("M04_6")), "###,##0")
16500     L5_6 = Format(FNN(rs.fields("M05_6")), "###,##0")
16510     L6_6 = Format(FNN(rs.fields("M06_6")), "###,##0")
16520     L7_6 = Format(FNN(rs.fields("M07_6")), "###,##0")
16530     L8_6 = Format(FNN(rs.fields("M08_6")), "###,##0")
16540     L9_6 = Format(FNN(rs.fields("M09_6")), "###,##0")
16550     L10_6 = Format(FNN(rs.fields("M10_6")), "###,##0")
16560     L11_6 = Format(FNN(rs.fields("M11_6")), "###,##0")
16570     L12_6 = Format(FNN(rs.fields("M12_6")), "###,##0")
          
16580     L1_7 = Format(FNDBL(rs.fields("M01_7")), "##0%")
16590     L2_7 = Format(FNDBL(rs.fields("M02_7")), "##0%")
16600     L3_7 = Format(FNDBL(rs.fields("M03_7")), "##0%")
16610     L4_7 = Format(FNDBL(rs.fields("M04_7")), "##0%")
16620     L5_7 = Format(FNDBL(rs.fields("M05_7")), "##0%")
16630     L6_7 = Format(FNDBL(rs.fields("M06_7")), "##0%")
16640     L7_7 = Format(FNDBL(rs.fields("M07_7")), "##0%")
16650     L8_7 = Format(FNDBL(rs.fields("M08_7")), "##0%")
16660     L9_7 = Format(FNDBL(rs.fields("M09_7")), "##0%")
16670     L10_7 = Format(FNDBL(rs.fields("M10_7")), "##0%")
16680     L11_7 = Format(FNDBL(rs.fields("M11_7")), "##0%")
16690     L12_7 = Format(FNDBL(rs.fields("M12_7")), "##0%")
          
16700     L1_8 = Format(FNDBL(rs.fields("M01_8")), "##0%")
16710     L2_8 = Format(FNDBL(rs.fields("M02_8")), "###0%")
16720     L3_8 = Format(FNDBL(rs.fields("M03_8")), "##0%")
16730     L4_8 = Format(FNDBL(rs.fields("M04_8")), "##0%")
16740     L5_8 = Format(FNDBL(rs.fields("M05_8")), "##0%")
16750     L6_8 = Format(FNDBL(rs.fields("M06_8")), "##0%")
16760     L7_8 = Format(FNDBL(rs.fields("M07_8")), "##0%")
16770     L8_8 = Format(FNDBL(rs.fields("M08_8")), "##0%")
16780     L9_8 = Format(FNDBL(rs.fields("M09_8")), "##0%")
16790     L10_8 = Format(FNDBL(rs.fields("M10_8")), "##0%")
16800     L11_8 = Format(FNDBL(rs.fields("M11_8")), "##0%")
16810     L12_8 = Format(FNDBL(rs.fields("M12_8")), "##0%")
          
16820     L1_9 = Format(FNDBL(rs.fields("M01_9")), "##0%")
16830     L2_9 = Format(FNDBL(rs.fields("M02_9")), "##0%")
16840     L3_9 = Format(FNDBL(rs.fields("M03_9")), "##0%")
16850     L4_9 = Format(FNDBL(rs.fields("M04_9")), "##0%")
16860     L5_9 = Format(FNDBL(rs.fields("M05_9")), "##0%")
16870     L6_9 = Format(FNDBL(rs.fields("M06_9")), "##0%")
16880     L7_9 = Format(FNDBL(rs.fields("M07_9")), "##0%")
16890     L8_9 = Format(FNDBL(rs.fields("M08_9")), "##0%")
16900     L9_9 = Format(FNDBL(rs.fields("M09_9")), "##0%")
16910     L10_9 = Format(FNDBL(rs.fields("M10_9")), "##0%")
16920     L11_9 = Format(FNDBL(rs.fields("M11_9")), "##0%")
16930     L12_9 = Format(FNDBL(rs.fields("M12_9")), "##0%")
          
16940     L1_10 = Format(FNDBL(rs.fields("M01_10")), "##0%")
16950     L2_10 = Format(FNDBL(rs.fields("M02_10")), "##0%")
16960     L3_10 = Format(FNDBL(rs.fields("M03_10")), "##0%")
16970     L4_10 = Format(FNDBL(rs.fields("M04_10")), "##0%")
16980     L5_10 = Format(FNDBL(rs.fields("M05_10")), "##0%")
16990     L6_10 = Format(FNDBL(rs.fields("M06_10")), "##0%")
17000     L7_10 = Format(FNDBL(rs.fields("M07_10")), "##0%")
17010     L8_10 = Format(FNDBL(rs.fields("M08_10")), "##0%")
17020     L9_10 = Format(FNDBL(rs.fields("M09_10")), "##0%")
17030     L10_10 = Format(FNDBL(rs.fields("M10_10")), "##0%")
17040     L11_10 = Format(FNDBL(rs.fields("M11_10")), "##0%")
17050     L12_10 = Format(FNDBL(rs.fields("M12_10")), "##0%")
17060     lblMemo.Caption = "Data was last calculated on " & Format(FND(rs.fields("B_LastCalculated")), "DD-MM-YY Hh:Nn") & ". Note suppliers invoices value include unissued documents."
          
17070     Exit Sub
errHandler:
17080     If ErrMustStop Then Debug.Assert False: Resume
17090     ErrorIn "frmBudgetPreview.LoadData"
End Sub


Private Sub cmdHide_Click()
17100     Me.Height = 3920
17110     Me.Width = 11180
17120     Me.TOP = Forms(0).Height - Forms(0).TOP - Me.Height - 1600
17130     Me.Left = 0

17140     Me.Hide
End Sub

Private Sub cmdRefresh_Click()
17150     On Error GoTo errHandler
17160     If MsgBox("This refresh action may take a few minutes. Wait for it to complete before attempting further actions in Papyrus." & vbCrLf & "Click Cancel button to skip refresh action.", vbOKCancel, "Confirm refresh") = vbCancel Then
17170         Exit Sub
17180     End If
17190     Screen.MousePointer = vbHourglass
17200     oPC.ReloadBudget
17210     LoadData
17220     Screen.MousePointer = vbDefault
17230     Exit Sub
errHandler:
17240     If ErrMustStop Then Debug.Assert False: Resume
17250     ErrorIn "frmBudgetPreview.cmdRefresh_Click", , EA_NORERAISE
17260     HandleError
End Sub

Private Sub Form_Activate()
17270     On Error GoTo errHandler
17280     Me.Height = 3920
17290     Me.Width = 11180
17300     Me.TOP = Forms(0).Height - Forms(0).TOP - Me.Height - 1600
17310     Me.Left = 0
17320     LoadData
17330     Exit Sub
errHandler:
17340     If ErrMustStop Then Debug.Assert False: Resume
17350     ErrorIn "frmBudgetPreview.Form_Activate", , EA_NORERAISE
17360     HandleError
End Sub

Private Sub Form_Load()
17370     On Error GoTo errHandler
        '  MsgBox "Budget 1"
17380     Me.Height = 3920
17390     Me.Width = 11180
17400     Me.TOP = Forms(0).Height - Forms(0).TOP - Me.Height - 1600
17410     Me.Left = 0
      'MsgBox "Budget 2"
17420     LoadData
      'MsgBox "Budget 3"
17430     Exit Sub
errHandler:
17440     If ErrMustStop Then Debug.Assert False: Resume
17450     ErrorIn "frmBudgetPreview.Form_Load", , EA_NORERAISE
17460     HandleError
End Sub

Private Sub lblHelp_Click()
17470     On Error GoTo errHandler
      Dim f As New frmHelpGen
      Dim s As String

17480 s = "Notes on understanding the budget report" & vbCrLf
17490 s = s & "========================================" & vbCrLf & vbCrLf
17500 s = s & "1.  The columns are headed by dates and they refer to the Expected delivery date (ETA) i.r.o. purchase orders" & vbCrLf
17510 s = s & "    and to supplier invoice dates in respect of deliveries." & vbCrLf & vbCrLf
17520 s = s & "2. The first band represents actual sales made in that month." & vbCrLf & vbCrLf
17530 s = s & "3. The blue band represents the budgetted P.O.s receivable in that month and the returns to be effected in that month." & vbCrLf & vbCrLf
17540 s = s & "4. The band below the blue represents suppliers invoices received with dates in that month and actual returns removed from stock in that month." & vbCrLf & vbCrLf
17550 s = s & "5. The yellow band represents the difference (nett effect of the above). The cost incurred in that month." & vbCrLf & vbCrLf
17560 s = s & "6. The plain band below the yellow band represents the actual orders placed with delivery dates(ETA) in that month. " & vbCrLf
17570 s = s & "   The second row shows any unissued orders. The third row shows the actual orders value compared to the budgetted value as a percentage." & vbCrLf & vbCrLf
17580 s = s & "7. The red band shows the (deliveries - returns) against the budget." & vbCrLf & vbCrLf
17590 s = s & "8. The last band shows the average of the red band over the last 4 months." & vbCrLf
17600 s = s & "   Stock is not always delivered in the actual month that it is expected so a smoothed average is necessary." & vbCrLf
17610 s = s & "   The last line shows the spread between subscription and replenishment orders."
17620     f.component s, "Budget help", 13000, 6500
17630     f.Show
      'MsgBox "Still coming"
17640     Exit Sub
errHandler:
17650     If ErrMustStop Then Debug.Assert False: Resume
17660     ErrorIn "frmBudgetPreview.lblHelp_Click", , EA_NORERAISE
17670     HandleError
End Sub

Private Sub cmdExport_Click()
17680     On Error GoTo errHandler

17690     If oPC.UsesExcel Then
17700         ExportToExcel
17710     Else
17720         ExportToOO
17730     End If

17740     Exit Sub
errHandler:
17750     If ErrMustStop Then Debug.Assert False: Resume
17760     ErrorIn "frmBudgetPreview.cmdExport_Click", , EA_NORERAISE
17770     HandleError
End Sub
Public Sub ExportToExcel()
17780     On Error GoTo errHandler
17790         Screen.MousePointer = vbHourglass

17800         Set excApp = CreateObject("excel.application")
17810         Set excWb = excApp.Workbooks.open(oPC.SharedFolderRoot & "\Templates\budget.xls")
             ' Set excWb = excApp.Workbooks.Open(oPC.LocalFolder & "\Templates\budget.xls")
17820         Set excWs = excWb.Sheets.Item(1)
17830         excWs.Application.Visible = True
              
17840         excWs.Cells(2, 1) = "Prepared: " & Format(Now(), "DD-MMM-YYYY HH:NN")
17850         excWs.Cells(5, 3) = Format(FND(rs.fields("H01")), "DD-MMM-YYYY")
17860         excWs.Cells(5, 4) = Format(FND(rs.fields("H02")), "DD-MMM-YYYY")
17870         excWs.Cells(5, 5) = Format(FND(rs.fields("H03")), "DD-MMM-YYYY")
17880         excWs.Cells(5, 6) = Format(FND(rs.fields("H04")), "DD-MMM-YYYY")
17890         excWs.Cells(5, 7) = Format(FND(rs.fields("H05")), "DD-MMM-YYYY")
17900         excWs.Cells(5, 8) = Format(FND(rs.fields("H06")), "DD-MMM-YYYY")
17910         excWs.Cells(5, 9) = Format(FND(rs.fields("H07")), "DD-MMM-YYYY")
17920         excWs.Cells(5, 10) = Format(FND(rs.fields("H08")), "DD-MMM-YYYY")
17930         excWs.Cells(5, 11) = Format(FND(rs.fields("H09")), "DD-MMM-YYYY")
17940         excWs.Cells(5, 12) = Format(FND(rs.fields("H10")), "DD-MMM-YYYY")
17950         excWs.Cells(5, 13) = Format(FND(rs.fields("H11")), "DD-MMM-YYYY")
17960         excWs.Cells(5, 14) = Format(FND(rs.fields("H12")), "DD-MMM-YYYY")

      'Sales
17970         excWs.Cells(6, 3) = Format(FNN(rs.fields("M01_0")), "###,##0")
17980         excWs.Cells(6, 4) = Format(FNN(rs.fields("M02_0")), "###,##0")
17990         excWs.Cells(6, 5) = Format(FNN(rs.fields("M03_0")), "###,##0")
18000         excWs.Cells(6, 6) = Format(FNN(rs.fields("M04_0")), "###,##0")
18010         excWs.Cells(6, 7) = Format(FNN(rs.fields("M05_0")), "###,##0")
18020         excWs.Cells(6, 8) = Format(FNN(rs.fields("M06_0")), "###,##0")
18030         excWs.Cells(6, 9) = Format(FNN(rs.fields("M07_0")), "###,##0")
18040         excWs.Cells(6, 10) = Format(FNN(rs.fields("M08_0")), "###,##0")
18050         excWs.Cells(6, 11) = Format(FNN(rs.fields("M09_0")), "###,##0")
18060         excWs.Cells(6, 12) = Format(FNN(rs.fields("M10_0")), "###,##0")
18070         excWs.Cells(6, 13) = Format(FNN(rs.fields("M11_0")), "###,##0")
18080         excWs.Cells(6, 14) = Format(FNN(rs.fields("M12_0")), "###,##0")
              
              
18090         excWs.Cells(7, 3) = Format(FNN(rs.fields("M01_1")), "###,##0")
18100         excWs.Cells(7, 4) = Format(FNN(rs.fields("M02_1")), "###,##0")
18110         excWs.Cells(7, 5) = Format(FNN(rs.fields("M03_1")), "###,##0")
18120         excWs.Cells(7, 6) = Format(FNN(rs.fields("M04_1")), "###,##0")
18130         excWs.Cells(7, 7) = Format(FNN(rs.fields("M05_1")), "###,##0")
18140         excWs.Cells(7, 8) = Format(FNN(rs.fields("M06_1")), "###,##0")
18150         excWs.Cells(7, 9) = Format(FNN(rs.fields("M07_1")), "###,##0")
18160         excWs.Cells(7, 10) = Format(FNN(rs.fields("M08_1")), "###,##0")
18170         excWs.Cells(7, 11) = Format(FNN(rs.fields("M09_1")), "###,##0")
18180         excWs.Cells(7, 12) = Format(FNN(rs.fields("M10_1")), "###,##0")
18190         excWs.Cells(7, 13) = Format(FNN(rs.fields("M11_1")), "###,##0")
18200         excWs.Cells(7, 14) = Format(FNN(rs.fields("M12_1")), "###,##0")
          
18210         excWs.Cells(8, 3) = Format(FNN(rs.fields("M01_2")), "###,##0")
18220         excWs.Cells(8, 4) = Format(FNN(rs.fields("M02_2")), "###,##0")
18230         excWs.Cells(8, 5) = Format(FNN(rs.fields("M03_2")), "###,##0")
18240         excWs.Cells(8, 6) = Format(FNN(rs.fields("M04_2")), "###,##0")
18250         excWs.Cells(8, 7) = Format(FNN(rs.fields("M05_2")), "###,##0")
18260         excWs.Cells(8, 8) = Format(FNN(rs.fields("M06_2")), "###,##0")
18270         excWs.Cells(8, 9) = Format(FNN(rs.fields("M07_2")), "###,##0")
18280         excWs.Cells(8, 10) = Format(FNN(rs.fields("M08_2")), "###,##0")
18290         excWs.Cells(8, 11) = Format(FNN(rs.fields("M09_2")), "###,##0")
18300         excWs.Cells(8, 12) = Format(FNN(rs.fields("M10_2")), "###,##0")
18310         excWs.Cells(8, 13) = Format(FNN(rs.fields("M11_2")), "###,##0")
18320         excWs.Cells(8, 14) = Format(FNN(rs.fields("M12_2")), "###,##0")
          
18330         excWs.Cells(9, 3) = Format(FNN(rs.fields("M01_3")), "###,##0")
18340         excWs.Cells(9, 4) = Format(FNN(rs.fields("M02_3")), "###,##0")
18350         excWs.Cells(9, 5) = Format(FNN(rs.fields("M03_3")), "###,##0")
18360         excWs.Cells(9, 6) = Format(FNN(rs.fields("M04_3")), "###,##0")
18370         excWs.Cells(9, 7) = Format(FNN(rs.fields("M05_3")), "###,##0")
18380         excWs.Cells(9, 8) = Format(FNN(rs.fields("M06_3")), "###,##0")
18390         excWs.Cells(9, 9) = Format(FNN(rs.fields("M07_3")), "###,##0")
18400         excWs.Cells(9, 10) = Format(FNN(rs.fields("M08_3")), "###,##0")
18410         excWs.Cells(9, 11) = Format(FNN(rs.fields("M09_3")), "###,##0")
18420         excWs.Cells(9, 12) = Format(FNN(rs.fields("M10_3")), "###,##0")
18430         excWs.Cells(9, 13) = Format(FNN(rs.fields("M11_3")), "###,##0")
18440         excWs.Cells(9, 14) = Format(FNN(rs.fields("M12_3")), "###,##0")
          
18450         excWs.Cells(10, 3) = Format(FNN(rs.fields("M01_4")), "###,##0")
18460         excWs.Cells(10, 4) = Format(FNN(rs.fields("M02_4")), "###,##0")
18470         excWs.Cells(10, 5) = Format(FNN(rs.fields("M03_4")), "###,##0")
18480         excWs.Cells(10, 6) = Format(FNN(rs.fields("M04_4")), "###,##0")
18490         excWs.Cells(10, 7) = Format(FNN(rs.fields("M05_4")), "###,##0")
18500         excWs.Cells(10, 8) = Format(FNN(rs.fields("M06_4")), "###,##0")
18510         excWs.Cells(10, 9) = Format(FNN(rs.fields("M07_4")), "###,##0")
18520         excWs.Cells(10, 10) = Format(FNN(rs.fields("M08_4")), "###,##0")
18530         excWs.Cells(10, 11) = Format(FNN(rs.fields("M09_4")), "###,##0")
18540         excWs.Cells(10, 12) = Format(FNN(rs.fields("M10_4")), "###,##0")
18550         excWs.Cells(10, 13) = Format(FNN(rs.fields("M11_4")), "###,##0")
18560         excWs.Cells(10, 14) = Format(FNN(rs.fields("M12_4")), "###,##0")
          
      'Insert nett suppliers invoices here
18570         excWs.Cells(11, 3) = Format(FNN(rs.fields("M01_5b")), "###,##0")
18580         excWs.Cells(11, 4) = Format(FNN(rs.fields("M02_5b")), "###,##0")
18590         excWs.Cells(11, 5) = Format(FNN(rs.fields("M03_5b")), "###,##0")
18600         excWs.Cells(11, 6) = Format(FNN(rs.fields("M04_5b")), "###,##0")
18610         excWs.Cells(11, 7) = Format(FNN(rs.fields("M05_5b")), "###,##0")
18620         excWs.Cells(11, 8) = Format(FNN(rs.fields("M06_5b")), "###,##0")
18630         excWs.Cells(11, 9) = Format(FNN(rs.fields("M07_5b")), "###,##0")
18640         excWs.Cells(11, 10) = Format(FNN(rs.fields("M08_5b")), "###,##0")
18650         excWs.Cells(11, 11) = Format(FNN(rs.fields("M09_5b")), "###,##0")
18660         excWs.Cells(11, 12) = Format(FNN(rs.fields("M10_5b")), "###,##0")
18670         excWs.Cells(11, 13) = Format(FNN(rs.fields("M11_5b")), "###,##0")
18680         excWs.Cells(11, 14) = Format(FNN(rs.fields("M12_5b")), "###,##0")
          
18690         excWs.Cells(12, 3) = Format(FNN(rs.fields("M01_5")), "###,##0")
18700         excWs.Cells(12, 4) = Format(FNN(rs.fields("M02_5")), "###,##0")
18710         excWs.Cells(12, 5) = Format(FNN(rs.fields("M03_5")), "###,##0")
18720         excWs.Cells(12, 6) = Format(FNN(rs.fields("M04_5")), "###,##0")
18730         excWs.Cells(12, 7) = Format(FNN(rs.fields("M05_5")), "###,##0")
18740         excWs.Cells(12, 8) = Format(FNN(rs.fields("M06_5")), "###,##0")
18750         excWs.Cells(12, 9) = Format(FNN(rs.fields("M07_5")), "###,##0")
18760         excWs.Cells(12, 10) = Format(FNN(rs.fields("M08_5")), "###,##0")
18770         excWs.Cells(12, 11) = Format(FNN(rs.fields("M09_5")), "###,##0")
18780         excWs.Cells(12, 12) = Format(FNN(rs.fields("M10_5")), "###,##0")
18790         excWs.Cells(12, 13) = Format(FNN(rs.fields("M11_5")), "###,##0")
18800         excWs.Cells(12, 14) = Format(FNN(rs.fields("M12_5")), "###,##0")
          
18810         excWs.Cells(13, 3) = Format(FNN(rs.fields("M01_6")), "###,##0")
18820         excWs.Cells(13, 4) = Format(FNN(rs.fields("M02_6")), "###,##0")
18830         excWs.Cells(13, 5) = Format(FNN(rs.fields("M03_6")), "###,##0")
18840         excWs.Cells(13, 6) = Format(FNN(rs.fields("M04_6")), "###,##0")
18850         excWs.Cells(13, 7) = Format(FNN(rs.fields("M05_6")), "###,##0")
18860         excWs.Cells(13, 8) = Format(FNN(rs.fields("M06_6")), "###,##0")
18870         excWs.Cells(13, 9) = Format(FNN(rs.fields("M07_6")), "###,##0")
18880         excWs.Cells(13, 10) = Format(FNN(rs.fields("M08_6")), "###,##0")
18890         excWs.Cells(13, 11) = Format(FNN(rs.fields("M09_6")), "###,##0")
18900         excWs.Cells(13, 12) = Format(FNN(rs.fields("M10_6")), "###,##0")
18910         excWs.Cells(13, 13) = Format(FNN(rs.fields("M11_6")), "###,##0")
18920         excWs.Cells(13, 14) = Format(FNN(rs.fields("M12_6")), "###,##0")
          
18930         excWs.Cells(14, 3) = Format(FNDBL(rs.fields("M01_7")), "##0%")
18940         excWs.Cells(14, 4) = Format(FNDBL(rs.fields("M02_7")), "##0%")
18950         excWs.Cells(14, 5) = Format(FNDBL(rs.fields("M03_7")), "##0%")
18960         excWs.Cells(14, 6) = Format(FNDBL(rs.fields("M04_7")), "##0%")
18970         excWs.Cells(14, 7) = Format(FNDBL(rs.fields("M05_7")), "##0%")
18980         excWs.Cells(14, 8) = Format(FNDBL(rs.fields("M06_7")), "##0%")
18990         excWs.Cells(14, 9) = Format(FNDBL(rs.fields("M07_7")), "##0%")
19000         excWs.Cells(14, 10) = Format(FNDBL(rs.fields("M08_7")), "##0%")
19010         excWs.Cells(14, 11) = Format(FNDBL(rs.fields("M09_7")), "##0%")
19020         excWs.Cells(14, 12) = Format(FNDBL(rs.fields("M10_7")), "##0%")
19030         excWs.Cells(14, 13) = Format(FNDBL(rs.fields("M11_7")), "##0%")
19040         excWs.Cells(14, 14) = Format(FNDBL(rs.fields("M12_7")), "##0%")
          
19050         excWs.Cells(15, 3) = Format(FNDBL(rs.fields("M01_8")), "##0%")
19060         excWs.Cells(15, 4) = Format(FNDBL(rs.fields("M02_8")), "###0%")
19070         excWs.Cells(15, 5) = Format(FNDBL(rs.fields("M03_8")), "##0%")
19080         excWs.Cells(15, 6) = Format(FNDBL(rs.fields("M04_8")), "##0%")
19090         excWs.Cells(15, 7) = Format(FNDBL(rs.fields("M05_8")), "##0%")
19100         excWs.Cells(15, 8) = Format(FNDBL(rs.fields("M06_8")), "##0%")
19110         excWs.Cells(15, 9) = Format(FNDBL(rs.fields("M07_8")), "##0%")
19120         excWs.Cells(15, 10) = Format(FNDBL(rs.fields("M08_8")), "##0%")
19130         excWs.Cells(15, 11) = Format(FNDBL(rs.fields("M09_8")), "##0%")
19140         excWs.Cells(15, 12) = Format(FNDBL(rs.fields("M10_8")), "##0%")
19150         excWs.Cells(15, 13) = Format(FNDBL(rs.fields("M11_8")), "##0%")
19160         excWs.Cells(15, 14) = Format(FNDBL(rs.fields("M12_8")), "##0%")
          
19170         excWs.Cells(16, 3) = Format(FNDBL(rs.fields("M01_9")), "##0%")
19180         excWs.Cells(16, 4) = Format(FNDBL(rs.fields("M02_9")), "##0%")
19190         excWs.Cells(16, 5) = Format(FNDBL(rs.fields("M03_9")), "##0%")
19200         excWs.Cells(16, 6) = Format(FNDBL(rs.fields("M04_9")), "##0%")
19210         excWs.Cells(16, 7) = Format(FNDBL(rs.fields("M05_9")), "##0%")
19220         excWs.Cells(16, 8) = Format(FNDBL(rs.fields("M06_9")), "##0%")
19230         excWs.Cells(16, 9) = Format(FNDBL(rs.fields("M07_9")), "##0%")
19240         excWs.Cells(16, 10) = Format(FNDBL(rs.fields("M08_9")), "##0%")
19250         excWs.Cells(16, 11) = Format(FNDBL(rs.fields("M09_9")), "##0%")
19260         excWs.Cells(16, 12) = Format(FNDBL(rs.fields("M10_9")), "##0%")
19270         excWs.Cells(16, 13) = Format(FNDBL(rs.fields("M11_9")), "##0%")
19280         excWs.Cells(16, 14) = Format(FNDBL(rs.fields("M12_9")), "##0%")
              
19290         excWs.Cells(17, 3) = Format(FNDBL(rs.fields("M01_10")), "##0%")
19300         excWs.Cells(17, 4) = Format(FNDBL(rs.fields("M02_10")), "##0%")
19310         excWs.Cells(17, 5) = Format(FNDBL(rs.fields("M03_10")), "##0%")
19320         excWs.Cells(17, 6) = Format(FNDBL(rs.fields("M04_10")), "##0%")
19330         excWs.Cells(17, 7) = Format(FNDBL(rs.fields("M05_10")), "##0%")
19340         excWs.Cells(17, 8) = Format(FNDBL(rs.fields("M06_10")), "##0%")
19350         excWs.Cells(17, 9) = Format(FNDBL(rs.fields("M07_10")), "##0%")
19360         excWs.Cells(17, 10) = Format(FNDBL(rs.fields("M08_10")), "##0%")
19370         excWs.Cells(17, 11) = Format(FNDBL(rs.fields("M09_10")), "##0%")
19380         excWs.Cells(17, 12) = Format(FNDBL(rs.fields("M10_10")), "##0%")
19390         excWs.Cells(17, 13) = Format(FNDBL(rs.fields("M11_10")), "##0%")
19400         excWs.Cells(17, 14) = Format(FNDBL(rs.fields("M12_10")), "##0%")
              
              
19410         excWb.Save
          
              '‘don’t forget to do this or you’ll not be able to open
              '‘book1.xls again, untill you restart you pc.
           '   excApp.ActiveWorkbook.Close False, "c:\budget.xls"
           '   excApp.Quit
19420         Set excWb = Nothing
19430         Set excApp = Nothing
19440         Screen.MousePointer = vbDefault
              
19450     Exit Sub
errHandler:
19460     If ErrMustStop Then Debug.Assert False: Resume
19470     ErrorIn "frmBudgetPreview.ExportToExcel"
End Sub
Public Sub ExportToOO()
19480     On Error GoTo errHandler
      Dim oSM                   'Root object for accessing OpenOffice from VB
      Dim oDesk, oDOC As Object 'First objects from the API
      Dim arg()                 'Ignore it for the moment !
      Dim aNoArgs()
      Dim strFilename As String
      Dim fs As New FileSystemObject
19490  On Error Resume Next
19500         Screen.MousePointer = vbHourglass
      'Instanciate OOo : this line is mandatory with VB for OOo API
19510     Set oSM = CreateObject("com.sun.star.ServiceManager")
      'Create the first and most important service
19520     Set oDesk = oSM.CreateInstance("com.sun.star.frame.Desktop")
      'Open an existing doc (pay attention to the syntax for first argument)
19530     strFilename = Replace(fs.GetDriveName(oPC.SharedFolderRoot), "\", "/") & "/Templates/OO_Budget.ods"
        ' strFilename = Replace(fs.GetDriveName(oPC.LocalFolder), "\", "/") & "/Templates/OO_Budget.ods"
        '  strFilename = "file:///PBKS_S/Templates/OO_Budget.ods"
19540     Set oDOC = oDesk.loadComponentFromURL("file:" & strFilename, "_blank", 0, arg())
      'Save the doc
19550     Set oSheet = oDOC.getSheets().getByIndex(0)
19560         oSheet.getCellRangeByName("A2").setString ("Prepared: " & Format(Now(), "DD-MMM-YYYY HH:NN"))
              
19570         oSheet.getCellRangeByName("C5").Value = (FND(rs.fields("H01")))
19580         oSheet.getCellRangeByName("D5").Value = (FND(rs.fields("H02")))
19590         oSheet.getCellRangeByName("E5").Value = (FND(rs.fields("H03")))
19600         oSheet.getCellRangeByName("F5").Value = (FND(rs.fields("H04")))
19610         oSheet.getCellRangeByName("G5").Value = (FND(rs.fields("H05")))
19620         oSheet.getCellRangeByName("H5").Value = (FND(rs.fields("H06")))
19630         oSheet.getCellRangeByName("I5").Value = (FND(rs.fields("H07")))
19640         oSheet.getCellRangeByName("J5").Value = (FND(rs.fields("H08")))
19650         oSheet.getCellRangeByName("K5").Value = (FND(rs.fields("H09")))
19660         oSheet.getCellRangeByName("L5").Value = (FND(rs.fields("H10")))
19670         oSheet.getCellRangeByName("M5").Value = (FND(rs.fields("H11")))
19680         oSheet.getCellRangeByName("N5").Value = (FND(rs.fields("H12")))
          
19690         oSheet.getCellRangeByName("C6").Value = FNDBL(rs.fields("M01_0"))
19700         oSheet.getCellRangeByName("D6").Value = FNDBL(rs.fields("M02_0"))
19710         oSheet.getCellRangeByName("E6").Value = FNDBL(rs.fields("M03_0"))
19720         oSheet.getCellRangeByName("F6").Value = FNDBL(rs.fields("M04_0"))
19730         oSheet.getCellRangeByName("G6").Value = FNDBL(rs.fields("M05_0"))
19740         oSheet.getCellRangeByName("H6").Value = FNDBL(rs.fields("M06_0"))
19750         oSheet.getCellRangeByName("I6").Value = FNDBL(rs.fields("M07_0"))
19760         oSheet.getCellRangeByName("J6").Value = FNDBL(rs.fields("M08_0"))
19770         oSheet.getCellRangeByName("K6").Value = FNDBL(rs.fields("M09_0"))
19780         oSheet.getCellRangeByName("L6").Value = FNDBL(rs.fields("M10_0"))
19790         oSheet.getCellRangeByName("M6").Value = FNDBL(rs.fields("M11_0"))
19800         oSheet.getCellRangeByName("N6").Value = FNDBL(rs.fields("M12_0"))
              
19810         oSheet.getCellRangeByName("C7").Value = FNDBL(rs.fields("M01_1"))
19820         oSheet.getCellRangeByName("D7").Value = FNDBL(rs.fields("M02_1"))
19830         oSheet.getCellRangeByName("E7").Value = FNDBL(rs.fields("M03_1"))
19840         oSheet.getCellRangeByName("F7").Value = FNDBL(rs.fields("M04_1"))
19850         oSheet.getCellRangeByName("G7").Value = FNDBL(rs.fields("M05_1"))
19860         oSheet.getCellRangeByName("H7").Value = FNDBL(rs.fields("M06_1"))
19870         oSheet.getCellRangeByName("I7").Value = FNDBL(rs.fields("M07_1"))
19880         oSheet.getCellRangeByName("J7").Value = FNDBL(rs.fields("M08_1"))
19890         oSheet.getCellRangeByName("K7").Value = FNDBL(rs.fields("M09_1"))
19900         oSheet.getCellRangeByName("L7").Value = FNDBL(rs.fields("M10_1"))
19910         oSheet.getCellRangeByName("M7").Value = FNDBL(rs.fields("M11_1"))
19920         oSheet.getCellRangeByName("N7").Value = FNDBL(rs.fields("M12_1"))
          
19930         oSheet.getCellRangeByName("C8").Value = FNDBL(rs.fields("M01_2"))
19940         oSheet.getCellRangeByName("D8").Value = FNDBL(rs.fields("M02_2"))
19950         oSheet.getCellRangeByName("E8").Value = FNDBL(rs.fields("M03_2"))
19960         oSheet.getCellRangeByName("F8").Value = FNDBL(rs.fields("M04_2"))
19970         oSheet.getCellRangeByName("G8").Value = FNDBL(rs.fields("M05_2"))
19980         oSheet.getCellRangeByName("H8").Value = FNDBL(rs.fields("M06_2"))
19990         oSheet.getCellRangeByName("I8").Value = FNDBL(rs.fields("M07_2"))
20000         oSheet.getCellRangeByName("J8").Value = FNDBL(rs.fields("M08_2"))
20010         oSheet.getCellRangeByName("K8").Value = FNDBL(rs.fields("M09_2"))
20020         oSheet.getCellRangeByName("L8").Value = FNDBL(rs.fields("M10_2"))
20030         oSheet.getCellRangeByName("M8").Value = FNDBL(rs.fields("M11_2"))
20040         oSheet.getCellRangeByName("N8").Value = FNDBL(rs.fields("M12_2"))
          
20050         oSheet.getCellRangeByName("C9").Value = FNDBL(rs.fields("M01_3"))
20060         oSheet.getCellRangeByName("D9").Value = FNDBL(rs.fields("M02_3"))
20070         oSheet.getCellRangeByName("E9").Value = FNDBL(rs.fields("M03_3"))
20080         oSheet.getCellRangeByName("F9").Value = FNDBL(rs.fields("M04_3"))
20090         oSheet.getCellRangeByName("G9").Value = FNDBL(rs.fields("M05_3"))
20100         oSheet.getCellRangeByName("H9").Value = FNDBL(rs.fields("M06_3"))
20110         oSheet.getCellRangeByName("I9").Value = FNDBL(rs.fields("M07_3"))
20120         oSheet.getCellRangeByName("J9").Value = FNDBL(rs.fields("M08_3"))
20130         oSheet.getCellRangeByName("K9").Value = FNDBL(rs.fields("M09_3"))
20140         oSheet.getCellRangeByName("L9").Value = FNDBL(rs.fields("M10_3"))
20150         oSheet.getCellRangeByName("M9").Value = FNDBL(rs.fields("M11_3"))
20160         oSheet.getCellRangeByName("N9").Value = FNDBL(rs.fields("M12_3"))
          
20170         oSheet.getCellRangeByName("C10").Value = FNDBL(rs.fields("M01_4"))
20180         oSheet.getCellRangeByName("D10").Value = FNDBL(rs.fields("M02_4"))
20190         oSheet.getCellRangeByName("E10").Value = FNDBL(rs.fields("M03_4"))
20200         oSheet.getCellRangeByName("F10").Value = FNDBL(rs.fields("M04_4"))
20210         oSheet.getCellRangeByName("G10").Value = FNDBL(rs.fields("M05_4"))
20220         oSheet.getCellRangeByName("H10").Value = FNDBL(rs.fields("M06_4"))
20230         oSheet.getCellRangeByName("I10").Value = FNDBL(rs.fields("M07_4"))
20240         oSheet.getCellRangeByName("J10").Value = FNDBL(rs.fields("M08_4"))
20250         oSheet.getCellRangeByName("K10").Value = FNDBL(rs.fields("M09_4"))
20260         oSheet.getCellRangeByName("L10").Value = FNDBL(rs.fields("M10_4"))
20270         oSheet.getCellRangeByName("M10").Value = FNDBL(rs.fields("M11_4"))
20280         oSheet.getCellRangeByName("N10").Value = FNDBL(rs.fields("M12_4"))
         
20290         oSheet.getCellRangeByName("C11").Value = FNDBL(rs.fields("M01_5b"))
20300         oSheet.getCellRangeByName("D11").Value = FNDBL(rs.fields("M02_5b"))
20310         oSheet.getCellRangeByName("E11").Value = FNDBL(rs.fields("M03_5b"))
20320         oSheet.getCellRangeByName("F11").Value = FNDBL(rs.fields("M04_5b"))
20330         oSheet.getCellRangeByName("G11").Value = FNDBL(rs.fields("M05_5b"))
20340         oSheet.getCellRangeByName("H11").Value = FNDBL(rs.fields("M06_5b"))
20350         oSheet.getCellRangeByName("I11").Value = FNDBL(rs.fields("M07_5b"))
20360         oSheet.getCellRangeByName("J11").Value = FNDBL(rs.fields("M08_5b"))
20370         oSheet.getCellRangeByName("K11").Value = FNDBL(rs.fields("M09_5b"))
20380         oSheet.getCellRangeByName("L11").Value = FNDBL(rs.fields("M10_5b"))
20390         oSheet.getCellRangeByName("M11").Value = FNDBL(rs.fields("M11_5b"))
20400         oSheet.getCellRangeByName("N11").Value = FNDBL(rs.fields("M12_5b"))
             
20410         oSheet.getCellRangeByName("C12").Value = FNDBL(rs.fields("M01_5"))
20420         oSheet.getCellRangeByName("D12").Value = FNDBL(rs.fields("M02_5"))
20430         oSheet.getCellRangeByName("E12").Value = FNDBL(rs.fields("M03_5"))
20440         oSheet.getCellRangeByName("F12").Value = FNDBL(rs.fields("M04_5"))
20450         oSheet.getCellRangeByName("G12").Value = FNDBL(rs.fields("M05_5"))
20460         oSheet.getCellRangeByName("H12").Value = FNDBL(rs.fields("M06_5"))
20470         oSheet.getCellRangeByName("I12").Value = FNDBL(rs.fields("M07_5"))
20480         oSheet.getCellRangeByName("J12").Value = FNDBL(rs.fields("M08_5"))
20490         oSheet.getCellRangeByName("K12").Value = FNDBL(rs.fields("M09_5"))
20500         oSheet.getCellRangeByName("L12").Value = FNDBL(rs.fields("M10_5"))
20510         oSheet.getCellRangeByName("M12").Value = FNDBL(rs.fields("M11_5"))
20520         oSheet.getCellRangeByName("N12").Value = FNDBL(rs.fields("M12_5"))
          
20530         oSheet.getCellRangeByName("C13").Value = FNDBL(rs.fields("M01_6"))
20540         oSheet.getCellRangeByName("D13").Value = FNDBL(rs.fields("M02_6"))
20550         oSheet.getCellRangeByName("E13").Value = FNDBL(rs.fields("M03_6"))
20560         oSheet.getCellRangeByName("F13").Value = FNDBL(rs.fields("M04_6"))
20570         oSheet.getCellRangeByName("G13").Value = FNDBL(rs.fields("M05_6"))
20580         oSheet.getCellRangeByName("H13").Value = FNDBL(rs.fields("M06_6"))
20590         oSheet.getCellRangeByName("I13").Value = FNDBL(rs.fields("M07_6"))
20600         oSheet.getCellRangeByName("J13").Value = FNDBL(rs.fields("M08_6"))
20610         oSheet.getCellRangeByName("K13").Value = FNDBL(rs.fields("M09_6"))
20620         oSheet.getCellRangeByName("L13").Value = FNDBL(rs.fields("M10_6"))
20630         oSheet.getCellRangeByName("M13").Value = FNDBL(rs.fields("M11_6"))
20640         oSheet.getCellRangeByName("N13").Value = FNDBL(rs.fields("M12_6"))
          
20650         oSheet.getCellRangeByName("C14").Value = FNDBL(rs.fields("M01_7"))
20660         oSheet.getCellRangeByName("D14").Value = FNDBL(rs.fields("M02_7"))
20670         oSheet.getCellRangeByName("E14").Value = FNDBL(rs.fields("M03_7"))
20680         oSheet.getCellRangeByName("F14").Value = FNDBL(rs.fields("M04_7"))
20690         oSheet.getCellRangeByName("G14").Value = FNDBL(rs.fields("M05_7"))
20700         oSheet.getCellRangeByName("H14").Value = FNDBL(rs.fields("M06_7"))
20710         oSheet.getCellRangeByName("I14").Value = FNDBL(rs.fields("M07_7"))
20720         oSheet.getCellRangeByName("J14").Value = FNDBL(rs.fields("M08_7"))
20730         oSheet.getCellRangeByName("K14").Value = FNDBL(rs.fields("M09_7"))
20740         oSheet.getCellRangeByName("L14").Value = FNDBL(rs.fields("M10_7"))
20750         oSheet.getCellRangeByName("M14").Value = FNDBL(rs.fields("M11_7"))
20760         oSheet.getCellRangeByName("N14").Value = FNDBL(rs.fields("M12_7"))
          
20770         oSheet.getCellRangeByName("C15").Value = FNDBL(rs.fields("M01_8"))
20780         oSheet.getCellRangeByName("D15").Value = FNDBL(rs.fields("M02_8"))
20790         oSheet.getCellRangeByName("E15").Value = FNDBL(rs.fields("M03_8"))
20800         oSheet.getCellRangeByName("F15").Value = FNDBL(rs.fields("M04_8"))
20810         oSheet.getCellRangeByName("G15").Value = FNDBL(rs.fields("M05_8"))
20820         oSheet.getCellRangeByName("H15").Value = FNDBL(rs.fields("M06_8"))
20830         oSheet.getCellRangeByName("I15").Value = FNDBL(rs.fields("M07_8"))
20840         oSheet.getCellRangeByName("J15").Value = FNDBL(rs.fields("M08_8"))
20850         oSheet.getCellRangeByName("K15").Value = FNDBL(rs.fields("M09_8"))
20860         oSheet.getCellRangeByName("L15").Value = FNDBL(rs.fields("M10_8"))
20870         oSheet.getCellRangeByName("M15").Value = FNDBL(rs.fields("M11_8"))
20880         oSheet.getCellRangeByName("N15").Value = FNDBL(rs.fields("M12_8"))
          
20890         oSheet.getCellRangeByName("C16").Value = FNDBL(rs.fields("M01_9"))
20900         oSheet.getCellRangeByName("D16").Value = FNDBL(rs.fields("M02_9"))
20910         oSheet.getCellRangeByName("E16").Value = FNDBL(rs.fields("M03_9"))
20920         oSheet.getCellRangeByName("F16").Value = FNDBL(rs.fields("M04_9"))
20930         oSheet.getCellRangeByName("G16").Value = FNDBL(rs.fields("M05_9"))
20940         oSheet.getCellRangeByName("H16").Value = FNDBL(rs.fields("M06_9"))
20950         oSheet.getCellRangeByName("I16").Value = FNDBL(rs.fields("M07_9"))
20960         oSheet.getCellRangeByName("J16").Value = FNDBL(rs.fields("M08_9"))
20970         oSheet.getCellRangeByName("K16").Value = FNDBL(rs.fields("M09_9"))
20980         oSheet.getCellRangeByName("L16").Value = FNDBL(rs.fields("M10_9"))
20990         oSheet.getCellRangeByName("M16").Value = FNDBL(rs.fields("M11_9"))
21000         oSheet.getCellRangeByName("N16").Value = FNDBL(rs.fields("M12_9"))
              
21010         oSheet.getCellRangeByName("C17").Value = FNDBL(rs.fields("M01_10"))
21020         oSheet.getCellRangeByName("D17").Value = FNDBL(rs.fields("M02_10"))
21030         oSheet.getCellRangeByName("E17").Value = FNDBL(rs.fields("M03_10"))
21040         oSheet.getCellRangeByName("F17").Value = FNDBL(rs.fields("M04_10"))
21050         oSheet.getCellRangeByName("G17").Value = FNDBL(rs.fields("M05_10"))
21060         oSheet.getCellRangeByName("H17").Value = FNDBL(rs.fields("M06_10"))
21070         oSheet.getCellRangeByName("I17").Value = FNDBL(rs.fields("M07_10"))
21080         oSheet.getCellRangeByName("J17").Value = FNDBL(rs.fields("M08_10"))
21090         oSheet.getCellRangeByName("K17").Value = FNDBL(rs.fields("M09_10"))
21100         oSheet.getCellRangeByName("L17").Value = FNDBL(rs.fields("M10_10"))
21110         oSheet.getCellRangeByName("M17").Value = FNDBL(rs.fields("M11_10"))
21120         oSheet.getCellRangeByName("N17").Value = FNDBL(rs.fields("M12_10"))
          
21130     oDOC.Store
         ' oDOC.Close (True)
          
21140     Set oDOC = Nothing
21150     Set oDesk = Nothing
         ' oSM.Close
21160     Set oSM = Nothing
21170     Screen.MousePointer = vbDefault

      'errHandler:
      '    ErrorIn "frmBudgetPreview.ExportToOO"
      'errHandler:
      '    ErrorIn "frmBudgetPreview.ExportToOO"
21180     Exit Sub
errHandler:
21190     If ErrMustStop Then Debug.Assert False: Resume
21200     ErrorIn "frmBudgetPreview.ExportToOO"
End Sub
Public Function ConvertToUrl(strFile) As String
21210     On Error GoTo errHandler
      'With thanks to didier.alain@kalitech.fr code used from http://www.kalitech.fr/clients/doc/VB_APIOOo_en.html#replacement_functions

        '  strFile = Replace(strFile, "\", "/")
         ' strFile = Replace(strFile, ":", "|")
21220     strFile = Replace(strFile, " ", "%20")
21230     strFile = "file:///" + strFile
21240     ConvertToUrl = strFile

21250     Exit Function
errHandler:
21260     If ErrMustStop Then Debug.Assert False: Resume
21270     ErrorIn "frmBudgetPreview.ConvertToUrl(strFile)", strFile
End Function
