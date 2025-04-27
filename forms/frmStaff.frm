VERSION 5.00
Begin VB.Form frmStaff 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Staff member"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   10680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDefaultsStoreManager 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Go"
      Height          =   375
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   1185
      Width           =   480
   End
   Begin VB.CommandButton cmdDefaultSupervisor 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Go"
      Height          =   375
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   660
      Width           =   480
   End
   Begin VB.CommandButton cmdDefaultOperator 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Go"
      Height          =   375
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   150
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5340
      Picture         =   "frmStaff.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6810
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
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
      Height          =   615
      Left            =   4320
      Picture         =   "frmStaff.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6810
      Width           =   1000
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   810
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   870
      Width           =   2115
   End
   Begin VB.TextBox txtSignature 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4020
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   870
      Width           =   2115
   End
   Begin VB.CommandButton cmdComm 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Commissions"
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
      Height          =   435
      Left            =   4605
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1455
      Width           =   1305
   End
   Begin VB.CheckBox chkRep 
      BackColor       =   &H00D3D3CB&
      Caption         =   "is a rep"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   780
      TabIndex        =   20
      Top             =   1545
      Width           =   1260
   End
   Begin VB.CheckBox chkOp 
      BackColor       =   &H00D3D3CB&
      Caption         =   "is an operator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2760
      TabIndex        =   19
      Top             =   1530
      Width           =   1815
   End
   Begin VB.Frame Permissions 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Permissions"
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
      Height          =   4755
      Left            =   165
      TabIndex        =   18
      Top             =   1965
      Width           =   9450
      Begin VB.CheckBox chkSaleLineDelete 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Delete sale line"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3435
         TabIndex        =   72
         Top             =   4245
         Width           =   2325
      End
      Begin VB.CheckBox chkNewStockItems 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Create stock items"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6345
         TabIndex        =   71
         Top             =   3900
         Width           =   2730
      End
      Begin VB.CheckBox chkIssuePettyCash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Issue petty cash"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3435
         TabIndex        =   70
         Top             =   3600
         Width           =   2325
      End
      Begin VB.CheckBox chkOpenDrawer 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Open drawer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3435
         TabIndex        =   69
         Top             =   3300
         Width           =   2325
      End
      Begin VB.CheckBox chkDebtors 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Block/unblock debtors"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6345
         TabIndex        =   68
         Top             =   3570
         Width           =   2730
      End
      Begin VB.CheckBox chkQU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Sign quotation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   240
         TabIndex        =   59
         Top             =   4245
         Width           =   2565
      End
      Begin VB.CheckBox chkSTKADJ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Sign stock adjustment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   240
         TabIndex        =   58
         Top             =   3900
         Width           =   2565
      End
      Begin VB.CheckBox chkINV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Sign invoice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   240
         TabIndex        =   57
         Top             =   3570
         Width           =   2565
      End
      Begin VB.CheckBox chkRET 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Sign return finalise"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   240
         TabIndex        =   56
         Top             =   3225
         Width           =   2565
      End
      Begin VB.CheckBox chkRETREQ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Sign return request"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   240
         TabIndex        =   55
         Top             =   2895
         Width           =   2565
      End
      Begin VB.CheckBox chkCN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Sign credit note"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   240
         TabIndex        =   54
         Top             =   2565
         Width           =   2565
      End
      Begin VB.CheckBox chkTFR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Sign transfer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   240
         TabIndex        =   53
         Top             =   2235
         Width           =   2565
      End
      Begin VB.CheckBox chkAPPR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Sign appro return"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   240
         TabIndex        =   52
         Top             =   1890
         Width           =   2565
      End
      Begin VB.CheckBox chkAPP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Sign appro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   240
         TabIndex        =   51
         Top             =   1560
         Width           =   2565
      End
      Begin VB.CheckBox chkGRN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Sign G.R.N."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   240
         TabIndex        =   50
         Top             =   1230
         Width           =   2565
      End
      Begin VB.CheckBox chkCO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Sign C.O"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   240
         TabIndex        =   49
         Top             =   885
         Width           =   2565
      End
      Begin VB.CheckBox chkPO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Sign P.O."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   240
         TabIndex        =   48
         Top             =   555
         Width           =   2565
      End
      Begin VB.CheckBox chkPOSCN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Issue POS credit note"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3435
         TabIndex        =   46
         Top             =   2385
         Width           =   2325
      End
      Begin VB.CheckBox chkTakePayment 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Accept a/c payment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3435
         TabIndex        =   45
         Top             =   2995
         Width           =   2325
      End
      Begin VB.CheckBox chkPOSDepositRefund 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Refund deposit at POS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3435
         TabIndex        =   44
         Top             =   2690
         Width           =   2325
      End
      Begin VB.CheckBox chkDICT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Edit dictionary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6345
         TabIndex        =   43
         Top             =   1887
         Width           =   2730
      End
      Begin VB.CheckBox chkCONFIG 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Edit configuration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6345
         TabIndex        =   42
         Top             =   1554
         Width           =   2730
      End
      Begin VB.CheckBox chkCASHUP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Open cash-up sheet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3435
         TabIndex        =   41
         Top             =   3930
         Width           =   2325
      End
      Begin VB.CheckBox chkPT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Product types/categories"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6345
         TabIndex        =   40
         Top             =   2220
         Width           =   2730
      End
      Begin VB.CheckBox chkRR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Edit rounding rules"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6345
         TabIndex        =   39
         Top             =   2553
         Width           =   2730
      End
      Begin VB.CheckBox chkDD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Accept direct deposit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3435
         TabIndex        =   38
         Top             =   555
         Width           =   2325
      End
      Begin VB.CheckBox chkPOSRefund 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Issue refund"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3435
         TabIndex        =   37
         Top             =   2080
         Width           =   2325
      End
      Begin VB.CheckBox chkPriceChange 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Change price on POS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3435
         TabIndex        =   36
         Top             =   1775
         Width           =   2325
      End
      Begin VB.CheckBox chkPOSAppro 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Issue appro on POS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3435
         TabIndex        =   35
         Top             =   1470
         Width           =   2325
      End
      Begin VB.CheckBox chkPOSDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Give POS discount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3435
         TabIndex        =   34
         Top             =   1165
         Width           =   2325
      End
      Begin VB.CheckBox chkVoid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Void exchange"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3435
         TabIndex        =   33
         Top             =   860
         Width           =   2325
      End
      Begin VB.CheckBox chkMB 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Alter multi-buy allocations"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6345
         TabIndex        =   32
         Top             =   2886
         Width           =   2730
      End
      Begin VB.CheckBox chkCustDiscounts 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Prices && discounts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6345
         TabIndex        =   25
         Top             =   1221
         Width           =   2730
      End
      Begin VB.CheckBox chkMergeTP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Merge custs. and suppls."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6345
         TabIndex        =   24
         Top             =   540
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.CheckBox chkMergeStock 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Merge stock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6345
         TabIndex        =   23
         Top             =   888
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.CheckBox chkComm 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "View commissions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6345
         TabIndex        =   22
         Top             =   3225
         Width           =   2730
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   7065
         TabIndex        =   61
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Signing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   765
         TabIndex        =   60
         Top             =   315
         Width           =   1200
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Point of sale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   3840
         TabIndex        =   47
         Top             =   315
         Width           =   1200
      End
   End
   Begin VB.CheckBox chkSupervisor 
      BackColor       =   &H00D3D3CB&
      Caption         =   "is a supervisor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   780
      TabIndex        =   17
      Top             =   1260
      Width           =   1725
   End
   Begin VB.TextBox txtShortName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5325
      TabIndex        =   13
      Top             =   75
      Width           =   795
   End
   Begin VB.CheckBox chkActive 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Active"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2760
      TabIndex        =   12
      Top             =   1260
      Width           =   870
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Access level"
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
      Height          =   885
      Left            =   510
      TabIndex        =   8
      Top             =   7980
      Visible         =   0   'False
      Width           =   4320
      Begin VB.OptionButton Option4 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Level 4"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   3300
         TabIndex        =   15
         Top             =   390
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Level 3"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   2260
         TabIndex        =   11
         Top             =   375
         Width           =   900
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Level 2"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   1220
         TabIndex        =   10
         Top             =   390
         Width           =   900
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Level 1"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   180
         TabIndex        =   9
         Top             =   390
         Width           =   900
      End
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2190
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   7065
      Width           =   1095
   End
   Begin VB.TextBox txtCell 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4005
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   2115
   End
   Begin VB.TextBox txtTel 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   795
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   2115
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   810
      TabIndex        =   0
      Top             =   90
      Width           =   2055
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Set defaults for store manager"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   6300
      TabIndex        =   67
      Top             =   1245
      Width           =   2670
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Set defaults for supervisor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   6525
      TabIndex        =   65
      Top             =   720
      Width           =   2445
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Set defaults for operator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   6525
      TabIndex        =   63
      Top             =   210
      Width           =   2445
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   210
      TabIndex        =   29
      Top             =   900
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Signature"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2955
      TabIndex        =   28
      Top             =   900
      Width           =   960
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   720
      Left            =   2535
      TabIndex        =   16
      Top             =   7035
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Short name (max 4 chars)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2940
      TabIndex        =   14
      Top             =   120
      Width           =   2280
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   1140
      TabIndex        =   7
      Top             =   7080
      Width           =   945
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   3270
      TabIndex        =   5
      Top             =   510
      Width           =   630
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   315
      TabIndex        =   3
      Top             =   510
      Width           =   360
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   135
      Width           =   600
   End
End
Attribute VB_Name = "frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents oStaff As a_Staff
Attribute oStaff.VB_VarHelpID = -1
Dim flgLoading As Boolean

Private Sub chkNewStockItems_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_CREATENEWSTOCKITEM, chkNewStockItems = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkNewStockItems_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdComm_Click()
    On Error GoTo errHandler
Dim frm As New frmCR
    
    frm.LoadForSM oStaff.ID
    frm.Show vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.cmdComm_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oStaff_SupervisorstatusChange()
    On Error GoTo errHandler
    cmdComm.Enabled = oStaff.IsRep
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.oStaff_SupervisorstatusChange", , EA_NORERAISE
    HandleError
End Sub
Private Sub oStaff_RepstatusChange()
    On Error GoTo errHandler
    cmdComm.Enabled = oStaff.IsRep
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.oStaff_RepstatusChange", , EA_NORERAISE
    HandleError
End Sub
'==============================================
Private Sub chkDD_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_ACCEPTDIRECTDEPOSIT, chkDD = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkDD_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkVoid_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_VOIDEXCHANGE, chkVoid = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkVoid_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkPOSAppro_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_ISSUEPOSAPPRO, chkPOSAppro = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkPOSAppro_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkPOSCN_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_ISSUEPOSCREDITNOTE, chkPOSCN = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkPOSCN_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkPOSDepositRefund_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_REFUNDDEPOSITONPOS, chkPOSDepositRefund = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkPOSDepositRefund_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkPOSDiscount_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_POSDISCOUNT, chkPOSDiscount = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkPOSDiscount_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkPOSRefund_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_ISSUEPOSREFUND, chkPOSRefund = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkPOSRefund_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkPriceChange_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_POSPRICECHANGE, chkPriceChange = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkPriceChange_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkTakePayment_Click()
    oStaff.SetRole enSECURITY_ACCEPTACPAYMENT, chkTakePayment = 1
End Sub

Private Sub chkOpenDrawer_Click()
    oStaff.SetRole enSECURITY_OPENDRAWER, chkOpenDrawer = 1
End Sub
Private Sub chkIssuePettyCash_Click()
    oStaff.SetRole enSECURITY_ISSUEPETTYCASH, chkIssuePettyCash = 1
End Sub

Private Sub chkMB_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_MULTIBUY, chkMB = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkMB_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkActive_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_ACTIVE, chkActive = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkActive_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkAPP_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_APP_SIGN, chkAPP = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkAPP_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkAPPR_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_APPR_SIGN, chkAPPR = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkAPPR_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkCASHUP_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_CASHUP_SIGN, chkCASHUP = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkCASHUP_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkSaleLineDelete_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_SALELINEDELETE, chkCASHUP = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkSaleLineDelete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkCN_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_CN_SIGN, chkCN = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkCN_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkCO_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_CO_SIGN, chKCO = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkCO_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkCONFIG_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_CONFIG_SIGN, chkCONFIG = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkCONFIG_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkDICT_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_DICT_SIGN, chkDICT = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkDICT_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkGRN_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_GRN_SIGN, chkGRN = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkGRN_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkINV_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_INV_SIGN, chkINV = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkINV_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkOp_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_ISOPERATOR, chkOp = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkOp_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkPO_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_PO_SIGN, chkPO = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkPO_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkRep_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_ISREP, chkRep = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkRep_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkRET_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_RETFIN_SIGN, chkRET = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkRET_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkRETREQ_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_RETREQ_SIGN, chkRETREQ = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkRETREQ_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkSTKADJ_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_STKADJ_SIGN, chkSTKADJ = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkSTKADJ_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkSupervisor_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_ISSUPERVISOR, chkSupervisor = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkSupervisor_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkTFR_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_TFR_SIGN, chkTFR = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkTFR_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkComm_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_COMM_AUTH, chkComm = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkComm_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkDebtors_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_BLOCK_DEBTORS, chkDebtors = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkDebtors_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkCustDiscounts_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_CUSTDISCOUNT_AUTH, chkCustDiscounts = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkCustDiscounts_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkPT_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_EDITPRODUCTTYPES_AUTH, chkPT = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkPT_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkRR_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_EDITROUNDINGRULES_AUTH, chkRR = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkRR_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkMergeStock_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_MERGESTOCK_AUTH, chkMergeStock = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkMergeStock_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkMergeTP_Click()
    On Error GoTo errHandler
    oStaff.SetRole enSECURITY_MERGECUSTSUPP_AUTH, chkMergeTP = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.chkMergeTP_Click", , EA_NORERAISE
    HandleError
End Sub

'========================
Private Sub oStaff_Valid(pMsg As String)
    On Error GoTo errHandler
    EnableOK pMsg = ""
    lblErrors = pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.oStaff_Valid(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub

Private Sub EnableOK(pOK As Boolean)
    On Error GoTo errHandler
    Me.cmdOK.Enabled = pOK
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.EnableOK(pOK)", pOK
End Sub

Public Sub component(poStaff As a_Staff)
    On Error GoTo errHandler
    Set oStaff = poStaff
    'oStaff.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.component(poStaff)", poStaff
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    
    txtName = oStaff.StaffName
    txtShortname = oStaff.Shortname
    txtTel = oStaff.StaffTel
    txtCell = oStaff.StaffCell
    txtPassword = oStaff.Password
    txtEmail = oStaff.EMail
    txtSignature = oStaff.Signature
    cmdComm.Enabled = oStaff.IsRep
    
    chkActive = IIf(oStaff.GetRole(enSECURITY_ACTIVE), 1, 0)
    chkSupervisor = IIf(oStaff.GetRole(enSECURITY_ISSUPERVISOR), 1, 0)
    chkOp = IIf(oStaff.GetRole(enSECURITY_ISOPERATOR), 1, 0)
    chkRep = IIf(oStaff.GetRole(enSECURITY_ISREP), 1, 0)
    
    
    chkAPP = IIf(oStaff.GetRole(enSECURITY_APP_SIGN), 1, 0)
    chkAPPR = IIf(oStaff.GetRole(enSECURITY_APPR_SIGN), 1, 0)
    chkCASHUP = IIf(oStaff.GetRole(enSECURITY_CASHUP_SIGN), 1, 0)
    chkSaleLineDelete = IIf(oStaff.GetRole(enSECURITY_SALELINEDELETE), 1, 0)
    chkCN = IIf(oStaff.GetRole(enSECURITY_CN_SIGN), 1, 0)
    chKCO = IIf(oStaff.GetRole(enSECURITY_CO_SIGN), 1, 0)
    chkCONFIG = IIf(oStaff.GetRole(enSECURITY_CONFIG_SIGN), 1, 0)
    chkDICT = IIf(oStaff.GetRole(enSECURITY_DICT_SIGN), 1, 0)
    chkGRN = IIf(oStaff.GetRole(enSECURITY_GRN_SIGN), 1, 0)
    chkINV = IIf(oStaff.GetRole(enSECURITY_INV_SIGN), 1, 0)
    chkPO = IIf(oStaff.GetRole(enSECURITY_PO_SIGN), 1, 0)
    chkRET = IIf(oStaff.GetRole(enSECURITY_RETFIN_SIGN), 1, 0)
    chkTFR = IIf(oStaff.GetRole(enSECURITY_TFR_SIGN), 1, 0)
    chkSTKADJ = IIf(oStaff.GetRole(enSECURITY_STKADJ_SIGN), 1, 0)
    chkRETREQ = IIf(oStaff.GetRole(enSECURITY_RETREQ_SIGN), 1, 0)
    
    chkComm = IIf(oStaff.GetRole(enSECURITY_COMM_AUTH), 1, 0)
    chkCustDiscounts = IIf(oStaff.GetRole(enSECURITY_CUSTDISCOUNT_AUTH), 1, 0)
    chkPT = IIf(oStaff.GetRole(enSECURITY_EDITPRODUCTTYPES_AUTH), 1, 0)
    chkRR = IIf(oStaff.GetRole(enSECURITY_EDITROUNDINGRULES_AUTH), 1, 0)
    chkMergeStock = IIf(oStaff.GetRole(enSECURITY_MERGESTOCK_AUTH), 1, 0)
    chkMergeTP = IIf(oStaff.GetRole(enSECURITY_MERGECUSTSUPP_AUTH), 1, 0)
    chkMB = IIf(oStaff.GetRole(enSECURITY_MULTIBUY), 1, 0)
    chkDD = IIf(oStaff.GetRole(enSECURITY_ACCEPTDIRECTDEPOSIT), 1, 0)
    
    chkVoid = IIf(oStaff.GetRole(enSECURITY_VOIDEXCHANGE), 1, 0)
    chkPOSDiscount = IIf(oStaff.GetRole(enSECURITY_POSDISCOUNT), 1, 0)
    chkPOSAppro = IIf(oStaff.GetRole(enSECURITY_ISSUEPOSAPPRO), 1, 0)
    chkPOSRefund = IIf(oStaff.GetRole(enSECURITY_ISSUEPOSREFUND), 1, 0)
    chkPOSCN = IIf(oStaff.GetRole(enSECURITY_ISSUEPOSCREDITNOTE), 1, 0)
    chkPriceChange = IIf(oStaff.GetRole(enSECURITY_POSPRICECHANGE), 1, 0)
    chkPOSDepositRefund = IIf(oStaff.GetRole(enSECURITY_REFUNDDEPOSITONPOS), 1, 0)
    chkTakePayment = IIf(oStaff.GetRole(enSECURITY_ACCEPTACPAYMENT), 1, 0)
    chkQU = IIf(oStaff.GetRole(enSECURITY_QU_SIGN), 1, 0)
    chkDebtors = IIf(oStaff.GetRole(enSECURITY_BLOCK_DEBTORS), 1, 0)
    chkOpenDrawer = IIf(oStaff.GetRole(enSECURITY_OPENDRAWER), 1, 0)
    chkNewStockItems = IIf(oStaff.GetRole(enSECURITY_CREATENEWSTOCKITEM), 1, 0)
    Me.chkIssuePettyCash = IIf(oStaff.GetRole(enSECURITY_ISSUEPETTYCASH), 1, 0)
    If oPC.Configuration.Staff.IsLastSupervisor(oStaff) Then
        Me.Frame1.Enabled = False
    End If
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.LoadControls"
End Sub
Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oStaff.CancelEdit
    oStaff.BeginEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdDefaultOperator_Click()
    chkPO = 0
    chKCO = 0
    chkGRN = 0
    chkAPP = 0
    chkAPPR = 0
    chkTFR = 0
    chkCN = 0
    chkRETREQ = 0
    chkRET = 0
    chkINV = 0
    chkSTKADJ = 0
    chkQU = 0
    chkDD = 0
    chkVoid = 0
    chkPOSDiscount = 0
    chkPOSAppro = 0
    chkPOSRefund = 0
    chkPOSCN = 0
    chkPOSDepositRefund = 0
    chkTakePayment = 0
    chkCASHUP = 0
    chkMergeTP = 0
    chkMergeStock = 0
    chkCustDiscounts = 0
    chkCONFIG = 0
    chkDICT = 0
    chkPT = 0
    chkRR = 0
    chkMB = 0
    chkComm = 0
    chkOp = 1
    chkActive = 1
    chkSupervisor = 0
    chkRep = 0
End Sub

Private Sub cmdDefaultsStoreManager_Click()
    chkPO = 1
    chKCO = 1
    chkGRN = 1
    chkAPP = 1
    chkAPPR = 1
    chkTFR = 1
    chkCN = 1
    chkRETREQ = 1
    chkRET = 1
    chkINV = 1
    chkSTKADJ = 1
    chkQU = 1
    chkDD = 1
    chkVoid = 1
    chkPOSDiscount = 1
    chkPOSAppro = 1
    chkPOSRefund = 1
    chkPOSCN = 1
    chkPOSDepositRefund = 1
    chkTakePayment = 1
    chkCASHUP = 1
    chkMergeTP = 1
    chkMergeStock = 1
    chkCustDiscounts = 1
    chkCONFIG = 1
    chkDICT = 1
    chkPT = 1
    chkRR = 1
    chkMB = 1
    chkComm = 1
    chkOp = 1
    chkActive = 1
    chkSupervisor = 1
    chkRep = 0
End Sub

Private Sub cmdDefaultSUpervisor_Click()
    chkPO = 1
    chKCO = 1
    chkGRN = 1
    chkAPP = 1
    chkAPPR = 1
    chkTFR = 1
    chkCN = 1
    chkRETREQ = 1
    chkRET = 1
    chkINV = 1
    chkSTKADJ = 1
    chkQU = 1
    chkDD = 1
    chkVoid = 1
    chkPOSDiscount = 1
    chkPOSAppro = 1
    chkPOSRefund = 1
    chkPOSCN = 1
    chkPOSDepositRefund = 1
    chkTakePayment = 1
    chkCASHUP = 1
    chkMergeTP = 1
    chkMergeStock = 1
    chkCustDiscounts = 1
    chkCONFIG = 0
    chkDICT = 1
    chkPT = 1
    chkRR = 1
    chkMB = 1
    chkComm = 0
    chkOp = 1
    chkActive = 1
    chkSupervisor = 1
    chkRep = 0
    
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long

    oStaff.ApplyEdit
    oStaff.BeginEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtName_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oStaff.SetStaffName txtName
    If Err Then
      Beep
      intPos = txtName.SelStart
      txtName = oStaff.StaffName
      txtName.SelStart = intPos - 1
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.txtName_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtName_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtName")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.txtName_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtName_LostFocus()
    On Error GoTo errHandler
   txtName.text = oStaff.StaffName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.txtName_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCell_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oStaff.SetStaffCell txtCell
    If Err Then
      Beep
      intPos = txtCell.SelStart
      txtCell = oStaff.StaffCell
      txtCell.SelStart = intPos - 1
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.txtCell_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCell_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtCell")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.txtCell_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCell_LostFocus()
    On Error GoTo errHandler
   txtCell.text = oStaff.StaffCell
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.txtCell_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTel_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
     On Error Resume Next
   oStaff.SetStaffTel txtTel
    If Err Then
      Beep
      intPos = txtTel.SelStart
      txtTel = oStaff.StaffTel
      txtTel.SelStart = intPos - 1
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.txtTel_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTel_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtTel")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.txtTel_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTel_LostFocus()
    On Error GoTo errHandler
   txtTel.text = oStaff.StaffTel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.txtTel_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPassword_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oStaff.SetPassword txtPassword
    If Err Then
      Beep
      intPos = txtTel.SelStart
      txtPassword = oStaff.Password
      txtPassword.SelStart = intPos - 1
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.txtPassword_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEmail_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
     On Error Resume Next
   oStaff.SetEmail txtEmail
    If Err Then
      Beep
      intPos = txtEmail.SelStart
      txtEmail = oStaff.EMail
      txtEmail.SelStart = intPos - 1
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.txtEmail_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtSignature_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oStaff.SetSignature txtSignature
    If Err Then
      Beep
      intPos = txtSignature.SelStart
      txtSignature = oStaff.Signature
      txtSignature.SelStart = intPos - 1
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.txtSignature_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtShortName_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oStaff.SetShortname txtShortname
    If Err Then
      Beep
      intPos = txtShortname.SelStart
      txtShortname = oStaff.Shortname
      txtShortname.SelStart = intPos - 1
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.txtShortName_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtShortName_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtShortName")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.txtShortName_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtShortName_LostFocus()
    On Error GoTo errHandler
   txtShortname.text = oStaff.Shortname
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmStaff.txtShortName_LostFocus", , EA_NORERAISE
    HandleError
End Sub

