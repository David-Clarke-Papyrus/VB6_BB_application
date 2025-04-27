VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Check "
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox checkDeleteLine 
      Caption         =   "Secure deleting sale line"
      Height          =   345
      Left            =   525
      TabIndex        =   15
      Top             =   4650
      Width           =   2280
   End
   Begin VB.CheckBox checkCloseApp 
      Caption         =   "Secure closing of application"
      Height          =   345
      Left            =   540
      TabIndex        =   14
      Top             =   4350
      Width           =   3630
   End
   Begin VB.CheckBox CheckPOSDiscount 
      Caption         =   "Secure POS discount"
      Height          =   345
      Left            =   540
      TabIndex        =   13
      Top             =   3990
      Width           =   2280
   End
   Begin VB.CheckBox checkPOSRefund 
      Caption         =   "Secure POS refund"
      Height          =   345
      Left            =   525
      TabIndex        =   12
      Top             =   2775
      Width           =   2280
   End
   Begin VB.CheckBox checkDIrectDeposit 
      Caption         =   "Secure accept direct deposit"
      Height          =   345
      Left            =   525
      TabIndex        =   11
      Top             =   3090
      Width           =   2910
   End
   Begin VB.CheckBox CheckPettyCash 
      Caption         =   "Secure petty cash"
      Height          =   345
      Left            =   525
      TabIndex        =   10
      Top             =   3375
      Width           =   2280
   End
   Begin VB.CheckBox CheckPriceChange 
      Caption         =   "Secure price change"
      Height          =   345
      Left            =   525
      TabIndex        =   9
      Top             =   3690
      Width           =   2280
   End
   Begin VB.CheckBox CheckIssueAppro 
      Caption         =   "Secure Issue appro"
      Height          =   345
      Left            =   525
      TabIndex        =   8
      Top             =   1560
      Width           =   2280
   End
   Begin VB.CheckBox checkRefundDeposit 
      Caption         =   "Secure Refund deposit"
      Height          =   345
      Left            =   525
      TabIndex        =   7
      Top             =   1860
      Width           =   2280
   End
   Begin VB.CheckBox CheckAcPayment 
      Caption         =   "Secure Accept a/c payment"
      Height          =   345
      Left            =   525
      TabIndex        =   6
      Top             =   2160
      Width           =   2925
   End
   Begin VB.CheckBox checkPOSCN 
      Caption         =   "SecurePOSCreditNote"
      Height          =   345
      Left            =   525
      TabIndex        =   5
      Top             =   2475
      Width           =   2280
   End
   Begin VB.CheckBox CheckVoid 
      Caption         =   "Secure Void"
      Height          =   345
      Left            =   525
      TabIndex        =   4
      Top             =   315
      Width           =   2280
   End
   Begin VB.CheckBox CheckOpenDrawer 
      Caption         =   "Secure Open Drawer"
      Height          =   345
      Left            =   525
      TabIndex        =   3
      Top             =   630
      Width           =   2280
   End
   Begin VB.CheckBox CheckDiscount 
      Caption         =   "Secure Discount"
      Height          =   345
      Left            =   525
      TabIndex        =   2
      Top             =   915
      Width           =   2280
   End
   Begin VB.CheckBox CheckCreditNote 
      Caption         =   "Secure Credit Note"
      Height          =   345
      Left            =   525
      TabIndex        =   1
      Top             =   1230
      Width           =   2280
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "GetResult"
      Height          =   345
      Left            =   465
      TabIndex        =   0
      Top             =   5145
      Width           =   1065
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmp As Long

Const M_VOID As Integer = 1
Const M_OPENDRAWER As Integer = 2
Const M_DISCOUNT As Integer = 4
Const M_CREDITNOTE As Integer = 8
Const M_ISSUEAPPRO As Integer = 16
Const M_REFUNDDEPOSIT As Integer = 32
Const M_ACCEPTACPAYMENT As Integer = 64
Const M_ISSUEPOSCREDITNOTE As Integer = 128
Const M_ISSUEPOSREFUND As Integer = 256
Const M_ACCEPTDIRECTDEPOSIT As Integer = 512
Const M_PETTYCASH As Integer = 1024
Const M_POSPRICECHANGE As Integer = 2048
Const M_POSDISCOUNT As Integer = 4096
Const M_CLOSEAPPLICATION As Integer = 8192
Const M_DELETELINE As Integer = 16384
Private Sub cmdSet_Click()
    
    tmp = (M_VOID * CheckVoid) Or _
            (M_OPENDRAWER * CheckOpenDrawer) Or _
            (M_DISCOUNT * CheckDiscount) Or _
            (M_CREDITNOTE * CheckCreditNote) Or _
            (M_ISSUEAPPRO * CheckIssueAppro) Or _
            (M_REFUNDDEPOSIT * checkRefundDeposit) Or _
            (M_ACCEPTACPAYMENT * CheckAcPayment) Or _
            (M_ISSUEPOSCREDITNOTE * checkPOSCN) Or _
            (M_ISSUEPOSREFUND * checkPOSRefund) Or _
            (M_ACCEPTDIRECTDEPOSIT * checkDIrectDeposit) Or _
            (M_PETTYCASH * CheckPettyCash) Or _
            (M_POSPRICECHANGE * CheckPriceChange) Or _
            (M_POSDISCOUNT * CheckPOSDiscount) Or _
            (M_CLOSEAPPLICATION * checkCloseApp) Or _
            (M_DELETELINE * checkDeleteLine)
            
End Sub


Private Sub cmdGet_Click()
cmdSet_Click
MsgBox tmp
End Sub
