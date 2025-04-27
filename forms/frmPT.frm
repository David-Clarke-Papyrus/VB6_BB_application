VERSION 5.00
Begin VB.Form frmPT 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Product type"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   8805
   Begin VB.TextBox txtNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4755
      MaxLength       =   10
      TabIndex        =   56
      Top             =   120
      Width           =   765
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00D3D3CB&
      Caption         =   "General ledger accounts"
      ForeColor       =   &H8000000D&
      Height          =   1860
      Left            =   180
      TabIndex        =   37
      Top             =   4305
      Width           =   8460
      Begin VB.CommandButton cmdCopy 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Copy from other &product type"
         Height          =   615
         Left            =   6390
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1050
         Width           =   1935
      End
      Begin VB.TextBox txtPURCHASES_CONTRA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4695
         TabIndex        =   48
         Top             =   1260
         Width           =   1425
      End
      Begin VB.TextBox txtCASHSALES_CONTRA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2595
         TabIndex        =   47
         Top             =   1260
         Width           =   1425
      End
      Begin VB.TextBox txtCRSALES_CONTRA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   585
         TabIndex        =   46
         Top             =   1260
         Width           =   1425
      End
      Begin VB.TextBox txtVAT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6450
         TabIndex        =   45
         Top             =   630
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox txtPURCHASES 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4695
         TabIndex        =   40
         Top             =   630
         Width           =   1425
      End
      Begin VB.TextBox txtCASHSALES 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2595
         TabIndex        =   39
         Top             =   630
         Width           =   1425
      End
      Begin VB.TextBox txtCRSALES 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   585
         TabIndex        =   38
         Top             =   630
         Width           =   1425
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CR"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   4365
         TabIndex        =   55
         Top             =   1290
         Width           =   285
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CR"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   2265
         TabIndex        =   54
         Top             =   1290
         Width           =   285
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CR"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   255
         TabIndex        =   53
         Top             =   1260
         Width           =   285
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DR"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   4365
         TabIndex        =   52
         Top             =   660
         Width           =   285
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DR"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   2265
         TabIndex        =   51
         Top             =   660
         Width           =   285
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DR"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   255
         TabIndex        =   50
         Top             =   645
         Width           =   285
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "VAT"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   6555
         TabIndex        =   44
         Top             =   345
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit sales"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   180
         TabIndex        =   43
         Top             =   375
         Width           =   1215
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash sales"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   2310
         TabIndex        =   42
         Top             =   375
         Width           =   1050
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Purchases"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   4380
         TabIndex        =   41
         Top             =   375
         Width           =   1170
      End
   End
   Begin VB.CheckBox chkVoucher 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Is a voucher"
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   5970
      TabIndex        =   36
      Top             =   705
      Width           =   1560
   End
   Begin VB.CommandButton cmdComm 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Commissions"
      Default         =   -1  'True
      Height          =   435
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6345
      Width           =   1305
   End
   Begin VB.CheckBox chkActive 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Active"
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   5970
      TabIndex        =   2
      Top             =   75
      Width           =   1560
   End
   Begin VB.TextBox txtMarkdown 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3345
      TabIndex        =   4
      Top             =   1020
      Width           =   870
   End
   Begin VB.CheckBox chkSaleOrReturn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Sale or return"
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   5970
      TabIndex        =   3
      Top             =   390
      Width           =   1560
   End
   Begin VB.TextBox txtError 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   1635
      MultiLine       =   -1  'True
      TabIndex        =   33
      Top             =   6285
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Band 2  (Prices in whole currency units e.g. rands)"
      ForeColor       =   &H8000000D&
      Height          =   840
      Left            =   165
      TabIndex        =   28
      Top             =   2310
      Width           =   8460
      Begin VB.TextBox txtB2Min 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1725
         TabIndex        =   8
         Top             =   345
         Width           =   870
      End
      Begin VB.TextBox txtB2Max 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3165
         TabIndex        =   9
         Top             =   345
         Width           =   1095
      End
      Begin VB.TextBox txtB2MU 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5310
         TabIndex        =   10
         Top             =   345
         Width           =   870
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Add"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   4755
         TabIndex        =   32
         Top             =   375
         Width           =   390
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "and"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   2700
         TabIndex        =   31
         Top             =   375
         Width           =   390
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Price between"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   180
         TabIndex        =   30
         Top             =   375
         Width           =   1365
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "currency units to price."
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   6285
         TabIndex        =   29
         Top             =   375
         Width           =   2115
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Band 3  (Prices in whole currency units e.g. rands)"
      ForeColor       =   &H8000000D&
      Height          =   840
      Left            =   165
      TabIndex        =   23
      Top             =   3225
      Width           =   8460
      Begin VB.TextBox txtB3Min 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1710
         TabIndex        =   11
         Top             =   345
         Width           =   870
      End
      Begin VB.TextBox txtB3Max 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3180
         TabIndex        =   12
         Top             =   345
         Width           =   1095
      End
      Begin VB.TextBox txtB3MU 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5310
         TabIndex        =   13
         Top             =   345
         Width           =   870
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Add"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   4755
         TabIndex        =   27
         Top             =   375
         Width           =   390
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "and"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   2700
         TabIndex        =   26
         Top             =   375
         Width           =   390
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Price between"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   195
         TabIndex        =   25
         Top             =   345
         Width           =   1365
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "currency units to price."
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   6285
         TabIndex        =   24
         Top             =   375
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Band 1 (Prices in whole currency units e.g. rands)"
      ForeColor       =   &H8000000D&
      Height          =   840
      Left            =   165
      TabIndex        =   18
      Top             =   1395
      Width           =   8460
      Begin VB.TextBox txtB1Min 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1710
         TabIndex        =   5
         Top             =   345
         Width           =   870
      End
      Begin VB.TextBox txtB1Max 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3180
         TabIndex        =   6
         Top             =   345
         Width           =   1095
      End
      Begin VB.TextBox txtB1MU 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5310
         TabIndex        =   7
         Top             =   345
         Width           =   870
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Add"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   4755
         TabIndex        =   22
         Top             =   375
         Width           =   390
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "and"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   2700
         TabIndex        =   21
         Top             =   375
         Width           =   390
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Price between"
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   165
         TabIndex        =   20
         Top             =   375
         Width           =   1365
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "currency units to price."
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   6285
         TabIndex        =   19
         Top             =   375
         Width           =   2115
      End
   End
   Begin VB.TextBox txtRound 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3345
      TabIndex        =   1
      Top             =   600
      Width           =   1140
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1065
      MaxLength       =   50
      TabIndex        =   0
      Top             =   120
      Width           =   3030
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   6600
      Picture         =   "frmPT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6315
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Height          =   615
      Left            =   7620
      Picture         =   "frmPT.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6315
      Width           =   1000
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   4245
      TabIndex        =   57
      Top             =   150
      Width           =   435
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Markdown"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   2310
      TabIndex        =   34
      Top             =   1050
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Round up to nearest currency unit when price is within X (cents) of a full unit"
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   180
      TabIndex        =   17
      Top             =   555
      Width           =   3090
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   210
      TabIndex        =   15
      Top             =   180
      Width           =   780
   End
End
Attribute VB_Name = "frmPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oPT As a_PT
Attribute oPT.VB_VarHelpID = -1
Dim flgLoading As Boolean
Public Sub component(pPT As a_PT)
    On Error GoTo errHandler
    Set oPT = pPT
    oPT.BeginEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.component(pPT)", pPT
End Sub

Private Sub chkActive_Click()
    On Error GoTo errHandler
If flgLoading Then Exit Sub
    oPT.ActiveYN = (chkActive = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.chkActive_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkActive_LostFocus()
    On Error GoTo errHandler
   ' chkActive = IIf(oPT.ActiveYN, 1, 0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.chkActive_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkSaleOrReturn_Click()
    On Error GoTo errHandler
If flgLoading Then Exit Sub
    oPT.SaleOrReturn = (chkSaleOrReturn = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.chkSaleOrReturn_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkVoucher_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oPT.SetVoucher = (chkVoucher = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.chkVoucher_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkVoucher_LostFocus()
    On Error GoTo errHandler
   ' chkVoucher = IIf(oPT.IsVoucher, 1, 0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.chkVoucher_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdComm_Click()
    On Error GoTo errHandler
Dim frm As New frmCR
    
    If SecurityControl(enSECURITY_COMM_AUTH, , "Entering commissions", "You do not have permission to open the commission form (or your signature is invalid).") = False Then Exit Sub
        frm.LoadForPT oPT.PTID
        frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdComm_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCopy_Click()
    On Error GoTo errHandler
Dim frm As New frmPTsimple
    frm.Show vbModal
    Me.txtCRSALES = frm.CRSALES
    Me.txtCRSALES_CONTRA = frm.CRSALES_CONTRA
    Me.txtCASHSALES = frm.CASHSALES
    Me.txtCASHSALES_CONTRA = frm.CASHSALES_CONTRA
    Me.txtPURCHASES = frm.PURCHASES
    Me.txtPURCHASES_CONTRA = frm.PURCHASES_CONTRA
    Me.txtVAT = frm.VAT
    oPT.SetCRSALES txtCRSALES
    oPT.SetCRSALES_CONTRA txtCRSALES_CONTRA
    oPT.SetCASHSALES txtCASHSALES
    oPT.SetCASHSALES_CONTRA txtCASHSALES_CONTRA
    oPT.SetPURCHASES txtPURCHASES
    oPT.SetPURCHASES_CONTRA txtPURCHASES_CONTRA
    oPT.SetVAT txtVAT

    Unload frm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdCopy_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oPT_Valid(pMsg As String)
    On Error GoTo errHandler
    txtError = pMsg
    Me.cmdOK.Enabled = (pMsg = "")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.oPT_Valid(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        TOP = 850
        Left = 250
        Width = 8900
        Height = 7500
    End If
    LoadControls
    Me.cmdCopy.Visible = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
flgLoading = True
    txtCode = oPT.code
    Me.txtNumber = oPT.Number
    chkSaleOrReturn = IIf(oPT.SaleOrReturn, 1, 0)
    chkActive = IIf(oPT.ActiveYN, 1, 0)
    chkVoucher = IIf(oPT.IsVoucher, 1, 0)
    txtRound = oPT.RoundedS
    txtB1Min = oPT.B1MinS
    txtB1Max = oPT.B1MaxS
    txtB1MU = oPT.B1MUS
    txtB2Min = oPT.B2MinS
    txtB2Max = oPT.B2MaxS
    txtB2MU = oPT.B2MUS
    txtB3Min = oPT.B3MinS
    txtB3Max = oPT.B3MaxS
    txtB3MU = oPT.B3MUS
    Me.txtCASHSALES = oPT.CASHSALES
    Me.txtCASHSALES_CONTRA = oPT.CASHSALES_CONTRA
    Me.txtCRSALES = oPT.CRSALES
    Me.txtCRSALES_CONTRA = oPT.CRSALES_CONTRA
    Me.txtPURCHASES = oPT.PURCHASES
    Me.txtPURCHASES_CONTRA = oPT.PURCHASES_CONTRA
    Me.txtVAT = oPT.VAT
    Me.txtMarkdown = oPT.DiscountF
flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.LoadControls"
End Sub


Private Sub Text3_Change()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.Text3_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtB1Max_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtB1Max
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB1Max_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtB1Min_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtB1Min
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB1Min_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtB1Min_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPT.SetB1Min txtB1Min
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB1Min_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtB1Max_Validate(Cancel As Boolean)
    On Error GoTo errHandler
If flgLoading Then Exit Sub
    oPT.SetB1Max txtB1Max
     txtB2Min = oPT.B1Max

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB1Max_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtB1MU_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtB1MU
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB1MU_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtB1MU_Validate(Cancel As Boolean)
    On Error GoTo errHandler
If flgLoading Then Exit Sub
    oPT.SetB1MU txtB1MU
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB1MU_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtB2Max_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtB2Max
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB2Max_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtB2Min_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtB2Min
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB2Min_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtB2Min_Validate(Cancel As Boolean)
    On Error GoTo errHandler
If flgLoading Then Exit Sub
    oPT.SetB2Min txtB2Min
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB2Min_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtB2Max_Validate(Cancel As Boolean)
    On Error GoTo errHandler
If flgLoading Then Exit Sub
    oPT.SetB2Max txtB2Max
    txtB3Min = oPT.B2Max

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB2Max_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtB2MU_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtB2MU
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB2MU_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtB2MU_Validate(Cancel As Boolean)
    On Error GoTo errHandler
If flgLoading Then Exit Sub
    oPT.SetB2MU txtB2MU
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB2MU_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtB3Max_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtB3Max
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB3Max_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtB3Min_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtB3Min
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB3Min_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtB3Min_Validate(Cancel As Boolean)
    On Error GoTo errHandler
If flgLoading Then Exit Sub
    oPT.SetB3Min txtB3Min
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB3Min_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtB3Max_Validate(Cancel As Boolean)
    On Error GoTo errHandler
If flgLoading Then Exit Sub
    oPT.SetB3Max txtB3Max
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB3Max_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtB3MU_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtB3MU
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB3MU_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtB3MU_Validate(Cancel As Boolean)
    On Error GoTo errHandler
If flgLoading Then Exit Sub
    oPT.SetB3MU txtB3MU
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtB3MU_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtCode_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtCode = oPT.code
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtCode_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtCode_Change()
            On Error Resume Next
Dim intPos As Integer
If flgLoading Then Exit Sub
    On Error Resume Next
    oPT.SetCode txtCode
    If Err Then
      Beep
      intPos = txtCode.SelStart
      txtCode = oPT.code
      txtCode.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtCode_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtCode_Validate(Cancel As Boolean)
            On Error Resume Next
If flgLoading Then Exit Sub
    oPT.SetCode txtCode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNumber_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtNumber = oPT.Number
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtNumber_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNumber_Change()
            On Error Resume Next
Dim intPos As Integer
If flgLoading Then Exit Sub
    On Error Resume Next
    oPT.SetNumber txtNumber
    If Err Then
      Beep
      intPos = txtNumber.SelStart
      txtNumber = oPT.Number
      txtNumber.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtNumber_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNumber_Validate(Cancel As Boolean)
            On Error Resume Next
If flgLoading Then Exit Sub
    oPT.SetNumber txtNumber
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtNumber_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oPT.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim s As String
    oPT.ApplyEdit
    
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtCRSALES_Validate(Cancel As Boolean)
            On Error Resume Next
If flgLoading Then Exit Sub
    oPT.SetCRSALES txtCRSALES
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtCRSALES_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtCRSALES_CONTRA_Validate(Cancel As Boolean)
            On Error Resume Next
If flgLoading Then Exit Sub
    oPT.SetCRSALES_CONTRA txtCRSALES_CONTRA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtCRSALES_CONTRA_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtCASHSALES_Validate(Cancel As Boolean)
            On Error Resume Next
If flgLoading Then Exit Sub
    oPT.SetCASHSALES txtCASHSALES
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtCASHSALES_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtCASHSALES_CONTRA_Validate(Cancel As Boolean)
            On Error Resume Next
If flgLoading Then Exit Sub
    oPT.SetCASHSALES_CONTRA txtCASHSALES_CONTRA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtCASHSALES_CONTRA_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPURCHASES_Validate(Cancel As Boolean)
            On Error Resume Next
If flgLoading Then Exit Sub
    oPT.SetPURCHASES txtPURCHASES
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtPURCHASES_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPURCHASES_CONTRA_Validate(Cancel As Boolean)
            On Error Resume Next
If flgLoading Then Exit Sub
    oPT.SetPURCHASES_CONTRA txtPURCHASES_CONTRA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtPURCHASES_CONTRA_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtVAT_Validate(Cancel As Boolean)
            On Error Resume Next
If flgLoading Then Exit Sub
    oPT.SetVAT txtVAT
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtVAT_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtMarkdown_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtMarkdown
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtMarkdown_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtMarkdown_LostFocus()
    On Error GoTo errHandler
    txtMarkdown = oPT.DiscountF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtMarkdown_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtMarkdown_Validate(Cancel As Boolean)
            On Error Resume Next
If flgLoading Then Exit Sub
    Cancel = Not oPT.SetDiscount(txtMarkdown)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtMarkdown_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtRound_Validate(Cancel As Boolean)
            On Error Resume Next
If flgLoading Then Exit Sub
    oPT.SetRound txtRound
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.txtRound_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
