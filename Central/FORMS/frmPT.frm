VERSION 5.00
Begin VB.Form frmPT 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Product type"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   8985
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
      Left            =   3675
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5130
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CheckBox chkActive 
      Alignment       =   1  'Right Justify
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
      Height          =   420
      Left            =   6330
      TabIndex        =   2
      Top             =   405
      Width           =   1560
   End
   Begin VB.TextBox txtMarkdown 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   7710
      TabIndex        =   4
      Top             =   1395
      Width           =   870
   End
   Begin VB.CheckBox chkSaleOrReturn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Sale or return"
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
      Height          =   420
      Left            =   6345
      TabIndex        =   3
      Top             =   780
      Width           =   1560
   End
   Begin VB.TextBox txtError 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   165
      MultiLine       =   -1  'True
      TabIndex        =   33
      Top             =   5100
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Band 2"
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
      Height          =   840
      Left            =   165
      TabIndex        =   28
      Top             =   2955
      Width           =   8460
      Begin VB.TextBox txtB2Min 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1710
         TabIndex        =   8
         Top             =   345
         Width           =   870
      End
      Begin VB.TextBox txtB2Max 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   3165
         TabIndex        =   9
         Top             =   345
         Width           =   1095
      End
      Begin VB.TextBox txtB2MU 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   4755
         TabIndex        =   32
         Top             =   405
         Width           =   390
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "and"
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
         Left            =   2700
         TabIndex        =   31
         Top             =   405
         Width           =   390
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Price between"
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
         Left            =   180
         TabIndex        =   30
         Top             =   405
         Width           =   1365
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "currency units to price."
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
         Left            =   6285
         TabIndex        =   29
         Top             =   405
         Width           =   2115
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Band 3"
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
      Height          =   840
      Left            =   165
      TabIndex        =   23
      Top             =   3945
      Width           =   8460
      Begin VB.TextBox txtB3Min 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1710
         TabIndex        =   11
         Top             =   345
         Width           =   870
      End
      Begin VB.TextBox txtB3Max 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   3180
         TabIndex        =   12
         Top             =   345
         Width           =   1095
      End
      Begin VB.TextBox txtB3MU 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   4755
         TabIndex        =   27
         Top             =   405
         Width           =   390
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "and"
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
         Left            =   2700
         TabIndex        =   26
         Top             =   405
         Width           =   390
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Price between"
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
         Left            =   195
         TabIndex        =   25
         Top             =   375
         Width           =   1365
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "currency units to price."
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
         Left            =   6285
         TabIndex        =   24
         Top             =   405
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Band 1"
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
      Height          =   840
      Left            =   165
      TabIndex        =   18
      Top             =   1965
      Width           =   8460
      Begin VB.TextBox txtB1Min 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1710
         TabIndex        =   5
         Top             =   345
         Width           =   870
      End
      Begin VB.TextBox txtB1Max 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   3180
         TabIndex        =   6
         Top             =   345
         Width           =   1095
      End
      Begin VB.TextBox txtB1MU 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   4755
         TabIndex        =   22
         Top             =   405
         Width           =   390
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "and"
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
         Left            =   2700
         TabIndex        =   21
         Top             =   405
         Width           =   390
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Price between"
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
         Left            =   165
         TabIndex        =   20
         Top             =   405
         Width           =   1365
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "currency units to price."
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
         Left            =   6285
         TabIndex        =   19
         Top             =   405
         Width           =   2115
      End
   End
   Begin VB.TextBox txtRound 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   3375
      TabIndex        =   1
      Top             =   840
      Width           =   1140
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   1245
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   6075
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5010
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7365
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5010
      Width           =   1245
   End
   Begin VB.Label Label16 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "(Prices in bands are  in whole currency units)"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   165
      TabIndex        =   35
      Top             =   4815
      Width           =   3810
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Markdown"
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
      Left            =   6405
      TabIndex        =   34
      Top             =   1455
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Round up to nearest currency unit when price is within X (cents) of a full unit"
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
      Height          =   810
      Left            =   210
      TabIndex        =   17
      Top             =   795
      Width           =   3090
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Left            =   630
      TabIndex        =   15
      Top             =   180
      Width           =   1395
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
Public Sub Component(pPT As a_PT)
    Set oPT = pPT
    oPT.BeginEdit
End Sub

Private Sub chkActive_Click()
    oPT.ActiveYN = (chkActive = 1)
End Sub

Private Sub chkActive_LostFocus()
    chkActive = IIf(oPT.ActiveYN, 1, 0)
End Sub

Private Sub chkSaleOrReturn_Click()
    oPT.SaleOrReturn = (chkSaleOrReturn = 1)
End Sub
'
'Private Sub cmdComm_Click()
'Dim frm As New frmCR
'
'    If SecurityControl(enSECURITY_COMM, gSTAFFID, , "Entering commissions", "You do not have permission to open the commission form.") = False Then Exit Sub
'        frm.LoadForPT oPT.PTID
'        frm.Show vbModal
'End Sub

Private Sub oPT_Valid(pMsg As String)
    txtError = pMsg
    Me.cmdOK.Enabled = (pMsg = "")
End Sub
Private Sub Form_Load()
    top = 750
    left = 250
    Width = 9000
    Height = 6000
    LoadControls
End Sub
Private Sub LoadControls()
    txtCode = oPT.code
    chkSaleOrReturn = IIf(oPT.SaleOrReturn, 1, 0)
    chkActive = IIf(oPT.ActiveYN, 1, 0)
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
    Me.txtMarkdown = oPT.DiscountF
End Sub


Private Sub txtB1Max_GotFocus()
    AutoSelect txtB1Max
End Sub

Private Sub txtB1Min_GotFocus()
    AutoSelect txtB1Min
End Sub

Private Sub txtB1Min_Validate(Cancel As Boolean)
    oPT.setB1Min txtB1Min
End Sub
Private Sub txtB1Max_Validate(Cancel As Boolean)
    oPT.setB1Max txtB1Max
     txtB2Min = oPT.B1Max

End Sub

Private Sub txtB1MU_GotFocus()
    AutoSelect txtB1MU
End Sub

Private Sub txtB1MU_Validate(Cancel As Boolean)
    oPT.setB1Mu txtB1MU
End Sub

Private Sub txtB2Max_GotFocus()
    AutoSelect txtB2Max
End Sub

Private Sub txtB2Min_GotFocus()
    AutoSelect txtB2Min
End Sub

Private Sub txtB2Min_Validate(Cancel As Boolean)
    oPT.setB2Min txtB2Min
End Sub
Private Sub txtB2Max_Validate(Cancel As Boolean)
    oPT.setB2Max txtB2Max
    txtB3Min = oPT.B2Max

End Sub

Private Sub txtB2MU_GotFocus()
    AutoSelect txtB2MU
End Sub

Private Sub txtB2MU_Validate(Cancel As Boolean)
    oPT.setB2Mu txtB2MU
End Sub

Private Sub txtB3Max_GotFocus()
    AutoSelect txtB3Max
End Sub

Private Sub txtB3Min_GotFocus()
    AutoSelect txtB3Min
End Sub

Private Sub txtB3Min_Validate(Cancel As Boolean)
    oPT.setB3Min txtB3Min
End Sub
Private Sub txtB3Max_Validate(Cancel As Boolean)
    oPT.setB3Max txtB3Max
End Sub

Private Sub txtB3MU_GotFocus()
    AutoSelect txtB3MU
End Sub

Private Sub txtB3MU_Validate(Cancel As Boolean)
    oPT.setB3Mu txtB3MU
End Sub

Private Sub txtCode_LostFocus()
    If flgLoading Then Exit Sub
    txtCode = oPT.code
End Sub
Private Sub txtCode_Change()
Dim intPos As Integer
    On Error Resume Next
    oPT.SetCode txtCode
    If Err Then
      Beep
      intPos = txtCode.SelStart
      txtCode = oPT.code
      txtCode.SelStart = intPos - 1
    End If
End Sub
Private Sub txtCode_Validate(Cancel As Boolean)
    Cancel = Not oPT.SetCode(txtCode)
End Sub

Private Sub cmdCancel_Click()
    oPT.CancelEdit
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim s As String
    oPT.ApplyEdit
    
    Unload Me
End Sub


Private Sub txtMarkdown_GotFocus()
    AutoSelect txtMarkdown
End Sub

Private Sub txtMarkdown_LostFocus()
    txtMarkdown = oPT.DiscountF
End Sub

Private Sub txtMarkdown_Validate(Cancel As Boolean)
    Cancel = Not oPT.SetDiscount(txtMarkdown)
End Sub

Private Sub txtRound_Validate(Cancel As Boolean)
    oPT.SetRound txtRound
End Sub
