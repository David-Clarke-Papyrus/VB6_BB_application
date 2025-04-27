VERSION 5.00
Begin VB.Form frmAudit_ProductPrices 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Price change log"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6540
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstreasons 
      Height          =   1815
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   5655
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
      Height          =   615
      Left            =   330
      Picture         =   "frmAudit_ProductPrices.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1000
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
      Height          =   615
      Left            =   4980
      Picture         =   "frmAudit_ProductPrices.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3750
      Width           =   1000
   End
   Begin VB.TextBox txtReason 
      Alignment       =   2  'Center
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
      ForeColor       =   &H8000000D&
      Height          =   780
      Left            =   330
      TabIndex        =   2
      Top             =   2760
      Width           =   5700
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Reason for price/discount/terms alteration (min 12 charsacters)"
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
      Height          =   405
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   5610
   End
End
Attribute VB_Name = "frmAudit_ProductPrices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancel As Boolean


Private Sub Form_Load()
    lstreasons.AddItem "1. New product creation with price"
    lstreasons.AddItem "2. Marketing dept Was/Now price change"
    lstreasons.AddItem "3. Supplier price change"
    lstreasons.AddItem "4. Promotion price change"
    lstreasons.AddItem "5. Match warehouse price"
    lstreasons.AddItem "6. Marked Down to Paperback Price"
    lstreasons.AddItem "7. Other"
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancel = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAudit_ProductPrices.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler

    bCancel = False
    If Len(Me.txtReason) > 11 Then
        Me.Hide
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAudit_ProductPrices.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub lstreasons_Click()
    If lstreasons.text <> "7. Other" Then
        txtReason.text = lstreasons.text
    Else
        txtReason.text = ""
    End If
End Sub

Private Sub txtReason_Change()
    On Error GoTo errHandler
    Me.cmdOK.Enabled = (Len(StripToAlphanumeric(txtReason)) > 10)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAudit_ProductPrices.txtReason_Change", , EA_NORERAISE
    HandleError
End Sub

Public Property Get Reason() As String
    On Error GoTo errHandler
    Reason = stripCRLF(txtReason)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAudit_ProductPrices.Reason"
End Property

Public Property Get Cancelled() As Boolean
    On Error GoTo errHandler
    Cancelled = bCancel
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAudit_ProductPrices.Cancelled"
End Property
