VERSION 5.00
Begin VB.Form frmGeneralDiscount 
   BackColor       =   &H00D3D3CB&
   Caption         =   "General discount"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
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
      Height          =   615
      Left            =   1830
      Picture         =   "frmGeneralDiscount.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1020
      Width           =   1000
   End
   Begin VB.TextBox txtGeneralDiscount 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Top             =   450
      Width           =   1395
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
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
      Height          =   285
      Left            =   765
      TabIndex        =   1
      Top             =   465
      Width           =   870
   End
End
Attribute VB_Name = "frmGeneralDiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oInvoice As a_Invoice
Public Sub component(pINV As a_Invoice)
    On Error GoTo errHandler
    Set oInvoice = pINV
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGeneralDiscount.component(pINV)", pINV
End Sub
Private Sub txtDiscount_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtDiscount")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGeneralDiscount.txtDiscount_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGeneralDiscount.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtGeneralDiscount_LostFocus()
    On Error GoTo errHandler
    txtGeneralDiscount = oInvoice.DocDiscountRateF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGeneralDiscount.txtGeneralDiscount_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtGeneralDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
   Cancel = Not oInvoice.SetGeneralDiscount(txtGeneralDiscount)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGeneralDiscount.txtGeneralDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
