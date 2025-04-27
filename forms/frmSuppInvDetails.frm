VERSION 5.00
Begin VB.Form frmSuppInvDetails 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Supplier invoice details"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5460
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBatchQtyTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1665
      TabIndex        =   4
      Top             =   4290
      Width           =   1965
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6195
      Width           =   135
   End
   Begin VB.TextBox txtBatchTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1635
      TabIndex        =   3
      Top             =   3450
      Width           =   1965
   End
   Begin VB.TextBox txtBatchTotalExtras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1665
      TabIndex        =   2
      Top             =   2580
      Width           =   1965
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Default         =   -1  'True
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
      Height          =   450
      Left            =   4575
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5085
      Width           =   795
   End
   Begin VB.TextBox txtInvDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1665
      TabIndex        =   1
      Top             =   1530
      Width           =   1965
   End
   Begin VB.TextBox txtInvRef 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1650
      TabIndex        =   0
      Top             =   645
      Width           =   1965
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantity items"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1215
      TabIndex        =   12
      Top             =   4020
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "(Click ESC to cancel)"
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
      Height          =   240
      Left            =   15
      TabIndex        =   11
      Top             =   5460
      Width           =   1800
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total invoice value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1200
      TabIndex        =   9
      Top             =   3180
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Additional charges (e.g. freight and insurance)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   -8
      TabIndex        =   8
      Top             =   2235
      Width           =   5310
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1665
      TabIndex        =   7
      Top             =   1260
      Width           =   1965
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice ref."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1747
      TabIndex        =   6
      Top             =   375
      Width           =   1800
   End
End
Attribute VB_Name = "frmSuppInvDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strRef As String
Dim dteInvDate As Date
Dim flgLoading As Boolean
Dim oDel As a_Delivery
Dim bCancel As Boolean
Public Property Get Cancelled() As Boolean
    Cancelled = bCancel
End Property

Public Sub Component(pDel As a_Delivery)
    Set oDel = pDel
End Sub

Private Sub cmdCancel_Click()
    bCancel = True
    Me.Hide
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    flgLoading = True
    txtInvDate = oDel.SupplierInvoiceDateF
    txtInvRef = oDel.SupplierInvoiceRef
    txtBatchTotal = oDel.BatchTotal
    txtBatchQtyTotal = oDel.BatchQtyTotalF
    txtBatchTotalExtras = oDel.BatchTotalExtras
    CheckCloseButton
    bCancel = False
    flgLoading = False
End Sub

Private Sub txtBatchTotalExtras_LostFocus()
    txtBatchTotalExtras = oDel.BatchTotalExtras
End Sub
Private Sub txtBatchTotalExtras_Validate(Cancel As Boolean)
    If flgLoading Then Exit Sub
    Cancel = Not oDel.SetBatchTotalExtras(txtBatchTotalExtras)
    CheckCloseButton
End Sub
Private Sub txtBatchTotalExtras_GotFocus()
    AutoSelect txtBatchTotalExtras
End Sub

Private Sub txtBatchTotal_LostFocus()
    txtBatchTotal = oDel.BatchTotal
End Sub
Private Sub txtBatchTotal_Validate(Cancel As Boolean)
    If flgLoading Then Exit Sub
    Cancel = Not oDel.SetBatchTotal(txtBatchTotal)
    CheckCloseButton
End Sub
Private Sub txtBatchTotal_GotFocus()
    AutoSelect txtBatchTotal
End Sub

Private Sub txtInvDate_Validate(Cancel As Boolean)
Dim lngMonths As Long

    If flgLoading Then Exit Sub
    If Not IsDate(txtInvDate) Then
        Cancel = True
        Exit Sub
    Else
        lngMonths = DateDiff("m", Date, CDate(txtInvDate))
        If lngMonths > 0 Then
            Cancel = True
            Exit Sub
        ElseIf lngMonths < -2 Then
            Select Case lngMonths
            Case Is > -13
                If MsgBox("The invoice date is more than two months ago. Is it correct?", vbQuestion + vbYesNo, "Warning") = vbNo Then
                    Cancel = True
                    Exit Sub
                End If
            Case Else
                MsgBox "The invoice date is too old. You must correct it.", vbExclamation, "Warning"
                Cancel = True
                Exit Sub
            End Select
        End If
    End If
    oDel.SupplierInvoiceDate = CDate(txtInvDate)
    CheckCloseButton
End Sub
Private Sub txtInvDate_GotFocus()
    AutoSelect txtInvDate
End Sub

Private Sub txtInvRef_GotFocus()
    AutoSelect txtInvRef
End Sub

Private Sub txtInvRef_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Set oDel = Nothing
    bCancel = True
    Me.Hide
End If
End Sub

Private Sub txtInvRef_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Set oDel = Nothing
    bCancel = True
    Me.Hide
End If
End Sub

Private Sub txtInvRef_Validate(Cancel As Boolean)
    If flgLoading Then Exit Sub
    If Not Len(txtInvRef) > 2 Then
        Cancel = True
    Else
        Cancel = False
        oDel.SupplierInvoiceRef = txtInvRef
    End If
    CheckCloseButton
End Sub

Private Sub CheckCloseButton()
    cmdClose.Enabled = (IsDate(txtInvDate) And Len(txtInvRef) > 2)
End Sub
Public Property Get Ref() As String
    Ref = strRef
End Property
Public Property Get InvDate() As Date
    InvDate = dteInvDate
End Property

Private Sub txtBatchQtyTotal_GotFocus()
    AutoSelect txtBatchQtyTotal
End Sub

Private Sub txtBatchQtyTotal_Validate(Cancel As Boolean)
    If flgLoading Then Exit Sub
    Cancel = Not oDel.SetBatchQtyTotal(txtBatchQtyTotal)
    CheckCloseButton

End Sub
