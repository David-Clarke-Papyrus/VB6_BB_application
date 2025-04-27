VERSION 5.00
Begin VB.Form frmHeader_GRN 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Supplier invoice details"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   3930
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBatchQtyTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   990
      TabIndex        =   4
      Top             =   3255
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
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   2580
      Width           =   1965
   End
   Begin VB.TextBox txtBatchTotalExtras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   990
      TabIndex        =   2
      Top             =   1905
      Width           =   1965
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
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
      Left            =   2775
      Picture         =   "frmHeader_GRN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3765
      Width           =   1000
   End
   Begin VB.TextBox txtInvDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   990
      TabIndex        =   1
      Top             =   1230
      Width           =   1965
   End
   Begin VB.TextBox txtInvRef 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   735
      MaxLength       =   50
      TabIndex        =   0
      Top             =   600
      Width           =   2445
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantity items"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   540
      TabIndex        =   12
      Top             =   3030
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "(Click ESC to cancel)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   30
      TabIndex        =   11
      Top             =   4185
      Width           =   1800
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total invoice value"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   525
      TabIndex        =   9
      Top             =   2355
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Additional charges (e.g. freight and insurance)"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   -690
      TabIndex        =   8
      Top             =   1650
      Width           =   5310
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice date"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   990
      TabIndex        =   7
      Top             =   1005
      Width           =   1965
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice ref."
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1065
      TabIndex        =   6
      Top             =   375
      Width           =   1800
   End
End
Attribute VB_Name = "frmHeader_GRN"
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

Public Sub component(pDel As a_Delivery)
    On Error GoTo errHandler
    Set oDel = pDel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.component(pDel)", pDel
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancel = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    txtInvDate = oDel.SupplierInvoiceDateF
    txtInvRef = oDel.SupplierInvoiceRef
    txtBatchTotal = oDel.BatchTotal
    txtBatchQtyTotal = oDel.BatchQtyTotalF
    txtBatchTotalExtras = oDel.BatchTotalExtras
    CheckCloseButton
    bCancel = False
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtBatchTotalExtras_LostFocus()
    On Error GoTo errHandler
    txtBatchTotalExtras = oDel.BatchTotalExtras
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.txtBatchTotalExtras_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtBatchTotalExtras_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    Cancel = Not oDel.SetBatchTotalExtras(txtBatchTotalExtras)
    CheckCloseButton
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.txtBatchTotalExtras_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtBatchTotalExtras_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtBatchTotalExtras
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.txtBatchTotalExtras_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtBatchTotal_LostFocus()
    On Error GoTo errHandler
    txtBatchTotal = oDel.BatchTotal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.txtBatchTotal_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtBatchTotal_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    Cancel = Not oDel.SetBatchTotal(txtBatchTotal)
    CheckCloseButton
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.txtBatchTotal_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtBatchTotal_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtBatchTotal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.txtBatchTotal_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtInvDate_Validate(Cancel As Boolean)
    On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.txtInvDate_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtInvDate_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtInvDate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.txtInvDate_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtInvRef_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtInvRef
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.txtInvRef_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtInvRef_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
If KeyCode = 27 Then
    Set oDel = Nothing
    bCancel = True
    Me.Hide
End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.txtInvRef_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub txtInvRef_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
If KeyAscii = 27 Then
    Set oDel = Nothing
    bCancel = True
    Me.Hide
End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.txtInvRef_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub txtInvRef_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim oSQL As New z_SQL
Dim strResult As String

    If flgLoading Then Exit Sub
    If Not Len(txtInvRef) > 2 Then
        Cancel = True
    Else
        Cancel = False
        oDel.SupplierInvoiceRef = txtInvRef
    End If
    If oPC.GetProperty("CheckRefsOnGRN") = "TRUE" Then
        oSQL.FindSuppInvMatch oDel.SupplierInvoiceRef, strResult, oDel.TRID
        If strResult > "" Then
            If MsgBox("This reference number is found on document(s): " & strResult & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
                Cancel = True
            End If
        End If
    End If
    CheckCloseButton
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.txtInvRef_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub CheckCloseButton()
    On Error GoTo errHandler
    cmdClose.Enabled = (IsDate(txtInvDate) And Len(txtInvRef) > 2)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.CheckCloseButton"
End Sub
Public Property Get Ref() As String
    Ref = strRef
End Property
Public Property Get InvDate() As Date
    InvDate = dteInvDate
End Property

Private Sub txtBatchQtyTotal_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtBatchQtyTotal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.txtBatchQtyTotal_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtBatchQtyTotal_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    Cancel = Not oDel.SetBatchQtyTotal(txtBatchQtyTotal)
    CheckCloseButton

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_GRN.txtBatchQtyTotal_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
