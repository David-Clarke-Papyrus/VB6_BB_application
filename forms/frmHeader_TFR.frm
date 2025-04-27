VERSION 5.00
Begin VB.Form frmHeader_TFR 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Transfer details"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   3735
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBatchQtyTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   915
      TabIndex        =   3
      Top             =   2475
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
      TabIndex        =   8
      Top             =   6195
      Width           =   135
   End
   Begin VB.TextBox txtBatchTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   915
      TabIndex        =   2
      Top             =   1800
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
      Left            =   2640
      Picture         =   "frmHeader_TFR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2835
      Width           =   1000
   End
   Begin VB.TextBox txtDocDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   915
      TabIndex        =   1
      Top             =   1080
      Width           =   1965
   End
   Begin VB.TextBox txtDocRef 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   900
      TabIndex        =   0
      Top             =   420
      Width           =   1965
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Quantity items"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   465
      TabIndex        =   10
      Top             =   2250
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
      Left            =   15
      TabIndex        =   9
      Top             =   3315
      Width           =   1800
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total transfer value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   465
      TabIndex        =   7
      Top             =   1575
      Width           =   2895
   End
   Begin VB.Label lblDocDate 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sender's document date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   600
      TabIndex        =   6
      Top             =   855
      Width           =   2625
   End
   Begin VB.Label lblDocRef 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sender's document code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   630
      TabIndex        =   5
      Top             =   195
      Width           =   2550
   End
End
Attribute VB_Name = "frmHeader_TFR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strRef As String
Dim dteInvDate As Date
Dim flgLoading As Boolean
Dim oTFR As a_TF
Dim bCancel As Boolean
Dim mInOut As String
Public Property Get Cancelled() As Boolean
    On Error GoTo errHandler
    Cancelled = bCancel
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.Cancelled"
End Property

Public Sub component(pTFR As a_TF)
    On Error GoTo errHandler
    Set oTFR = pTFR
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.component(pTFR)", pTFR
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancel = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    txtDocDate = oTFR.SendersDocDateF
    txtDocRef = oTFR.SendersDocRef
    txtBatchTotal = oTFR.BatchTotal
    txtBatchQtyTotal = oTFR.BatchQtyTotal
    CheckCloseButton
    bCancel = False
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.Form_Load", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtBatchTotal_LostFocus()
    On Error GoTo errHandler
    txtBatchTotal = oTFR.BatchTotal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.txtBatchTotal_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtBatchTotal_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    Cancel = Not oTFR.SetBatchTotal(txtBatchTotal)
    CheckCloseButton
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.txtBatchTotal_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtBatchTotal_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtBatchTotal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.txtBatchTotal_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDocDate_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim lngMonths As Long

    If flgLoading Then Exit Sub
    If Not IsDate(txtDocDate) Then
        Cancel = True
        Exit Sub
    Else
        lngMonths = DateDiff("m", Date, CDate(txtDocDate))
        If lngMonths > 0 Then
            Cancel = True
            Exit Sub
        ElseIf lngMonths < -2 Then
            Select Case lngMonths
            Case Is > -13
                If MsgBox("The transfer date is more than two months ago. Is it correct?", vbQuestion + vbYesNo, "Warning") = vbNo Then
                    Cancel = True
                    Exit Sub
                End If
            Case Else
                MsgBox "The transfer date is too old. You must correct it.", vbExclamation, "Warning"
                Cancel = True
                Exit Sub
            End Select
        End If
    End If
    oTFR.SendersDocDate = CDate(txtDocDate)
    CheckCloseButton
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.txtDocDate_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtDocDate_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtDocDate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.txtDocDate_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDocRef_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtDocRef
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.txtDocRef_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDocRef_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
If KeyCode = 27 Then
    Set oTFR = Nothing
    bCancel = True
    Me.Hide
End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.txtDocRef_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub txtDocRef_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
If KeyAscii = 27 Then
    Set oTFR = Nothing
    bCancel = True
    Me.Hide
End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.txtDocRef_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub txtDocRef_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not Len(txtDocRef) > 2 Then
        Cancel = True
    Else
        Cancel = False
        oTFR.SendersDocRef = txtDocRef
    End If
    CheckCloseButton
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.txtDocRef_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub CheckCloseButton()
    On Error GoTo errHandler
    cmdclose.Enabled = (IsDate(txtDocDate) And Len(txtDocRef) > 2)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.CheckCloseButton"
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
    ErrorIn "frmHeader_TFR.txtBatchQtyTotal_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtBatchQtyTotal_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    Cancel = Not oTFR.SetBatchQtyTotal(txtBatchQtyTotal)
    CheckCloseButton

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_TFR.txtBatchQtyTotal_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Public Property Get InOut() As String
    InOut = mInOut
End Property

