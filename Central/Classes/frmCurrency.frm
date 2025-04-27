VERSION 5.00
Begin VB.Form frmCurrency 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Currency"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1905
      TabIndex        =   20
      Top             =   3090
      Width           =   645
   End
   Begin VB.TextBox txtFactor2 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1905
      TabIndex        =   19
      Top             =   1905
      Width           =   1350
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Standard currencies"
      ForeColor       =   &H8000000D&
      Height          =   2100
      Left            =   2700
      TabIndex        =   14
      Top             =   780
      Visible         =   0   'False
      Width           =   2145
      Begin VB.OptionButton optNone 
         BackColor       =   &H00D3D3CB&
         Caption         =   "None of these"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   315
         TabIndex        =   24
         Top             =   1650
         Width           =   1710
      End
      Begin VB.OptionButton optZAR 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Is ZAR"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   315
         TabIndex        =   18
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton optEUR 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Is EUR"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   315
         TabIndex        =   17
         Top             =   990
         Width           =   1215
      End
      Begin VB.OptionButton optUSD 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Is USD"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   315
         TabIndex        =   16
         Top             =   660
         Width           =   1215
      End
      Begin VB.OptionButton optST 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Is GBP"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   315
         TabIndex        =   15
         Top             =   330
         Width           =   1215
      End
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1890
      TabIndex        =   12
      Top             =   150
      Width           =   2655
   End
   Begin VB.TextBox txtDivisor 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1905
      TabIndex        =   10
      Top             =   2475
      Width           =   1380
   End
   Begin VB.TextBox txtFactor 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1905
      TabIndex        =   8
      Top             =   1545
      Width           =   1350
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1905
      Picture         =   "frmCurrency.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3540
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3030
      Picture         =   "frmCurrency.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3540
      Width           =   1125
   End
   Begin VB.TextBox txtFormatString 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1890
      TabIndex        =   3
      Top             =   1010
      Width           =   2655
   End
   Begin VB.TextBox txtSymbol 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1890
      TabIndex        =   1
      Top             =   580
      Width           =   525
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Use one of the two boxes to enter the conversion factor"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   0
      TabIndex        =   23
      Top             =   1815
      Width           =   1785
   End
   Begin VB.Label lblConvertMessage2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   3300
      TabIndex        =   22
      Top             =   1965
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Standard code"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   165
      TabIndex        =   21
      Top             =   3150
      Width           =   1635
   End
   Begin VB.Label lblConvertMessage1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   3300
      TabIndex        =   13
      Top             =   1545
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Divisor"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   180
      TabIndex        =   11
      Top             =   2505
      Width           =   1635
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Factor"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   1080
      TabIndex        =   9
      Top             =   1590
      Width           =   675
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   1515
      Left            =   390
      TabIndex        =   7
      Top             =   3555
      Width           =   1830
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Format string"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   150
      TabIndex        =   4
      Top             =   1060
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Symbol"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   150
      TabIndex        =   2
      Top             =   635
      Width           =   1635
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Currency name"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   1635
   End
End
Attribute VB_Name = "frmCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oCurr As a_Currency
Dim flgLoading As Boolean
Dim flgLoadingfactor1 As Boolean
Dim flgLoadingfactor2 As Boolean
Dim tlStaff As z_TextListCol

Private Sub EnableOK(pOK As Boolean)
    On Error GoTo errHandler
    Me.cmdOK.Enabled = pOK
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.EnableOK(pOK)", pOK
End Sub
Private Sub oCurr_Valid(pErrors As String, Status As Boolean)
    On Error GoTo errHandler
    EnableOK Status
    lblErrors = pErrors
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.oCurr_Valid(pErrors,Status)", Array(pErrors, Status), EA_NORERAISE
    HandleError
End Sub

Public Sub component(poCurr As a_Currency)
    On Error GoTo errHandler
    Set oCurr = poCurr
   ' oCurr.BeginEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.component(poCurr)", poCurr
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    Me.txtName = oCurr.Description
    Me.txtSymbol = oCurr.Symbol
    Me.txtFormatString = oCurr.FormatString
    Me.txtFactor = oCurr.Factor
    Me.txtDivisor = oCurr.Divisor
    Me.txtCode = oCurr.SYSNAME
    If Not oPC.Configuration.LocalCurrency Is Nothing Then
        Me.lblConvertMessage1.Caption = oCurr.Description & " per " & oPC.Configuration.LocalCurrency.Description
        Me.lblConvertMessage2.Caption = oPC.Configuration.LocalCurrency.Description & " per " & oCurr.Description
    Select Case oCurr.SYSNAME
    Case "ZAR"
        Me.optZAR = 1
    Case "USD"
        Me.optUSD = 1
    Case "EUR"
        Me.optEUR = 1
    Case "GBP"
        Me.optST = 1
    End Select
     End If
   
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.LoadControls"
End Sub
Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oCurr.CancelEdit
    oCurr.BeginEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long

    If optZAR = 1 Then
        oCurr.SYSNAME = "ZAR"
    ElseIf optEUR Then
        oCurr.SYSNAME = "EUR"
    ElseIf optZAR Then
        oCurr.SYSNAME = "ZAR"
    ElseIf optST Then
        oCurr.SYSNAME = "GBP"
    End If
    oCurr.ApplyEdit
    oCurr.BeginEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.Form_Load", , EA_NORERAISE
    HandleError
End Sub


Private Sub optEUR_Click()
    On Error GoTo errHandler
    txtCode = "EUR"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.optEUR_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optST_Click()
    On Error GoTo errHandler
    txtCode = "GBP"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.optST_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optUSD_Click()
    On Error GoTo errHandler
    txtCode = "USD"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.optUSD_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optZAR_Click()
    On Error GoTo errHandler
    txtCode = "ZAR"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.optZAR_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCode_Change()
    On Error GoTo errHandler
    Select Case txtCode
    Case "GBP"
        Me.optST = True
    Case "USD"
        Me.optUSD = True
    Case "EUR"
        Me.optEUR = True
    Case "ZAR"
        Me.optZAR = True
    Case Else
        Me.optNone = True
    End Select
    oCurr.SYSNAME = txtCode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtCode_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtName_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oCurr.Description = txtName
    If Err Then
      Beep
      intPos = txtName.SelStart
      txtName = oCurr.Description
      txtName.SelStart = intPos - 1
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtName_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtName_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtName")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtName_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtName_LostFocus()
    On Error GoTo errHandler
   txtName.Text = oCurr.Description
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtName_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtSymbol_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oCurr.Symbol = txtSymbol
    If Err Then
      Beep
      intPos = txtSymbol.SelStart
      txtSymbol = oCurr.Symbol
      txtSymbol.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtSymbol_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtSymbol_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtSymbol")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtSymbol_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtSymbol_LostFocus()
    On Error GoTo errHandler
   txtSymbol.Text = oCurr.Symbol
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtSymbol_LostFocus", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtFormatString_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oCurr.FormatString = txtFormatString
    If Err Then
      Beep
      intPos = txtFormatString.SelStart
      txtFormatString = oCurr.FormatString
      txtFormatString.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtFormatString_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtFormatString_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtFormatString")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtFormatString_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtFormatString_LostFocus()
    On Error GoTo errHandler
   txtFormatString.Text = oCurr.FormatString
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtFormatString_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtFactor_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoadingfactor1 Then Exit Sub
    If Not IsNumeric(txtFactor) Then Exit Sub
    oCurr.SetFactor txtFactor
    flgLoadingfactor2 = True
    If IsNumeric(txtFactor) Then
        If CDbl(txtFactor) <> 0 Then
            txtFactor2 = Round(CStr(1# / CDbl(txtFactor)), 6)
        End If
    End If
    flgLoadingfactor2 = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtFactor_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFactor2_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoadingfactor2 Then Exit Sub
    If Not IsNumeric(txtFactor2) Then Exit Sub
    flgLoadingfactor1 = True
    oCurr.SetFactor Round((1# / CDbl(txtFactor2)), 6)
    txtFactor = oCurr.Factor
    flgLoadingfactor1 = False

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtFactor2_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtFactor_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtFactor")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtFactor_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtFactor_LostFocus()
    On Error GoTo errHandler
  ' txtFactor.Text = oCurr.factorFormatted
  '  txtFactor2 = Round(CStr(1# / CDbl(txtFactor)), 4)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtFactor_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFactor2_LostFocus()
    On Error GoTo errHandler
  '  txtFactor = Round(CStr(1# / CDbl(txtFactor)), 4)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtFactor2_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDivisor_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    oCurr.SetDivisor txtDivisor
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtDivisor_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDivisor_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtDivisor")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtDivisor_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDivisor_LostFocus()
    On Error GoTo errHandler
   txtDivisor.Text = oCurr.DivisorF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCurrency.txtDivisor_LostFocus", , EA_NORERAISE
    HandleError
End Sub

