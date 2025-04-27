VERSION 5.00
Begin VB.Form frmCustJnl 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customer account journal"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3870
   Icon            =   "frmCustJnl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   3870
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Type"
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   165
      TabIndex        =   7
      Top             =   1785
      Width           =   3405
      Begin VB.PictureBox Picture 
         BackColor       =   &H00D3D3CB&
         Height          =   435
         Left            =   180
         ScaleHeight     =   375
         ScaleWidth      =   3030
         TabIndex        =   8
         Top             =   210
         Width           =   3090
         Begin VB.OptionButton optDebit 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Debit"
            ForeColor       =   &H8000000D&
            Height          =   300
            Left            =   630
            TabIndex        =   10
            Top             =   60
            Width           =   750
         End
         Begin VB.OptionButton optCredit 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Credit"
            ForeColor       =   &H8000000D&
            Height          =   300
            Left            =   1590
            TabIndex        =   9
            Top             =   60
            Value           =   -1  'True
            Width           =   750
         End
      End
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   345
      Left            =   900
      TabIndex        =   1
      ToolTipText     =   "Enter product code,  Acc/ no. or document number or start of supplier name followed by '*'. Hit ENTER to fetch."
      Top             =   1350
      Width           =   1875
   End
   Begin VB.CommandButton cmdPost 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Post"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1305
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Click to find all customers matching the retrictions entered."
      Top             =   4035
      UseMaskColor    =   -1  'True
      Width           =   990
   End
   Begin VB.TextBox txtReason 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000D&
      Height          =   1110
      Left            =   587
      TabIndex        =   2
      ToolTipText     =   "Enter product code,  Acc/ no. or document number or start of supplier name followed by '*'. Hit ENTER to fetch."
      Top             =   2715
      Width           =   2500
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   345
      Left            =   900
      TabIndex        =   0
      ToolTipText     =   "Enter product code,  Acc/ no. or document number or start of supplier name followed by '*'. Hit ENTER to fetch."
      Top             =   510
      Width           =   1875
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Posting date"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   960
      TabIndex        =   6
      Top             =   1080
      Width           =   1980
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   600
      TabIndex        =   5
      Top             =   2475
      Width           =   1980
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount (V.A.T. inclusive)"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   870
      TabIndex        =   4
      Top             =   210
      Width           =   1980
   End
End
Attribute VB_Name = "frmCustJnl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTPID As Long
Dim dblAmount As Double
Dim dteDate As Date
Dim strReason As String
Dim bAmt As Boolean
Dim bDate As Boolean
Dim bReason As Boolean
Dim strCustomerName As String


Public Sub component(pTPID As Long, pCustomerName As String)
    On Error GoTo errHandler
    lngTPID = pTPID
    strCustomerName = pCustomerName
    Me.txtDate = Format(Date, "dd/mm/yyyy")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustJnl.component(pTPID,pCustomerName)", Array(pTPID, pCustomerName)
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
    bAmt = False
    bDate = False
    bReason = False
    If Me.WindowState <> 2 Then
        Width = 3990
        Height = 5220
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustJnl.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtAmount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim Res As Long
    bAmt = ConvertToDBL(txtAmount, dblAmount)
    If dblAmount < 0 Then bAmt = False
    CheckAction
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustJnl.txtAmount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtDate_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    bDate = ConvertToDate(txtDate, dteDate)
    CheckAction
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustJnl.txtDate_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtReason_Change()
    On Error GoTo errHandler
    strReason = FNS(txtReason)
    bReason = Len(strReason) > 4
    CheckAction
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustJnl.txtReason_Change", , EA_NORERAISE
    HandleError
End Sub

'Private Sub txtReason_Validate(Cancel As Boolean)
'    strReason = FNS(txtReason)
'    bReason = Len(strReason) > 4
'    CheckAction
'End Sub

Private Sub cmdPost_Click()
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
Dim strDebitCredit As String
Dim curAmt As Currency

    If optCredit = True Then
        strDebitCredit = "CREDIT"
        dblAmount = dblAmount * -1
    Else
        strDebitCredit = "DEBIT"
    End If
    
    curAmt = CCur(dblAmount) / oPC.Configuration.DefaultCurrency.Divisor
    If MsgBox("You are posting a journal " & strDebitCredit & " to " & strCustomerName & " valued " & vbCrLf & vbCrLf & Format(curAmt, "R#,##0.00;(R#,##0.00)"), vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "InsertDebtorsJournal_2"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput, , lngTPID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@AMT", adNumeric, adParamInput, , curAmt)
    par.NumericScale = 2
    par.Precision = 20
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@DATE", adDate, adParamInput, , ReverseDate(dteDate))
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@REASON", adVarChar, adParamInput, 255, strReason)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    
    Set cmd = Nothing

    Set cmd = New ADODB.Command
    cmd.CommandText = "AgeInvoices"
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.COShort
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput, , lngTPID)
    cmd.Parameters.Append par
    Set par = Nothing

    cmd.execute
    
    Set cmd = Nothing



'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    MsgBox "Journal posted", vbOKOnly, "Status"
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustJnl.cmdPost_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub CheckAction()
    On Error GoTo errHandler
    cmdPost.Enabled = (bAmt And bDate And bReason)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustJnl.CheckAction"
End Sub
