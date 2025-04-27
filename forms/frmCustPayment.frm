VERSION 5.00
Begin VB.Form frmCustPayment 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customer account payment"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   3870
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
      Left            =   1380
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Click to find all customers matching the retrictions entered."
      Top             =   3705
      UseMaskColor    =   -1  'True
      Width           =   990
   End
   Begin VB.TextBox txtReason 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000D&
      Height          =   1110
      Left            =   570
      TabIndex        =   2
      ToolTipText     =   "Enter product code,  Acc/ no. or document number or start of supplier name followed by '*'. Hit ENTER to fetch."
      Top             =   2310
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
      Caption         =   "Note"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   585
      TabIndex        =   5
      Top             =   2070
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
Attribute VB_Name = "frmCustPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTPID As Long
Dim lngAmt As Long
Dim dteDate As Date
Dim strReason As String
Dim bAmt As Boolean
Dim bDate As Boolean
Dim bReason As Boolean
Dim strCustomerName As String


Public Sub Component(pTPID As Long, pCustomerName As String)
    lngTPID = pTPID
    strCustomerName = pCustomerName
    Me.txtDate = Format(Date, "dd/mm/yyyy")
End Sub
Private Sub Form_Load()
    bAmt = False
    bDate = False
    bReason = False
    If Me.WindowState <> 2 Then
        Left = 70
        top = 70
        Width = 3990
        Height = 4620
    End If
End Sub

Private Sub txtAmount_Validate(Cancel As Boolean)
Dim res As Long
    bAmt = ConvertToLng(txtAmount, lngAmt)
    If lngAmt < 0 Then bAmt = False
    CheckAction
End Sub


Private Sub txtDate_Validate(Cancel As Boolean)
    bDate = ConvertToDate(txtDate, dteDate)
    CheckAction
End Sub

Private Sub txtReason_Change()
    strReason = FNS(txtReason)
    bReason = Len(strReason) > 4
    CheckAction
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
        lngAmt = lngAmt * -1
    Else
        strDebitCredit = "DEBIT"
    End If
    
    curAmt = CCur(lngAmt) / oPC.Configuration.DefaultCurrency.Divisor
    
    If MsgBox("You are posting a journal " & strDebitCredit & " to " & strCustomerName & " valued " & vbCrLf & vbCrLf & Format(curAmt, "R#,##0.00;(R#,##0.00)"), vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "InsertDebtorsJournal"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput, , lngTPID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@AMT", adInteger, adParamInput, , lngAmt)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@DATE", adDate, adParamInput, , ReverseDate(dteDate))
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@REASON", adVarChar, adParamInput, 255, strReason)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.Execute
    
    Set cmd = Nothing

    Set cmd = New ADODB.Command
    cmd.CommandText = "AgeInvoices"
    cmd.CommandType = adCmdStoredProc

    cmd.Execute
    
    Set cmd = Nothing



'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    MsgBox "Journal posted", vbOKOnly, "Status"
    Unload Me
    Exit Sub
errHandler:
    ErrPreserve
    MsgBox "Journal failed to be posted", vbOKOnly, "Status"
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustJnl.cmdPost_Click"
End Sub

Private Sub CheckAction()
    cmdPost.Enabled = (bAmt And bDate And bReason)
End Sub
