VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmCustPmt 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customer account payment"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   8820
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPostBatch 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Post"
      Height          =   465
      Left            =   7515
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Click to find all customers matching the retrictions entered."
      Top             =   3960
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox t2 
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
      Left            =   6540
      TabIndex        =   11
      ToolTipText     =   "Enter product code,  Acc/ no. or document number or start of supplier name followed by '*'. Hit ENTER to fetch."
      Top             =   1410
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox t1 
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
      Left            =   5445
      TabIndex        =   10
      ToolTipText     =   "Enter product code,  Acc/ no. or document number or start of supplier name followed by '*'. Hit ENTER to fetch."
      Top             =   1410
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Post"
      Height          =   465
      Left            =   7515
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Click to find all customers matching the retrictions entered."
      Top             =   1320
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2745
      Picture         =   "frmCustPmt2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   975
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   465
      Left            =   330
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Click to find all customers matching the retrictions entered."
      Top             =   1110
      UseMaskColor    =   -1  'True
      Width           =   990
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
      Height          =   315
      Left            =   2475
      TabIndex        =   1
      ToolTipText     =   "Enter product code,  Acc/ no. or document number or start of supplier name followed by '*'. Hit ENTER to fetch."
      Top             =   420
      Width           =   1875
   End
   Begin VB.CommandButton cmdPost 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Post"
      Enabled         =   0   'False
      Height          =   465
      Left            =   1320
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Click to find all customers matching the retrictions entered."
      Top             =   1110
      UseMaskColor    =   -1  'True
      Width           =   990
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
      Height          =   315
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Enter product code,  Acc/ no. or document number or start of supplier name followed by '*'. Hit ENTER to fetch."
      Top             =   435
      Width           =   1875
   End
   Begin TrueOleDBGrid60.TDBGrid gDeposits 
      Height          =   2025
      Left            =   225
      OleObjectBlob   =   "frmCustPmt2.frx":038A
      TabIndex        =   7
      Top             =   1905
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.Label lblDetails 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   705
      Left            =   3420
      TabIndex        =   5
      Top             =   135
      Width           =   3345
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Posting date"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   2535
      TabIndex        =   4
      Top             =   150
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
      Left            =   315
      TabIndex        =   3
      Top             =   135
      Width           =   1980
   End
End
Attribute VB_Name = "frmCustPmt"
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
Dim strInvoices As String
Dim lngInvoiceID As Long
Dim XA As New XArrayDB

Public Sub component(pTPID As Long, pCustomerName As String, Optional x As XArrayDB)
    On Error GoTo errHandler
Dim i As Integer
    lngTPID = pTPID
    strCustomerName = pCustomerName
    Me.txtDate = Format(Date, "dd/mm/yyyy")
    strInvoices = ""
    If Not x Is Nothing Then
        For i = 0 To x.UpperBound(1)
            strInvoices = strInvoices & x(i, 1)
        Next
        Me.lblDetails.Caption = strCustomerName & vbCrLf & strInvoices
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.component(pTPID,pCustomerName,x)", Array(pTPID, pCustomerName, x)
End Sub
Public Sub Component2(pTPID As Long, pCustomerName As String, pINVOICEID As Long)
    On Error GoTo errHandler
Dim i As Integer
Dim oInvoice As New a_Invoice
'Dim oMatch As New a_PaymentMatches
    lngInvoiceID = pINVOICEID
    oInvoice.Load lngInvoiceID, True
    lngTPID = pTPID
    strCustomerName = pCustomerName
'    oMatch.PostDebtorsPaymentOpenItem pInvoiceID,
    Me.txtDate = Format(Date, "dd/mm/yyyy")
    Me.txtAmount = oInvoice.TotalPayable(False)
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.Component2(pTPID,pCustomerName,pINVOICEID)", Array(pTPID, pCustomerName, _
         pINVOICEID)
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub


Private Sub Command1_Click()
MsgBox XA(CInt(t1), CInt(t2))
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    bAmt = False
    bDate = False
    If Me.WindowState <> 2 Then
        Left = 70
        TOP = 70
        Width = 3990
        Height = 4620
    End If
' Allocate space for 300 rows, 4 columns
    XA.ReDim 0, 299, 0, 3

    Dim row As Long, col As Integer


' Bind True DBGrid Control to this XArrayDB instance
    Set gDeposits.Array = XA
    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.Form_Load", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtAmount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim Res As Long
    bAmt = ConvertToLng(txtAmount, lngAmt)
    If lngAmt < 0 Then bAmt = False
    CheckAction
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.txtAmount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtDate_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    bDate = ConvertToDate(txtDate, dteDate)
    CheckAction
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.txtDate_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub cmdPost_Click()
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
Dim strDebitCredit As String
Dim curAmt As Currency


    curAmt = CCur(lngAmt) / oPC.Configuration.DefaultCurrency.Divisor
    
    If MsgBox("You are posting a Payment to " & strCustomerName & " valued " & vbCrLf & vbCrLf & Format(curAmt, "R#,##0.00;(R#,##0.00)"), vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "[CreatePaymentToAccount_NoExchange]"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput, , lngTPID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@DATE", adDate, adParamInput, , dteDate)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@TRSTATUS", adInteger, adParamInput, , 4)  'Create the Payment Transaction
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@Amount", adDouble, adParamInput, , curAmt)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@InvoiceID", adDouble, adParamInput, , lngInvoiceID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort  'THis is used only by PBKS_POSSVR so it must use this connection - DO NOT CHANGE
    
    cmd.execute
    
    Set cmd = Nothing



'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    MsgBox "Payment posted", vbOKOnly, "Status"
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.cmdPost_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdPostBatch_Click()
Dim xMLDoc As ujXML
Dim XMLArgs As String
Dim Strguid As String
Dim i As Integer
Dim oSM As New z_StockManager
Dim lngPaid As Long

'    If oPC.Configuration.Signtransactions = True Then
'        If SecurityControl(enSECURITY_ACCEPTACPAYMENT, , "Save payment batch", DOCAPPROVAL) = False Then
'               Exit Sub
'        End If
'    End If
    
    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "doc_PaymentBatch"
            .chCreate "MessageType"
                .elText = "PAYMENT_ACTION"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "TPID"
                .elText = CStr(lngTPID)
            .elCreateSibling "StaffID"
                .elText = CStr(gSTAFFID)
            .elCreateSibling "DetailLines", True
            i = 0
            Do While XA.Value(i, 0) > ""
                .chCreate "I"
                .chCreate "Reference"
                    .elText = XA.Value(i, 0)
                .elCreateSibling "Date", True
                    If IsDate(FND(XA.Value(i, 1))) Then
                        .elText = Format(FND(XA.Value(i, 1)), "YYYYMMDD")
                    End If
                .elCreateSibling "Amount", True
                    .elText = FNDBL(XA.Value(i, 2))
                .elCreateSibling "SettlementDiscount", True
                    .elText = FNDBL(XA.Value(i, 3))
                .navUP
                .navUP
                i = i + 1
            Loop

         XMLArgs = .docXML
  
    End With
    oSM.InsertScript Strguid, XMLArgs

    If Strguid > "" Then
        oSM.Action_InsertPayments Strguid, lngPaid
    End If

End Sub

Private Sub CheckAction()
    On Error GoTo errHandler
    cmdPost.Enabled = (bAmt And bDate)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.CheckAction"
End Sub
