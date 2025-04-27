VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmCreditAccountPOS 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Invoices outstanding for customer."
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   9975
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtInput 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   570
      Left            =   75
      TabIndex        =   0
      Top             =   5655
      Width           =   3390
   End
   Begin TrueOleDBGrid60.TDBGrid gInvLines 
      Height          =   2490
      Left            =   120
      OleObjectBlob   =   "frmCreditAccountPOS.frx":0000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2655
      Width           =   9720
   End
   Begin TrueOleDBGrid60.TDBGrid gInvoices 
      Height          =   1830
      Left            =   135
      OleObjectBlob   =   "frmCreditAccountPOS.frx":6282
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   345
      Width           =   5430
   End
   Begin VB.Label SB 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
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
      Height          =   435
      Left            =   60
      TabIndex        =   9
      Top             =   6330
      Width           =   9795
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCreditAccountPOS.frx":A024
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
      Height          =   1245
      Left            =   5760
      TabIndex        =   8
      Top             =   360
      Width           =   3525
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice lines"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   315
      Left            =   210
      TabIndex        =   7
      Top             =   2340
      Width           =   1950
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   30
      Width           =   1950
   End
   Begin VB.Label lblInput 
      BackStyle       =   0  'Transparent
      Caption         =   "Scan returns"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   405
      Left            =   90
      TabIndex        =   5
      Top             =   5220
      Width           =   5400
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   5715
      TabIndex        =   4
      Top             =   5250
      Width           =   2115
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   420
      Left            =   8025
      TabIndex        =   3
      Top             =   5235
      Width           =   1620
   End
End
Attribute VB_Name = "frmCreditAccountPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xInvoices As XArrayDB
Dim xInvLines As XArrayDB
Dim bCancelled As Boolean

Dim lngCustomerID As Long
Dim lngInvoiceID As Long
Dim strMode As String
Dim cInvoices As c_InvoicesContaining
Dim cInvoicelines As c_InvoiceLinesPOS

Public Sub component(pTPID As Long, pCustomerName As String)
    lngCustomerID = pTPID
    Caption = "Invoices outstanding for customer: " & pCustomerName
'    If Not xInvLines Is Nothing Then
'        Set xInvLines = Nothing
'    End If
End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property
Private Sub cmd_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    txtInput.Visible = True
    txtInput.SetFocus

End Sub

Private Sub Form_Load()
    Set xInvoices = New XArrayDB
    Set xInvLines = New XArrayDB
    
    Set cInvoices = New c_InvoicesContaining
    Set cInvoicelines = New c_InvoiceLinesPOS
    Me.SB.Caption = "Scan product,(C)Correct,(F)Finish"
    strMode = "R"
  '  cInvoicelines.Load lngCustomerID
   ' LoadInvoicelinesArray ""
    lblInput.Caption = "Scan returned items"
    lblInput.ForeColor = &H714942
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cInvoicelines = Nothing
    Set cInvoices = Nothing
    Set xInvoices = Nothing
    Set xInvLines = Nothing
End Sub

Private Sub refreshInvoicelines(PID As String)
    Set cInvoicelines = Nothing
    Set cInvoicelines = New c_InvoiceLinesPOS
    cInvoicelines.Load lngInvoiceID
    LoadInvoicelinesArray PID
End Sub





'
Private Sub gInvoices_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If IsNull(gInvoices.Bookmark) Then Exit Sub
    lngInvoiceID = FNN(xInvoices(gInvoices.Bookmark, 5))
    refreshInvoicelines FNS(xInvoices(gInvoices.Bookmark, 6))
End Sub



Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = 13 Then
        ActionInput txtInput
        txtInput = ""
     End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSAPPRet.txtInput_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub ActionInput(pIn As String)
    pIn = UCase(FNS(pIn))
    If pIn = "F" Then
        bCancelled = False
        Me.Hide
        Exit Sub
    ElseIf pIn = "C" Then
        bCancelled = True
        Me.Hide
        Exit Sub
    End If
'    If xInvLines.Count(1) < 1 Then
'        MsgBox "There are no lines.", vbInformation + vbOKOnly, "Can't do this"
'        Exit Sub
'    End If
    If LoadInvoiceHeaders(pIn) > 0 Then
        txtInput = ""
    ElseIf xInvLines.UpperBound(1) > 0 Then
        FindRowAmongInvLines (pIn)
    Else
        MsgBox "There is no product with this code invoiced to this customer.", vbInformation + vbOKOnly, "Can't do this"
    End If
End Sub
Private Function FindRowAmongInvLines(pIn As String) As String
Dim i As Integer
Dim iResult As Integer
Dim strPID As String
Dim oSM As New z_SM
    strPID = oSM.FindPIDFromCode(pIn)
    For i = 1 To xInvLines.Count(1)
        If xInvLines(i, 12) = strPID Then
            iResult = i
            Exit For
        End If
    Next i
    If iResult > 0 Then
        If xInvLines.Value(iResult, 4) - xInvLines.Value(iResult, 5) - xInvLines.Value(iResult, 6) > 0 Then
            xInvLines.Value(iResult, 6) = xInvLines.Value(iResult, 6) + 1
        Else
            MsgBox "This row is already fully credited", vbOKOnly, "Can't do this"
        End If
    Else
        MsgBox "Row not found", vbOKOnly, "Can't find row"
    End If
    Me.gInvLines.Refresh
End Function

Private Function LoadInvoiceHeaders(pIn As String) As Long
Dim i As Integer
Dim iRow As Integer
Dim lngToBeInvoiced As Long

    Set cInvoices = New c_InvoicesContaining
    cInvoices.Load lngCustomerID, Trim(pIn)
    If cInvoices.Count = 0 Then
        LoadInvoiceHeaders = 0
        Exit Function
    End If
    xInvoices.ReDim 1, cInvoices.Count, 1, 6
    For iRow = 1 To cInvoices.Count
        xInvoices.Value(iRow, 1) = iRow
        xInvoices.Value(iRow, 2) = cInvoices.Item(iRow).DocCode
        xInvoices.Value(iRow, 3) = cInvoices.Item(iRow).InvoiceDateF
        xInvoices.Value(iRow, 4) = cInvoices.Item(iRow).AmountF
        xInvoices.Value(iRow, 5) = cInvoices.Item(iRow).TRID
        xInvoices.Value(iRow, 6) = cInvoices.Item(iRow).PID
    Next iRow
    gInvoices.Array = xInvoices
    Me.gInvoices.ReBind
    LoadInvoiceHeaders = cInvoices.Count
    lngInvoiceID = FNN(xInvoices(gInvoices.Bookmark, 5))
    refreshInvoicelines FNS(xInvoices(gInvoices.Bookmark, 6))
End Function

Friend Function InvoiceReturnData(InvoiceID As Long, InvoiceCode As String, InvoiceDate As Date, InvoiceValue As Long, ra() As InvoiceRec) As Boolean
Dim i As Integer
Dim j As Integer
Dim lngTotalQty As Long

    InvoiceReturnData = True
    If IsNull(gInvoices.Bookmark) Then
        InvoiceReturnData = False
        Exit Function
    End If
    InvoiceID = lngInvoiceID
    InvoiceCode = FNS(xInvoices.Value(gInvoices.Bookmark, 2))
    InvoiceDate = CDate(xInvoices.Value(gInvoices.Bookmark, 3))
    InvoiceValue = InvoiceValue
    j = 1
    lngTotalQty = 0
    For i = 1 To xInvLines.UpperBound(1)
       ' If xInvLines(i, 6) > 0 Then
            ReDim Preserve ra(j)
            ra(j).title = xInvLines(i, 3)
            ra(j).Code = UCase(xInvLines(i, 2))
            ra(j).Qty = xInvLines(i, 6)
            ra(j).Price = xInvLines(i, 14)
            lngTotalQty = lngTotalQty + CLng(xInvLines(i, 6))
            ra(j).DiscountRate = xInvLines(i, 13)
            ra(j).PID = xInvLines(i, 12)
            ra(j).ILID = xInvLines(i, 11)
            j = j + 1
       ' End If
    Next
    InvoiceReturnData = (lngTotalQty > 0)
        
End Function
Private Sub LoadInvoicelinesArray(PID As String)
Dim i As Long

    xInvLines.Clear
    xInvLines.ReDim 1, cInvoicelines.Count, 1, 14
    For i = 1 To cInvoicelines.Count
        With cInvoicelines.Item(i)
            xInvLines.Value(i, 1) = i
            xInvLines.Value(i, 2) = .LineCode
            xInvLines.Value(i, 3) = .Description
            xInvLines.Value(i, 4) = .Qty
            xInvLines.Value(i, 5) = .QtyCredited
            If PID = .PID Then
                xInvLines.Value(i, 6) = xInvLines.Value(i, 6) + 1
            Else
                xInvLines.Value(i, 6) = xInvLines.Value(i, 6)
            End If
            xInvLines.Value(i, 7) = .PriceF
            xInvLines.Value(i, 8) = .DiscountRateF
            xInvLines.Value(i, 9) = .LineExtensionF
            xInvLines.Value(i, 11) = .ILID
            xInvLines.Value(i, 12) = .PID
            xInvLines.Value(i, 13) = .DiscountRate
            xInvLines.Value(i, 14) = .Price
        End With
    Next
    gInvLines.Array = xInvLines
    gInvLines.ReBind
End Sub

Private Sub gInvLines_AfterColUpdate(ByVal ColIndex As Integer)
    If IsNumeric(gInvLines.Text) Then
        xInvLines.Value(gInvLines.Bookmark, 6) = CLng(gInvLines.Text)
    End If
End Sub
Private Sub gInvLines_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    If Not IsNumeric(gInvLines.Text) Then
        Cancel = True
    End If
End Sub

