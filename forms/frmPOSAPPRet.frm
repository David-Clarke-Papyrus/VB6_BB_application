VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmPOSAPPRet 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Appros outstanding"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   9975
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   8670
      TabIndex        =   6
      Top             =   6270
      Width           =   1305
      Begin VB.Label lblNominal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   60
         TabIndex        =   7
         Top             =   180
         Width           =   1155
      End
   End
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
   Begin TrueOleDBGrid60.TDBGrid gAppLines 
      Height          =   2940
      Left            =   120
      OleObjectBlob   =   "frmPOSAPPRet.frx":0000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2055
      Width           =   9510
   End
   Begin TrueOleDBGrid60.TDBGrid gApps 
      Height          =   1770
      Left            =   135
      OleObjectBlob   =   "frmPOSAPPRet.frx":6142
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.Label SB 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Height          =   420
      Left            =   105
      TabIndex        =   8
      Top             =   6345
      Width           =   8430
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
Attribute VB_Name = "frmPOSAPPRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xAPPs As XArrayDB
Dim xAPPLines As XArrayDB
Dim cApps As c_APPs
Dim cAppLines As c_APPLines
Dim lngCustomerID As Long
Dim lngAPPID As Long
Dim strMode As String
Dim bCancelled As Boolean
Dim lngNominalPrice As Long
Dim strNominalPrice As String

Public Sub component(pTPID As Long)
    lngCustomerID = pTPID
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
    Set xAPPs = New XArrayDB
    Set xAPPLines = New XArrayDB
    
    Set cApps = New c_APPs
    cApps.Load lngCustomerID
    LoadAppros
    Me.SB.Caption = "Scan product,(C)Correct,(F)Finish"
    strMode = "R"
    lblInput.Caption = "Scan returned items"
    lblInput.ForeColor = &H714942
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cApps = Nothing
    Set xAPPs = Nothing
    Set xAPPLines = Nothing
End Sub

Private Sub LoadAppros()
    On Error GoTo errHandler
Dim i As Long

    gApps.Visible = True
    xAPPs.Clear
    xAPPs.ReDim 1, cApps.Count, 1, 5
    For i = 1 To cApps.Count
        With cApps(i)
            xAPPs.Value(i, 1) = i
            xAPPs.Value(i, 2) = .DocCode
            xAPPs.Value(i, 3) = .DOCDate
            xAPPs.Value(i, 4) = .APPID
        End With
    Next
    xAPPs.QuickSort 1, xAPPs.UpperBound(1), 3, XORDER_DESCEND, XTYPE_NUMBER
    
    gApps.Array = xAPPs
    Me.gApps.ReBind
    If Not IsNull(gApps.Bookmark) Then
        lngAPPID = FNN(xAPPs(gApps.Bookmark, 4))
        refreshApprolines
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadCOLS"
End Sub

Private Sub refreshApprolines()
    Set cAppLines = Nothing
    Set cAppLines = New c_APPLines
    cAppLines.Load lngAPPID
    LoadApproLines
End Sub

Private Sub LoadApproLines()
Dim i As Long
Dim lngToBeInvoiced As Long

    xAPPLines.Clear
    xAPPLines.ReDim 1, cAppLines.Count, 1, 16
    For i = 1 To cAppLines.Count
        With cAppLines(i)
            xAPPLines.Value(i, 1) = i
            xAPPLines.Value(i, 2) = .CodeF
            xAPPLines.Value(i, 3) = .title
            xAPPLines.Value(i, 4) = .QtyOut
            xAPPLines.Value(i, 5) = .QtyBack
            lngToBeInvoiced = (xAPPLines.Value(i, 4) - xAPPLines.Value(i, 5))
            xAPPLines.Value(i, 6) = lngToBeInvoiced
            xAPPLines.Value(i, 7) = .PriceF
            xAPPLines.Value(i, 8) = .DiscountRateF
            xAPPLines.Value(i, 10) = .Price
            xAPPLines.Value(i, 11) = .DiscountRate
            xAPPLines.Value(i, 12) = CalcExt(lngToBeInvoiced, xAPPLines.Value(i, 10), xAPPLines.Value(i, 11))
            xAPPLines.Value(i, 9) = Format(xAPPLines.Value(i, 12) / oPC.CurrencyDivisor, oPC.CurrencyFormat)
            xAPPLines.Value(i, 13) = .APPLID
            xAPPLines.Value(i, 14) = .VATRate
            xAPPLines.Value(i, 15) = .PID
            xAPPLines.Value(i, 16) = .Code
        End With
    Next
    gAppLines.Array = xAPPLines
    gAppLines.ReBind
    lblTotal.Caption = InvoiceValueF
End Sub
Private Function InvoiceValue() As Long
Dim lngTotal As Long
Dim lngToBeInvoiced As Long
Dim i As Integer
    
    lngNominalPrice = 0
    For i = 1 To xAPPLines.UpperBound(1)
        lngToBeInvoiced = (xAPPLines.Value(i, 4) - xAPPLines.Value(i, 5))
        lngTotal = lngTotal + CalcExt(lngToBeInvoiced, xAPPLines.Value(i, 10), xAPPLines.Value(i, 11))
        lngNominalPrice = lngNominalPrice + CalcExtNominal(lngToBeInvoiced, xAPPLines.Value(i, 10))
    Next i
    InvoiceValue = lngTotal
    If lngNominalPrice > 0 Then
        Me.lblNominal.Caption = CalcExtNominalF
    Else
        Me.lblNominal.Caption = ""
    End If
    
End Function
Private Function InvoiceValueF() As String
    InvoiceValueF = Format(InvoiceValue / oPC.CurrencyDivisor, oPC.CurrencyFormat)
End Function
Private Function CalcExt(pQty As Long, pPrice As Long, pDiscountRate As Long) As Long
    CalcExt = (pPrice * CDbl((100 - pDiscountRate) / 100)) * pQty
End Function
Private Function CalcExtF(pQty As Long, pPrice As Long, pDiscountRate As Long) As String
    CalcExtF = Format(CalcExt(pQty, pPrice, pDiscountRate) / oPC.CurrencyDivisor, oPC.CurrencyFormat)
End Function
Private Function CalcExtNominal(pQty As Long, pPrice As Long) As Long
    CalcExtNominal = pPrice * pQty
End Function
Public Function CalcExtNominalF() As String
    CalcExtNominalF = Format(lngNominalPrice / oPC.CurrencyDivisor, oPC.CurrencyFormat)
End Function
Private Sub gApps_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If IsNull(gApps.Bookmark) Then Exit Sub
    lngAPPID = FNN(xAPPs(gApps.Bookmark, 4))
    refreshApprolines
End Sub
Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler

    If KeyCode = 13 Then
        ActionInput txtInput
        If gAppLines.Row <> -1 Then gAppLines.MoveRelative 0, gAppLines.Bookmark
     End If
       
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSAPPRet.txtInput_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub ActionInput(pIn As String)
    pIn = (UCase(FNS(pIn)))
    If pIn = "F" Then
        Me.Hide
    ElseIf FindRow(pIn) > 0 Then
        txtInput = ""
    ElseIf pIn = "C" Then
        If MsgBox("You want to cancel this Appro return?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
            bCancelled = True
            Me.Hide
            Exit Sub
        End If
    
'        strMode = "C"
'        lblInput.Caption = "Correction"
'        lblInput.BackColor = vbRed
    Else
        MsgBox "There is no product with this code on the appro to be returned."
    End If
End Sub
Private Function FindRow(pIn As String) As Long
Dim i As Integer
Dim iRow As Integer
Dim lngToBeInvoiced As Long

    iRow = 0
    If xAPPLines.Count(1) = 0 Then
        FindRow = 0
        Exit Function
    End If
    For i = 1 To xAPPLines.UpperBound(1)
        If UCase(xAPPLines(i, 2)) = UCase(pIn) Or UCase(xAPPLines(i, 16)) = UCase(pIn) Then
            iRow = i
            Exit For
        End If
    Next
    If iRow = 0 Then
        FindRow = 0
        Exit Function
    End If
    If strMode = "R" Then
        If (xAPPLines.Value(i, 4) - xAPPLines.Value(i, 5)) > 0 Then
            xAPPLines(iRow, 5) = xAPPLines(iRow, 5) + 1
            lngToBeInvoiced = (xAPPLines.Value(i, 4) - xAPPLines.Value(i, 5))
            xAPPLines.Value(i, 6) = lngToBeInvoiced
            xAPPLines.Value(i, 12) = CalcExt(lngToBeInvoiced, xAPPLines.Value(i, 10), xAPPLines.Value(i, 11))
            xAPPLines.Value(i, 9) = Format(xAPPLines.Value(i, 12) / oPC.CurrencyDivisor, oPC.CurrencyFormat)
            gAppLines.ReBind
            lblTotal.Caption = InvoiceValueF
        Else
            iRow = 0
        End If
    ElseIf strMode = "C" Then 'Correction
        If xAPPLines(iRow, 5) > 0 Then
            xAPPLines(iRow, 5) = xAPPLines(iRow, 5) - 1
            lngToBeInvoiced = (xAPPLines.Value(i, 4) - xAPPLines.Value(i, 5))
            xAPPLines.Value(i, 6) = lngToBeInvoiced
            xAPPLines.Value(i, 12) = CalcExt(lngToBeInvoiced, xAPPLines.Value(i, 10), xAPPLines.Value(i, 11))
            xAPPLines.Value(i, 9) = Format(xAPPLines.Value(i, 12) / oPC.CurrencyDivisor, oPC.CurrencyFormat)
            gAppLines.ReBind
            lblTotal.Caption = InvoiceValueF
        Else
            iRow = 0
        End If
    End If
    FindRow = iRow
End Function

Friend Function ApproReturnData(APPID As Long, APPCode As String, APPDate As Date, AppValue As Long, ra() As ReturnRec) As Boolean
Dim i As Integer
Dim j As Integer

    ApproReturnData = True
    If IsNull(gApps.Bookmark) Then
        ApproReturnData = False
        Exit Function
    End If
    APPID = lngAPPID
    APPCode = FNS(xAPPs.Value(gApps.Bookmark, 2))
    APPDate = CDate(xAPPs.Value(gApps.Bookmark, 3))
    AppValue = InvoiceValue
    j = 1
    For i = 1 To xAPPLines.UpperBound(1)
       ' If xAPPLines(i, 6) > 0 Then
            ReDim Preserve ra(j)
            ra(j).title = xAPPLines(i, 3)
            ra(j).Code = UCase(xAPPLines(i, 16))
            ra(j).APPLQtySold = xAPPLines(i, 6)
            ra(j).APPLID = xAPPLines(i, 13)
            ra(j).APPLQtyReturned = xAPPLines(i, 5)
            ra(j).PID = xAPPLines(i, 15)
            ra(j).Price = xAPPLines(i, 10)
            ra(j).VATRate = xAPPLines(i, 14)
            ra(j).DiscountRate = xAPPLines(i, 11)
            j = j + 1
       ' End If
    Next
    
End Function
