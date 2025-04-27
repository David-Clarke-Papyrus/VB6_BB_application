VERSION 5.00
Begin VB.Form frmPayment 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Details"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   4110
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkType 
      BackColor       =   &H00404040&
      Height          =   285
      Index           =   3
      Left            =   270
      TabIndex        =   21
      Top             =   1950
      Width           =   240
   End
   Begin VB.CheckBox chkType 
      BackColor       =   &H00404040&
      Height          =   285
      Index           =   2
      Left            =   270
      TabIndex        =   20
      Top             =   1410
      Width           =   240
   End
   Begin VB.CheckBox chkType 
      BackColor       =   &H00404040&
      Height          =   285
      Index           =   1
      Left            =   270
      TabIndex        =   19
      Top             =   885
      Width           =   240
   End
   Begin VB.CheckBox chkType 
      BackColor       =   &H00404040&
      Height          =   285
      Index           =   0
      Left            =   270
      TabIndex        =   18
      Top             =   390
      Width           =   240
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   405
      Index           =   3
      Left            =   2175
      TabIndex        =   16
      Top             =   1890
      Width           =   1605
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   405
      Index           =   2
      Left            =   2175
      TabIndex        =   14
      Top             =   1380
      Width           =   1605
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   405
      Index           =   1
      Left            =   2190
      TabIndex        =   12
      Top             =   840
      Width           =   1605
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   405
      Index           =   0
      Left            =   2175
      TabIndex        =   10
      Top             =   315
      Width           =   1605
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080C0FF&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1485
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3915
      Width           =   1095
   End
   Begin VB.Frame fraCard 
      BackColor       =   &H00404040&
      Caption         =   "Cr&edit Card Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1410
      Left            =   135
      TabIndex        =   0
      Top             =   2955
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox txtCardNumber 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   150
         TabIndex        =   5
         Top             =   840
         Width           =   2280
      End
      Begin VB.OptionButton optCardType 
         BackColor       =   &H00404040&
         Caption         =   "Amer Express"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   270
         Index           =   2
         Left            =   2115
         TabIndex        =   3
         Top             =   300
         Width           =   1470
      End
      Begin VB.OptionButton optCardType 
         BackColor       =   &H00404040&
         Caption         =   "Visa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   270
         Index           =   1
         Left            =   1260
         TabIndex        =   2
         Top             =   300
         Width           =   690
      End
      Begin VB.OptionButton optCardType 
         BackColor       =   &H00404040&
         Caption         =   "Master"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   285
         Width           =   930
      End
      Begin VB.TextBox txtExpDate 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   2580
         MaxLength       =   5
         TabIndex        =   7
         Top             =   825
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "&Card Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   585
         Width           =   1245
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "mm/yy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   195
         Index           =   7
         Left            =   2670
         TabIndex        =   9
         Top             =   1170
         Width           =   810
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "E&xp. Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   195
         Index           =   10
         Left            =   2565
         TabIndex        =   6
         Top             =   585
         Width           =   1065
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label lblBalance 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   390
      Left            =   2175
      TabIndex        =   23
      Top             =   2460
      Width           =   1605
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBalanceLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   1095
      TabIndex        =   22
      Top             =   2505
      Width           =   885
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "&Voutcher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   225
      Index           =   4
      Left            =   540
      TabIndex        =   17
      Top             =   1950
      Width           =   1245
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "C&redit Card"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   225
      Index           =   3
      Left            =   540
      TabIndex        =   15
      Top             =   1425
      Width           =   1245
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Chec&k"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   225
      Index           =   2
      Left            =   540
      TabIndex        =   13
      Top             =   900
      Width           =   1245
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "&Cash Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   1
      Left            =   540
      TabIndex        =   11
      Top             =   405
      Width           =   1515
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public UnloadOK As Boolean
Public TotalPaid As Double

Dim dTotDue As Double
Dim flgLoading As Boolean
'Dim bValid As Boolean


Public Sub Component(xPayment() As tPayment, dTotal As Double)
    On Error GoTo errHandler
Dim i As Integer

    dTotDue = dTotal
    Me.Caption = "Total Amoutn Due: " & Format(dTotal, "R 0.00")
    xPay = xPayment
    flgLoading = True
    For i = 0 To 3
        If xPay(i).Amount > 0 Then
            Me.txtAmount(i) = Format(xPay(i).Amount, "0.00")
            Me.chkType(i).Value = 1
        Else
            txtAmount(i).Text = "0.00"
            txtAmount(i).Enabled = False
        End If
    Next i
    flgLoading = False
    CheckTotals (0)
    
    'If Me.chkType(0).Value = 1 Then Me.txtAmount(0).SetFocus
'    CheckValid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.Component(xPayment,dTotal)", Array(xPayment, dTotal)
End Sub

Private Sub CheckTotals(Index As Integer)
    On Error GoTo errHandler
Dim i As Integer
Dim k As Integer
Dim dTemp As Double
    
    flgLoading = True
    'calc total allocated
    TotalPaid = 0
    For i = 0 To 3
        If Me.chkType(i).Value = 1 Then
            TotalPaid = TotalPaid + val(txtAmount(i))
        End If
    Next i
    dTemp = dTotDue - TotalPaid
    lblBalance = Format(dTemp * -1, "0.00")
'    'just allocate the missing amount to the next open field.
'    'If no others are open, allocate it back to the initial field
'    If dTemp < dTotDue Then
'        For i = 0 To 3
'            If Me.chkType(i).Value = 1 And i <> Index Then
'                txtAmount(i) = Format(dTemp, "0.00")
'                dTemp = 0
'                Exit For
'            End If
'        Next i
'    End If
'    If dTemp > 0 Then
'        txtAmount(Index) = Format(dTotDue, "0.00")
'    End If
    If val(lblBalance) >= 0 Then
        lblBalance.ForeColor = &H80FF&
        If val(lblBalance) > 0 Then
            Me.lblBalanceLbl.Caption = "Change"
        Else
            Me.lblBalanceLbl.Caption = "Balance"
        End If
    Else
        'convert negative balance value to positive value
        'Me.lblBalance = Format(Val(Me.lblBalance) * -1, "0.00")
        lblBalance.ForeColor = vbRed
        Me.lblBalanceLbl.Caption = "Short"
    End If
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.CheckTotals(Index)", Index
End Sub

Public Function Payment() As Variant
    On Error GoTo errHandler
    Payment = xPay
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.Payment"
End Function

Private Sub chkType_Click(Index As Integer)
    On Error GoTo errHandler
    If Me.chkType(Index).Value = 1 Then
        Me.txtAmount(Index).Enabled = True
        If Me.txtAmount(Index).Visible Then Me.txtAmount(Index).SetFocus
    Else
        Me.txtAmount(Index).Enabled = False
        Me.txtAmount(Index) = "0.00"
    End If
    CheckTotals Index
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.chkType_Click(Index)", Index, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    If TotalPaid < dTotDue Then
        If MsgBox("Total amount entered is less then amount due!" & vbLf & _
                  vbLf & "YES = Return to Payment input." & vbLf & _
                  "NO = Cancel sale.", vbYesNo + vbExclamation, "Amount entered not valid!") = vbYes Then
            Form_Activate
            Exit Sub
        End If
    End If
    If xPay(0).Amount <> 0 Then xPay(0).Type = "M"  'Cash
    If xPay(1).Amount <> 0 Then xPay(1).Type = "P"  'Check
    If xPay(2).Amount <> 0 Then xPay(2).Type = "C"  'Credit Card
    If xPay(3).Amount <> 0 Then xPay(3).Type = "V"  'Voucher
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    Dim i As Integer
    
    For i = 0 To 3
        If Me.chkType(i).Value = 1 Then
            If Me.txtAmount(i).Enabled And Me.txtAmount(i).Visible Then Me.txtAmount(i).SetFocus
            Exit For
        End If
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If Not UnloadOK Then
        cmdOK_Click
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub optCardType_Click(Index As Integer)
    On Error GoTo errHandler
Dim i As Integer
    For i = 0 To 2
        If xPay(i).Type = "C" Then
            Select Case Index
                Case 0
                    xPay(i).CCType = "M"    'Master
                Case 1
                    xPay(i).CCType = "V"    'Visa
                Case 2
                    xPay(i).CCType = "A"    'American Express
            End Select
        End If
    Next i
    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.optCardType_Click(Index)", Index, EA_NORERAISE
    HandleError
End Sub


Private Sub txtAmount_Change(Index As Integer)
    On Error GoTo errHandler
    
    xPay(Index).Amount = val(Me.txtAmount(Index).Text)
    
    If flgLoading Then Exit Sub
'    Select Case Index
'        Case 0
''        If Val(txtAmount(Index)) < dTotDue Then
'            txtAmount(1) = Format(dTotDue - Val(txtAmount(Index)), "0.00")
''        End If
'
'        Case 1
'            If Val(txtAmount(Index)) + Val(txtAmount(1)) < dTotDue Then
'                txtAmount(2) = Format(dTotDue - Val(txtAmount(0)) - Val(txtAmount(1)), "0.00")
'            Else
'                txtAmount(2) = "0.00"
'            End If
'        Case 2
'
'        Case 3
'
'    End Select
    CheckTotals (Index)
'    CheckValid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.txtAmount_Change(Index)", Index, EA_NORERAISE
    HandleError
End Sub

Private Sub txtAmount_GotFocus(Index As Integer)
    On Error GoTo errHandler
    AutoSelect txtAmount(Index)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.txtAmount_GotFocus(Index)", Index, EA_NORERAISE
    HandleError
End Sub

Private Sub txtAmount_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
'    CurrencyInput txtAmount(Index), KeyCode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.txtAmount_KeyDown(Index,KeyCode,Shift)", Array(Index, KeyCode, Shift), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub txtAmount_LostFocus(Index As Integer)
    On Error GoTo errHandler
    flgLoading = True
        txtAmount(Index).Text = Format(txtAmount(Index).Text, "0.00")
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.txtAmount_LostFocus(Index)", Index, EA_NORERAISE
    HandleError
End Sub

Private Sub txtCardNumber_Change()
    On Error GoTo errHandler
Dim i As Integer

    For i = 0 To 2
        If xPay(i).Type = "C" Then
            xPay(i).CCNumber = Me.txtCardNumber.Text
        End If
    Next i
'    CheckValid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.txtCardNumber_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtExpDate_Change()
    On Error GoTo errHandler
Dim i As Integer

    With txtExpDate
        If Len(.Text) = 2 And InStr(1, .Text, "/") = 0 Then
            flgLoading = True
                .Text = .Text & "/"
                .SelStart = Len(.Text)
            flgLoading = False
        End If
    End With

    For i = 0 To 2
        If xPay(i).Type = "C" Then
            xPay(i).CCExpDate = Me.txtExpDate.Text
        End If
    Next i
'    CheckValid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.txtExpDate_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtExpDate_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If Not IsNumeric(Chr(KeyAscii)) Then
        If KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And Chr(KeyAscii) <> "/" Then KeyAscii = 0
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.txtExpDate_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub txtExpDate_LostFocus()
    On Error GoTo errHandler
    If Len(Me.txtExpDate) <> 5 Then GoTo errHandler
    If (val(Me.txtExpDate) < 1) Or (val(Me.txtExpDate) > 12) _
    Or (val(Right(Me.txtExpDate, 2)) < 1) Or (val(Right(Me.txtExpDate, 2)) > 10) Then GoTo errHandler
'EH:
'
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPayment.txtExpDate_LostFocus", , EA_NORERAISE
    HandleError
End Sub

'Private Sub CheckValid()
''Dim i As Integer
''Dim dSum As Double
''    For i = 0 To 2
''        dSum = dSum + xPay(i).Amount
''    Next i
''    If dSum >= dTotDue Then
''        If Me.fraCard.Visible Then
'''            Me.cmdOK.Enabled = Len(Me.txtCardNumber) > 8 And Len(Me.txtExpDate) = 5
''            Me.cmdOK.Enabled = Len(Me.txtCardNumber) > 3 And Len(Me.txtExpDate) = 5
''        Else
''            Me.cmdOK.Enabled = True
''        End If
''    Else
''        Me.cmdOK.Enabled = False
''
''    End If
'    Me.cmdOK.Enabled = lblBalance.ForeColor <> vbRed
'End Sub

