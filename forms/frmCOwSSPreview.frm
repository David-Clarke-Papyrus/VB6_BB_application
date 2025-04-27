VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmCOPreview 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D3D3CB&
   Caption         =   "Order"
   ClientHeight    =   5820
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11340
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmCOwSSPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCreateFinalInvoice 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Create final invoice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   4020
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Create invoice for whole order"
      Top             =   4605
      Width           =   1845
   End
   Begin VB.CommandButton cmdStatus 
      BackColor       =   &H00FFC0C0&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5295
      Width           =   255
   End
   Begin VB.Frame frHeader 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Heading data"
      ForeColor       =   &H8000000D&
      Height          =   3105
      Left            =   2985
      TabIndex        =   23
      Top             =   2385
      Visible         =   0   'False
      Width           =   5715
      Begin VB.TextBox txtForAttn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1020
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2130
         Width           =   3240
      End
      Begin VB.TextBox txtTPMemo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   945
         Left            =   1020
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1050
         Width           =   4440
      End
      Begin VB.TextBox txtRef 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1020
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   630
         Width           =   3240
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "For attn."
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   195
         TabIndex        =   29
         Top             =   2190
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Memo"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   195
         TabIndex        =   27
         Top             =   1095
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ref"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   195
         TabIndex        =   26
         Top             =   690
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdRef 
      BackColor       =   &H00FFC0C0&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4935
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdMemo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4620
      Width           =   255
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2520
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCOwSSPreview.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close the order"
      Top             =   4605
      Width           =   1000
   End
   Begin VB.TextBox txtCurrency 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00706034&
      Height          =   255
      Left            =   9450
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   420
      Width           =   1635
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1470
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCOwSSPreview.frx":2B2C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print or preview"
      Top             =   4605
      Width           =   1000
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   465
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCOwSSPreview.frx":2EB6
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print the invoice"
      Top             =   4605
      Width           =   1000
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   9555
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   165
      Width           =   1545
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   3000
      Left            =   45
      OleObjectBlob   =   "frmCOwSSPreview.frx":3240
      TabIndex        =   9
      Top             =   1515
      Width           =   10815
   End
   Begin CoolButtonControl.CoolButton cbDelto 
      Height          =   1425
      Left            =   6570
      TabIndex        =   10
      Top             =   0
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   2514
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      BackStyle       =   0
   End
   Begin CoolButtonControl.CoolButton cbBillTo 
      Height          =   1425
      Left            =   4020
      TabIndex        =   11
      Top             =   0
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   2514
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      BackStyle       =   0
   End
   Begin CoolButtonControl.CoolButton cbCust 
      Height          =   1425
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   2514
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      ShowFocusRect   =   -1  'True
      BackStyle       =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Height          =   450
      Left            =   465
      TabIndex        =   7
      Top             =   5250
      Width           =   1590
      Begin VB.Label lblOS 
         Alignment       =   2  'Center
         BackColor       =   &H00DBFAFB&
         Caption         =   "OS"
         Height          =   285
         Left            =   510
         TabIndex        =   21
         Top             =   135
         Width           =   345
      End
      Begin VB.Label lblFUL 
         Alignment       =   2  'Center
         BackColor       =   &H00FEABAD&
         Caption         =   "FUL"
         Height          =   285
         Left            =   840
         TabIndex        =   20
         Top             =   135
         Width           =   345
      End
      Begin VB.Label lblCAN 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "CAN"
         Height          =   285
         Left            =   1215
         TabIndex        =   19
         Top             =   135
         Width           =   345
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Key:"
         Height          =   300
         Left            =   90
         TabIndex        =   8
         Top             =   150
         Width           =   345
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Goods to:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6855
      TabIndex        =   18
      Top             =   30
      Width           =   1050
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill to:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4185
      TabIndex        =   17
      Top             =   45
      Width           =   660
   End
   Begin VB.Label lblBillToAddress 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   940
      Left            =   4185
      TabIndex        =   16
      Top             =   300
      Width           =   2055
   End
   Begin VB.Label lblDelToAddress 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   940
      Left            =   6770
      TabIndex        =   15
      Top             =   300
      Width           =   2055
   End
   Begin VB.Label lblTPName 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   75
      TabIndex        =   14
      Top             =   135
      Width           =   3600
      WordWrap        =   -1  'True
   End
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   10185
      X2              =   11655
      Y1              =   15
      Y2              =   840
   End
   Begin VB.Label lblTotalCaption 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
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
      Height          =   915
      Left            =   5445
      TabIndex        =   4
      Top             =   4575
      Width           =   3495
   End
   Begin VB.Label lblTotalValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
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
      Height          =   915
      Left            =   9015
      TabIndex        =   3
      Top             =   4575
      Width           =   1845
   End
End
Attribute VB_Name = "frmCOPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCO As c_COs
Dim oCO As a_CO
Dim dblTotal As Double
Dim XA As XArrayDB
Dim flgLoading As Boolean
Dim bMemoExpanded As Boolean
Dim mbShowMemo As Boolean
Dim mbShowLog As Boolean
Dim mbShowRef As Boolean
Dim oSM As z_StockManager

Dim PrintCommandButtonCTRLDown As Boolean
Private Sub Form_Initialize()
    PrintCommandButtonCTRLDown = False
End Sub
Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim ShiftTest As Integer
   PrintCommandButtonCTRLDown = False
   ShiftTest = Shift And 7
   Select Case ShiftTest
      Case 1 ' or vbShiftMask
      Case 2 ' or vbCtrlMask
         PrintCommandButtonCTRLDown = True
      End Select
End Sub

Private Sub cmdPrint_KeyUp(KeyCode As Integer, Shift As Integer)
        PrintCommandButtonCTRLDown = False
End Sub
Public Sub component(PID As Long, bForInvoicing As Boolean)
    On Error GoTo errHandler
Dim lngID As Long
    lngID = PID
    Set oCO = New a_CO
    oCO.Load lngID, True
    If oCO.OrderType = enWant Then
        Me.Caption = "Wants for " & oCO.Customer.Fullname & oCO.StaffNameB
    ElseIf oCO.OrderType = enNormalCO Then
       ' Me.Caption = "Order from " & oCO.Customer.Fullname & oCO.StaffNameB
        Me.Caption = "Sales order " & "  " & oCO.DOCCode & "    " & oCO.DOCDate & " " & oCO.StaffNameB & IIf(oCO.OrderRef > "", "  (ref:" & oCO.OrderRef & ")", "") & "   " & oCO.DOCCode
    End If
    If DateDiff("d", oCO.DOCDate, oCO.IssDate) > 1 Then
        Me.Caption = Me.Caption & "   Issued: " & oCO.IssDateF
    End If
    flgLoading = True
    LoadControls
    SetMenu
    lblOS.BackColor = COLOR_PALEYELLOW
    lblFUL.BackColor = COLOR_FULFILLED
    lblCAN.BackColor = COLOR_CANCELLED
    Me.cmdCreateFinalInvoice.Visible = bForInvoicing
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.component(PID)", PID
End Sub
Public Sub ComponentObject(pobj As a_CO)
    On Error GoTo errHandler
    Set oCO = pobj
    Me.Caption = "Order from " & oCO.TPNAME
    If oCO.OrderType = enWant Then
        Me.Caption = "Wants for " & oCO.Customer.Fullname & oCO.StaffNameB
    ElseIf oCO.OrderType = enNormalCO Then
       ' Me.Caption = "Order from " & oCO.Customer.Fullname & oCO.StaffNameB
        Me.Caption = "Sales order " & "  " & oCO.DOCCode & "    " & oCO.DOCDate & " " & oCO.StaffNameB & IIf(oCO.OrderRef > "", "  (ref:" & oCO.OrderRef & ")", "") & "   " & oCO.DOCCode
    End If
    If DateDiff("d", oCO.DOCDate, oCO.IssDate) > 1 Then
        Me.Caption = Me.Caption & "   Issued: " & oCO.IssDateF
    End If
    SetMenu
    flgLoading = True
    LoadControls
    flgLoading = False
    cmdCreateFinalInvoice.Visible = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.ComponentObject(pobj)", pobj
End Sub

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.G1, Me.Name, Me.Height, Me.Width
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.mnuSaveLayout"
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oCO.StatusF = "IN PROCESS" And oCO.IsNew = False)
    Forms(0).mnuCancel.Enabled = (oCO.StatusF = "ISSUED")
    Forms(0).mnuCancelLine.Enabled = (oCO.StatusF = "ISSUED" And oCO.IsNew = False)
    Forms(0).mnuFulfil.Enabled = (oCO.StatusF = "ISSUED") 'And oCO.CanCancel = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = False
   ' Forms(0).mnuCopyDoc.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Forms(0).mnuCopyLines.Enabled = True
    Forms(0).mnuPastelines.Enabled = True
    Forms(0).mnuPastelinestoNEW = True
    If oPC.EmailCO And (oCO.Status = stCOMPLETE Or oCO.Status = stISSUED) Then
        Forms(0).mnuEmail.Enabled = False
        Forms(0).mnuOutlook.Enabled = False
        If Not oCO.Customer.BillTOAddress Is Nothing Then
            If (oCO.Customer.DispatchMethod = "M" And oCO.Customer.BillTOAddress.EMail > "") Then
                Forms(0).mnuEmail.Enabled = Not oPC.UsesOutlookForCOEmail
                Forms(0).mnuOutlook.Enabled = oPC.UsesOutlookForCOEmail
            End If
        End If
    Else
        Forms(0).mnuEmail.Enabled = False
        Forms(0).mnuOutlook.Enabled = False
    End If
   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.SetMenu"
End Sub
Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
Dim oSM As New z_StockManager
    
    ofrm.component oCO.Memo
    ofrm.Show vbModal
    oSM.SetMemo ofrm.Memo, oCO.TRID
    
    txtTPMemo.Visible = (ofrm.Memo > "")
    txtTPMemo = "Note: " & ofrm.Memo
    oSM.SetMemo ofrm.Memo, oCO.TRID
    oCO.SetMemo ofrm.Memo
    
    Unload ofrm

    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.mnuMemo"
End Sub


Public Sub mnuFulfilLine()
    On Error GoTo errHandler
Dim oCOL As a_COL
    Set oCOL = oCO.COLines.FindLineByID(val(XA(G1.Bookmark, 24)))
            If MsgBox("Do you wish to mark the selected line as fulfilled?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
            Screen.MousePointer = vbHourglass
            oCO.COLines.FindLineByID(val(XA(G1.Bookmark, 24))).FulfilLine
            RefreshData
            Screen.MousePointer = vbDefault
    '    End If
  '  End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.mnuFulfilLine"
End Sub
Public Sub mnuCancelLine()
    On Error GoTo errHandler
Dim oP As a_Product
Dim str As String
Dim oCOL As a_COL
    Set oCOL = oCO.COLines.FindLineByID(val(XA(G1.Bookmark, 24)))
    If oCOL.QtyDispatched > 0 Then
        MsgBox "This line is partially fulfilled and can only be marked as fulfilled.", vbExclamation + vbOKOnly, "Can't do this"
    Else
    
        If oPC.Configuration.SignTransactions = True Then
            If SecurityControl(enSECURITY_CO_SIGN, , "Void this customer order", DOCAPPROVAL) = False Then
                   Exit Sub
            End If
        Else
            If oCO.Status = stInProcess Then
                If MsgBox("Do you want to cancel the selected line?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    
    
    
      '  If MsgBox("Do you wish to cancel the selected line?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
        Screen.MousePointer = vbHourglass
        Set oP = New a_Product
        oCO.COLines.FindLineByID(val(XA(G1.Bookmark, 24))).CancelLine
        RefreshData
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.mnuCancelLine"
End Sub


Private Sub cbBillTo_Click()
    On Error GoTo errHandler
Static iBillIdx As Integer
START:
    If oCO.Customer.ID = 0 Then Exit Sub
    If iBillIdx = 0 Then
        iBillIdx = setCurrentAddressIndex("BILL")
    End If
    iBillIdx = iBillIdx + 1
    If iBillIdx > oCO.Customer.Addresses.Count Then
        iBillIdx = 1
    End If
    lblBillToAddress.Caption = oCO.Customer.Addresses(iBillIdx).AddressMailing & vbCrLf & oCO.Customer.Addresses(iBillIdx).EMail
    oCO.SetBillToAddressImmediate oCO.Customer.Addresses(iBillIdx)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.cbBillTo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cbDelTo_Click()
    On Error GoTo errHandler
Static iDelIdx As Integer
START:
    If oCO.Customer.ID = 0 Then Exit Sub
    If iDelIdx = 0 Then iDelIdx = setCurrentAddressIndex("DEL")
    iDelIdx = iDelIdx + 1
    If iDelIdx > oCO.Customer.Addresses.Count Then
        iDelIdx = 1
    End If
    lblDelToAddress.Caption = oCO.Customer.Addresses(iDelIdx).AddressMailing & vbCrLf & oCO.Customer.Addresses(iDelIdx).EMail
    oCO.setDelToAddressImmediate oCO.Customer.Addresses(iDelIdx)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.cbDelTo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Function setCurrentAddressIndex(pType As String) As Integer
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To oCO.Customer.Addresses.Count
        If pType = "BILL" And Not oCO.BillTOAddress Is Nothing Then
            If oCO.BillTOAddress.ID = oCO.Customer.Addresses(i).ID Then
                setCurrentAddressIndex = i
            End If
        ElseIf pType = "DEL" And Not oCO.DelToAddress Is Nothing Then
            If oCO.DelToAddress.ID = oCO.Customer.Addresses(i).ID Then
                setCurrentAddressIndex = i
            End If
        End If
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.setCurrentAddressIndex(pType)", pType
End Function



Private Sub cmdStatus_Click()
    On Error GoTo errHandler
'Dim frm As New frmCustomerOrderStatusReport
'
'    frm.component oCO.TRID
'    frm.Show vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.cmdStatus_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdRef_Click()
'    ShowRef Not mbShowRef
'End Sub
'Private Sub ShowRef(bOn As Boolean)
'    On Error Resume Next
'        mbShowRef = bOn
'        txtRef.Visible = bOn
'        If bOn Then txtRef.SetFocus
'End Sub
'
'
'
'Private Sub cmdRef_Click()
'    ShowRef Not mbShowRef
'End Sub
'Private Sub ShowRef(bOn As Boolean)
'    On Error Resume Next
'        mbShowRef = bOn
'        txtRef.Visible = bOn
'        If bOn Then txtRef.SetFocus
'End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadControls()
    On Error GoTo errHandler
Dim dblVAT As Double
Dim dblConversionRate As Double
Dim strCurrencyFormat As String
Dim curTotalDeposits As Currency
Dim curTotalValue As Currency
Dim strAddress As String
Dim strTotalCaption As String
Dim strTotalValues As String
    
        With oCO
            Me.txtStatus = .StatusF
            CancelLine.Visible = (.Status = stCANCELLED Or .Status = stVOID)
            If oPC.GetProperty("CanEditCOs") = "TRUE" Then
                If .Status = stInProcess Or .Status = stISSUED Or .OrderType = enWant Then
                    cmdEdit.Enabled = True
                Else
                    cmdEdit.Enabled = False
                End If
            Else
                If .Status = stInProcess Or .OrderType = enWant Then
                    cmdEdit.Enabled = True
                Else
                    cmdEdit.Enabled = False
                End If
            End If
            lblTPName.Caption = .Customer.Fullname & IIf(Len(.TPACCNum) > 0, " (" & .TPACCNum & ")", "")
            If Not (.Customer.BillTOAddress Is Nothing) Then
                lblTPName.Caption = Me.lblTPName.Caption & vbCrLf & "Phone: " & .Customer.BillTOAddress.Phone & vbCrLf & "Fax: " & .Customer.BillTOAddress.Fax
            End If
            Me.txtRef = .Ref
            Me.txtTPMemo = IIf(Len(.Memo) > 0, .Memo, "")
            Me.txtForAttn = .ForAttn
            If .BillToAddressID > 0 Then
                If Not .BillTOAddress Is Nothing Then strAddress = .BillTOAddress.AddressMailing
            End If
            Me.lblBillToAddress.Caption = IIf(strAddress > "", strAddress, "unknown")
            If .GoodsAddressID > 0 Then
                If Not .DelToAddress Is Nothing Then strAddress = .DelToAddress.AddressMailing
            End If
            Me.lblDelToAddress.Caption = IIf(strAddress > "", strAddress, "unknown")
            .CalculateTotal
            .DisplayTotals strTotalCaption, strTotalValues, False
            lblTotalCaption.Caption = strTotalCaption
            lblTotalValues.Caption = strTotalValues
        End With
   '     LoadListView
        LoadGrid
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.LoadControls"
End Sub


Private Sub cbCust_Click()
    On Error GoTo errHandler
Dim frm As New frmCustomerPreview
    If flgLoading Then Exit Sub
    frm.component oCO.Customer
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.cbCust_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_CO
Dim oDOC As a_DocumentControl
Dim qtyLinesToPrint As Integer
Dim Dummy As String

    If PrintCommandButtonCTRLDown Then
        PrintCommandButtonCTRLDown = False

        Screen.MousePointer = vbHourglass
        oCO.COLines.SortInvoiceLines enSequence, True

        Set oDOC = oPC.Configuration.DocumentControls.FindDC(oCO.constDOCCODE)
        If oDOC Is Nothing Then
            qtyLinesToPrint = 1
        Else
            qtyLinesToPrint = oPC.Configuration.DocumentControls.FindDC(oCO.constDOCCODE).QtyCopies
        End If

       If oCO.ExportToXML(Dummy, enView, , , qtyLinesToPrint, True) = False Then
           Screen.MousePointer = vbDefault
           MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
       End If
       Screen.MousePointer = vbDefault
    Else
        Set frm = New frmPrintingOptions_CO
        frm.ComponentObject oCO
        frm.Show vbModal
    End If
EXIT_Handler:
 '   Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim blnEdit As Boolean
Dim frm As frmCO
Dim bCancel As Boolean
Dim lCO As New a_CO
    WaitMsg "Loading . . .", True, Me
    lCO.Load oCO.TRID, False
    Set frm = New frmCO
    blnEdit = True
    frm.component bCancel, lCO
    frm.Show
    WaitMsg "", False, Me

EXIT_Handler:
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadGrid()
    On Error Resume Next
Dim lstItem As ListItem
Dim i As Integer
Dim currDeposit As Currency
Dim currPrice As Currency
Dim dblVAT As Double
Dim strSummaryDescription As String
Dim strSummary As String
Dim lngTotal As Long
Dim lngDepositTotal As Long
Dim strExtra As String

    Set XA = New XArrayDB
    XA.Clear
    G1.Columns(9).Width = 0
    If oCO.OrderType = enWant Then
        G1.Columns(2).Width = 4000
        G1.Columns(3).Width = 1500
        G1.Columns(4).Width = 3500
        G1.Columns(5).Width = 0
        G1.Columns(6).Width = 0
        G1.Columns(7).Width = 0
        G1.Columns(8).Width = 0
        G1.Columns(9).Width = 0
    End If
    XA.ReDim 1, oCO.COLines.Count, 1, 25
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), G1.Columns(i - 1).Width)
    Next
  On Error GoTo errHandler
    For i = 1 To oCO.COLines.Count
             XA(i, 1) = oCO.COLines(i).CodeF & oCO.COLines(i).HasSupplierID
             XA(i, 2) = oCO.COLines(i).TitleAuthorPublisher
             If oCO.OrderType = enWant Then
                 XA(i, 3) = oCO.COLines(i).WantDateF
                 G1.Columns(2).Caption = "Date of want"
                 G1.Columns(3).Caption = "Note"
                 XA(i, 4) = oCO.COLines(i).Note
             Else
                 XA(i, 3) = oCO.COLines(i).Ref
                 G1.Columns(2).Caption = "Ref"
                 XA(i, 4) = oCO.COLines(i).QtyOrdered_QtyDispatched
            End If
             If oCO.COLines(i).Deposit > 0 Then
                 XA(i, 5) = oCO.COLines(i).DepositF
                 XA(i, 6) = oCO.COLines(i).DepositStatus
             Else
                 XA(i, 5) = " "
                 XA(i, 6) = " "
             End If
             XA(i, 7) = oCO.COLines(i).PriceF
             XA(i, 8) = oCO.COLines(i).DiscountF
             XA(i, 9) = oCO.COLines(i).ExtensionF
             XA(i, 10) = oCO.COLines(i).ExtraChargeDescription
             XA(i, 11) = oCO.COLines(i).ExtraChargeF
             XA(i, 13) = oCO.COLines(i).DeliveryDocument
             XA(i, 20) = oCO.COLines(i).Fulfilled
             XA(i, 21) = oCO.COLines(i).Key
             XA(i, 22) = oCO.COLines(i).code
             XA(i, 24) = oCO.COLines(i).POLID
             XA(i, 24) = oCO.COLines(i).COLineID
             XA(i, 25) = oCO.COLines(i).EAN
            strExtra = ""
            If oCO.COLines(i).ETAF > "" Then
                strExtra = oCO.COLines(i).ETAF
            End If
            If oCO.COLines(i).Note > "" Then
                strExtra = strExtra & " " & oCO.COLines(i).Note
            End If
            If oCO.COLines(i).LastAction > "" Then
                strExtra = strExtra & " Action:" & oCO.COLines(i).LastAction
            End If
            XA(i, 12) = strExtra
    Next i
    
    G1.Array = XA
    G1.ReBind

    
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.LoadGrid"
End Sub


Private Sub cmdCreateFinalInvoice_Click()
    On Error GoTo errHandler
Dim sResult As String
Dim oIG As New Z_InvoiceGeneration

    oIG.GetAlreadyInvoicedLines oCO.TRID, sResult
    If sResult > "" Then
    End If
    oIG.CreateFinalInvoiceFromDeliveredCustomerOrder oCO.TRID, stInProcess
    MsgBox "Invoice(s) generated", vbInformation + vbOKOnly, "Status"
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.cmdCreateFinalInvoice_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = vbKey4 Then
        If MsgBox("Confirm close?", vbOKCancel, "Close form") = vbOK Then
            Unload Me
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.Form_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    If Me.WindowState <> 2 Then
       Me.TOP = 50
        Me.Left = 50
    End If
    SetFormSize Me
    SetGridLayout G1, Me.Name
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    G1.Width = NonNegative_Lng(Me.Width - (G1.Left + 400))
    lngDiff = G1.Height
    G1.Height = NonNegative_Lng(Me.Height - (G1.TOP + 1660))
    lngDiff = (G1.Height - lngDiff)
    cmdEdit.TOP = cmdEdit.TOP + lngDiff
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    cmdClose.TOP = cmdClose.TOP + lngDiff
   ' txtTPMemo.Top = txtTPMemo.Top + lngDiff
    cmdMemo.TOP = cmdMemo.TOP + lngDiff
    cmdRef.TOP = cmdRef.TOP + lngDiff
    Frame1.TOP = Frame1.TOP + lngDiff
    lblTotalCaption.TOP = lblTotalCaption.TOP + lngDiff
    lblTotalValues.TOP = lblTotalValues.TOP + lngDiff
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set oCO = Nothing
   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Private Sub Label5_DblClick()
    On Error GoTo errHandler
Dim frm As frmCustomerPreview
    Set frm = New frmCustomerPreview
    frm.component oCO.Customer
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.Label5_DblClick", , EA_NORERAISE
    HandleError
End Sub


Private Sub G1_Click()
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 25)) > "", FNS(XA.Value(G1.Bookmark, 25)), FNS(XA.Value(G1.Bookmark, 21)))
    If str = "" Then Exit Sub
    On Error Resume Next
    
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.G1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub G1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 21))
    If str = "" Then Exit Sub
    Forms(0).mnuCancelLine.Enabled = oCO.COLines(str).QtyDispatched = 0 And (oCO.StatusF = "ISSUED" And oCO.IsNew = False)
    str = IIf(FNS(XA.Value(G1.Bookmark, 25)) > "", FNS(XA.Value(G1.Bookmark, 25)), FNS(XA.Value(G1.Bookmark, 21)))
    If str = "" Then Exit Sub
    On Error Resume Next
    
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.G1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub G1_SelChange(Cancel As Integer)
    On Error GoTo errHandler
Dim str As String

    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 21))
    If str = "" Then Exit Sub
    Forms(0).mnuCancelLine.Enabled = oCO.COLines(str).QtyDispatched = 0 And (oCO.StatusF = "ISSUED" And oCO.IsNew = False)
    str = FNS(XA.Value(G1.Bookmark, 22))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.G1_SelChange(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Public Sub mnuCancel()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    
    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_CO_SIGN, , "Void this customer order", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    Else
        If oCO.Status = stInProcess Then
            If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    
  ' If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oSM.CancelCO oCO
    RefreshData
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.mnuCancel"
End Sub

Private Sub G1_DblClick()
    On Error GoTo errHandler
Dim frm As frmProductPrev
Dim frmA As frmProductPrevAQ
Dim oP As a_Product
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 21))
    If str = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oP.Load oCO.COLines(str).PID, 0
    If oPC.Configuration.AntiquarianYN Then
        Set frmA = New frmProductPrevAQ
        frmA.component oP
        frmA.Show
    Else
        Set frm = New frmProductPrev
        frm.component oP
        frm.Show
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
  errRepeat = errRepeat + 1
  LogSaveToFile "Access violation in frmCOPreview: G1_DblClick"  'unknown source
  If errRepeat < 5 Then
      Resume Next
  Else
      LogSaveToFile "Access violation in frmCOPreview: G1_DblClick after 5 re-attempts"
      MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
      Err.Clear
      Exit Sub
  End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.G1_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuFileExit_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.mnuFileExit_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub G1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    G1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 3, 10
            GetRowType = XTYPE_STRING
        Case 4, 5, 6, 7, 8, 9
            GetRowType = XTYPE_INTEGER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.GetRowType(ColIndex)", ColIndex
End Function


Public Sub mnuVoid()
    On Error GoTo errHandler
    
        If oPC.Configuration.SignTransactions = True Then
            If SecurityControl(enSECURITY_CO_SIGN, , "Void this customer order", DOCAPPROVAL) = False Then
                   Exit Sub
            End If
        Else
            If oCO.Status = stInProcess Then
                If MsgBox("Do you want to void the selected line?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    
   ' If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oCO.VoidDocument
    RefreshData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.mnuVoid"
End Sub
Public Sub RefreshData()
    On Error GoTo errHandler
    oCO.Reload
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.RefreshData"
End Sub

Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    
    If XA(Bookmark, 20) = "CAN" Then
        RowStyle.BackColor = COLOR_CANCELLED
    ElseIf XA(Bookmark, 20) = "FUL" Then
        RowStyle.BackColor = COLOR_FULFILLED
    ElseIf XA(Bookmark, 23) > 0 Then
        RowStyle.BackColor = COLOR_PALEYELLOW
    End If
        
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub
Private Sub G1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnPrevVerCO   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.G1_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), EA_NORERAISE
    HandleError
End Sub
Public Sub BrowseTActions()
    On Error GoTo errHandler
Dim oSQL As New z_SQL
Dim x As New XArrayDB
Dim i, j As Integer
Dim xMLDoc As ujXML
Dim XMLArgs As String

    x.ReDim 1, G1.SelBookmarks.Count, 1, 1
    j = 1
    For i = 1 To G1.SelBookmarks.Count
        x(j, 1) = val(XA(G1.SelBookmarks.Item(i - 1), 14))
        j = j + 1
    Next i
    
    
    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "doc_TrackingActions"
            .chCreate "MessageType"
                .elText = "doc_TrackingActions"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "WORKSTATION"
                .elText = oPC.WorkstationName
            .elCreateSibling "DetailLines", True
            For i = 1 To G1.SelBookmarks.Count
                    .chCreate "ITEM"
                    .chCreate "POLID"
                        .elText = val(XA(G1.SelBookmarks.Item(i - 1), 14))
                    .navUP
                    .navUP
            Next i

         XMLArgs = .docXML
  
    End With
    If Forms(0).frmTRacking Is Nothing Then
        Set Forms(0).frmTRacking = New frmTrackingActions
    End If
    Forms(0).frmTRacking.component "", XMLArgs
    Forms(0).frmTRacking.Show

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.BrowseTActions"
End Sub

Public Sub ShowPreviousOLVersions()
    On Error GoTo errHandler
Dim frm As frmCOLHistory
Dim COLID As Long

    On Error Resume Next
    If XA.Count(1) = 0 Then Exit Sub
    If IsNull(G1.Bookmark) Then Exit Sub
    If Err Then Exit Sub
    
    On Error GoTo errHandler
    COLID = val(XA.Value(G1.Bookmark, 24))
    Set frm = New frmCOLHistory
    frm.component COLID
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.ShowPreviousOLVersions"
End Sub
Public Sub mnuCopyLines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oLine As a_COL
Dim fs As New FileSystemObject

    oPC.PrepareLinesClipboard
    Set rs = oPC.LinesClipboard
    rs.open
    For Each oLine In oCO.COLines
        rs.AddNew
        rs.fields("GUID") = CreateGUID
        rs.fields("PID") = oLine.PID
        rs.fields("Qty") = oLine.Qty
        rs.fields("QtyFirm") = oLine.QtyFirm
        rs.fields("QtySS") = oLine.QtySS
        rs.fields("Price") = oLine.Price
        rs.fields("DISCOUNTRATE") = oLine.Discount
        rs.fields("CODEF") = oLine.CodeF
        rs.fields("EANF") = oLine.EAN
        rs.fields("TITLE") = oLine.Title
        rs.fields("VATRATE") = oLine.VATRate
        rs.fields("REF") = oLine.Ref
        rs.fields("EXTRACHARGEPID") = oLine.ExtraPID
        rs.fields("EXTRACHARGEVALUE") = oLine.ExtraCharge
        rs.Update
    Next
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
        If Err <> 0 Then
            MsgBox "Cannot create folder for Papyrus clipboard", vbInformation + vbOKOnly, "Can't do this"
        End If
    End If
    If fs.FileExists(oPC.LocalFolder & "\TEMP\Clipboard.rs") Then
        fs.DeleteFile oPC.LocalFolder & "\TEMP\Clipboard.rs"
    Else
        If fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
            rs.Save oPC.LocalFolder & "\TEMP\Clipboard.rs"
        End If
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.mnuCopyLines"
End Sub

Public Sub mnuPastelines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim lngCOID As Long
Dim Qty As Long

    If oCO.Status <> stInProcess Then
        MsgBox "You can only add lines to an order that is still in process", vbInformation, "Warning"
        Exit Sub
    End If

    Set rs = oPC.LinesClipboard
    If rs.BOF And rs.eof Then Exit Sub
    If MsgBox("Confirm you are adding " & CStr(rs.RecordCount) & " lines to document " & oCO.DOCCode, vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    rs.MoveFirst
    Do While Not rs.eof
        If FNN(rs.fields("QTYFIRM")) > 0 Then
            Qty = FNN(rs.fields("QTYFIRM"))
        Else
            Qty = FNN(rs.fields("QTY"))
        End If
        oCO.PasteLine FNS(rs.fields("PID")), Qty, FNN(rs.fields("QTYSS")), FNN(rs.fields("PRICE")), FNDBL(rs.fields("DISCOUNTRATE")), _
            FNDBL(rs.fields("VATRATE")), FNS(rs.fields("REF")), FNS(rs.fields("EXTRACHARGEPID")), FNN(rs.fields("EXTRACHARGEVALUE")), DateAdd("ww", 3, Date), _
            FNN(rs.fields("FCPRICE")), FNDBL(rs.fields("FCFACTOR")), FNN(rs.fields("FCID"))
        rs.MoveNext
    Loop
    
    lngCOID = oCO.TRID
    Set oCO = Nothing
    Set oCO = New a_CO
    oCO.Load lngCOID, True
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.mnuPastelines"
End Sub

Private Sub txtTPMemo_Change()
    On Error GoTo errHandler
    txtTPMemo = HandleTextWithBites(txtTPMemo)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.txtTPMemo_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtTPMemo_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    txtTPMemo = HandleTextWithBites(txtTPMemo)

'    If InStr(1, txtTPMemo, Chr(22)) > 0 Then
'        If MsgBox("There are multiple lines in the memo you are saving.", vbExclamation + vbOKCancel, "Warning") = vbCancel Then
'            Cancel = True
'            Exit Sub
'        End If
'    End If
Dim oSM As New z_StockManager
    oSM.SetMemo txtTPMemo, oCO.TRID
    oCO.SetMemo txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.txtTPMemo_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_DragOver(Source As Control, x As Single, _
    Y As Single, State As Integer)
    On Error GoTo errHandler
    Dim picdocument As PictureBox
        ' Optionally move the cursor position so
        ' the user can see where the drop would happen.
        txtTPMemo.SelStart = TextBoxCursorPos(txtTPMemo, x, Y)
        txtTPMemo.SelLength = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.txtTPMemo_DragOver(Source,x,Y,State)", Array(Source, x, Y, State), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_DragDrop(Source As Control, x As Single, _
    Y As Single)
    On Error GoTo errHandler
    txtTPMemo.SelStart = TextBoxCursorPos(txtTPMemo, x, Y)
    txtTPMemo.SelLength = 0
    txtTPMemo.SelText = Source
Dim oSM As New z_StockManager
    oSM.SetMemo txtTPMemo, oCO.TRID
    oCO.SetMemo txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.txtTPMemo_DragDrop(Source,x,Y)", Array(Source, x, Y), EA_NORERAISE
    HandleError
End Sub


Private Sub txtRef_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.SetCORef txtRef, oCO.TRID
    oCO.SetRef txtRef
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.txtRef_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtForAttn_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.SetForAttnCO txtForAttn, oCO.TRID
    oCO.SetForAttn txtForAttn
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.txtForAttn_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdMemo_Click()
    On Error GoTo errHandler
    ShowMemo Not mbShowMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.cmdMemo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub ShowMemo(bON As Boolean)
    On Error GoTo errHandler
        mbShowMemo = bON
        frHeader.Visible = bON
        If bON Then txtRef.SetFocus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.ShowMemo(bOn)", bON
End Sub

Public Sub mnuEmail()
    On Error GoTo errHandler
Dim Res As Boolean
Dim oC As a_CO
Dim strFilename As String
Dim strDestinationEmail As String
Dim strWholeMessage As String
Dim strReference As String

    If oCO.Customer.DispatchMethod = "M" Then
        Screen.MousePointer = vbHourglass
        Set oC = New a_CO
        oC.Load oCO.TRID, True
        Res = oC.ExportToXML(strFilename, enMail, strDestinationEmail)
        Screen.MousePointer = vbDefault
    ElseIf oC.Customer.DispatchMethod = "E" Then
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.mnuEmail"
End Sub

Public Sub mnuOutlook()
    On Error GoTo errHandler
Dim ol As Object
Dim olns As Object
Dim oMI As Object
Dim mfol As Object
Dim fol As Object
Dim Res As Boolean
Dim oC As a_CO
Dim fold As Outlook.Folders
Dim pAttachmentfilename As String
Dim strDestinationEmail As String
Dim strWholeMessage As String
Dim strReference As String
Dim tmp As String
Dim fs As New FileSystemObject
Dim PapyrusDraftsFolder As String
Dim OutlookParentFolder As String

p 1
    If oCO.Customer.DispatchMethod = "M" Then
        Screen.MousePointer = vbHourglass
        Set oC = New a_CO
        oC.Load oCO.TRID, True
        Res = oC.ExportToXML(pAttachmentfilename, enMail, strDestinationEmail, strWholeMessage)
        Screen.MousePointer = vbDefault
    ElseIf oC.Customer.DispatchMethod = "E" Then
    End If
    Set ol = CreateObject("Outlook.Application")
    Set olns = ol.GetNamespace("MAPI")
 p 2
    OutlookParentFolder = GetIniKeyValue(oPC.LocalFolder & "\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERMAIN", "")
    PapyrusDraftsFolder = GetIniKeyValue(oPC.LocalFolder & "\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERSUB", "")
    
p 3
    If PapyrusDraftsFolder > "" Then
        Set fol = olns.Folders(OutlookParentFolder)
p 31
        Set fold = fol.Folders
p 32
        fold.Add PapyrusDraftsFolder
p 33
        Set mfol = fold(PapyrusDraftsFolder)
p 34
    End If
    Set oMI = ol.CreateItem(0)
  p 4
    If pAttachmentfilename > "" Then
        tmp = fs.GetBaseName(pAttachmentfilename)
        strReference = Right(tmp, Len(tmp) - InStr(1, tmp, "_") - 1)
    Else
        strReference = ""
    End If
  p 5
    With oMI
        If oPC.TestMode Then
            .To = oPC.EmailAddressForTesting
        Else
            .To = oCO.BillTOAddress.EMail
        End If
        .Subject = "Sales order: " & strReference
        .BodyFormat = 2   'HTML format
        .Body = ""
        .HTMLBody = FNS(strWholeMessage)
        
        .Attachments.Add (pAttachmentfilename)
        .ReadReceiptRequested = True
        .Close (0)  'save and close
p 6
        If PapyrusDraftsFolder > "" Then .Move mfol
p 7
    End With
    Set oMI = Nothing
    Set olns = Nothing
    Set ol = Nothing
    Set oSM = New z_StockManager
    oSM.LogTransmission oC.TRID, "Sent to Outlook: " & Format(Date, "dd/mm/yyyy")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.mnuOutlook"
End Sub
Public Sub mnuDeliveryDoc()
Dim frmI As frmInvoicePreview
Dim frmG As frmGDNPreview
Dim rs As ADODB.Recordset
Dim fI As frmInvoicePreview
Dim fG As frmGDNPreview
Dim i As Integer
Dim s As String
    oPC.OpenDBSHort
    Screen.MousePointer = vbHourglass
    Set rs = New ADODB.Recordset
    
    rs.open "SELECT * FROM vDeliveredCOLS_c WHERE COL_ID = " & CStr(FNN(XA(Me.G1.Bookmark, 24))), oPC.COShort, adOpenStatic, adLockOptimistic
    If rs.RecordCount = 1 Then
        If rs.fields("DeliveryDocType") = "INV" Then
            Set fI = New frmInvoicePreview
            fI.component FNN(rs.fields("DeliveryDocID"))
            Screen.MousePointer = vbDefault
            fI.Show
        Else
            Set fG = New frmGDNPreview
            fG.component FNN(rs.fields("DeliveryDocID"))
            Screen.MousePointer = vbDefault
            fG.Show
        End If
    Else
        If rs.RecordCount > 1 Then
            For i = 1 To rs.RecordCount
                s = s & IIf(s > "", ",", "") & FNS(rs.fields("DeliveryDoc"))
            Next
            MsgBox "There are multiple documents deliverying this item. Here they are:" & vbCrLf & vbCrLf & s, vbInformation + vbOKOnly, "Results"
        Else
            Screen.MousePointer = vbDefault
            MsgBox "There are no delivery documents found.", vbInformation + vbOKOnly, "Results"
        End If
    oPC.DisconnectDBShort
    End If
End Sub

