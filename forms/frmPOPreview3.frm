VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmPOPreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Purchase order preview"
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11220
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmPOPreview3.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDown 
      BackColor       =   &H00C4BCA4&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10785
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4545
      Width           =   330
   End
   Begin VB.CommandButton cmdUP 
      BackColor       =   &H00C4BCA4&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10785
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4215
      Width           =   330
   End
   Begin VB.CommandButton cmdTransmission 
      BackColor       =   &H00FFC0C0&
      Caption         =   "T"
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
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5340
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
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4875
      Width           =   255
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
      Height          =   1185
      Left            =   2895
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2850
      Visible         =   0   'False
      Width           =   5280
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Load tracking form for this item"
      Height          =   345
      Left            =   3135
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Close the purchase order"
      Top             =   5235
      Width           =   2580
   End
   Begin VB.CheckBox chklLastActions 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Show last actions"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   2100
      TabIndex        =   17
      Top             =   5700
      Width           =   1665
   End
   Begin VB.CommandButton cmdOrderStatus 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Update tracking notes for this document"
      Height          =   345
      Left            =   3135
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Close the purchase order"
      Top             =   4875
      Visible         =   0   'False
      Width           =   2970
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
      Left            =   345
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPOPreview3.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   885
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
      Left            =   1245
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPOPreview3.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   885
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
      Left            =   2145
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPOPreview3.frx":0C9E
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Close the purchase order"
      Top             =   4875
      Width           =   885
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   270
      TabIndex        =   12
      Top             =   7710
      Width           =   870
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Height          =   450
      Left            =   345
      TabIndex        =   8
      Top             =   5550
      Width           =   1650
      Begin VB.Label Label2 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Key"
         Height          =   225
         Left            =   120
         TabIndex        =   22
         Top             =   150
         Width           =   345
      End
      Begin VB.Label lblCAN 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "CAN"
         Height          =   285
         Left            =   1245
         TabIndex        =   11
         Top             =   135
         Width           =   345
      End
      Begin VB.Label lblFUL 
         Alignment       =   2  'Center
         BackColor       =   &H00FEABAD&
         Caption         =   "FUL"
         Height          =   285
         Left            =   870
         TabIndex        =   10
         Top             =   135
         Width           =   345
      End
      Begin VB.Label lblOS 
         Alignment       =   2  'Center
         BackColor       =   &H00DBFAFB&
         Caption         =   "OS"
         Height          =   285
         Left            =   540
         TabIndex        =   9
         Top             =   135
         Width           =   345
      End
   End
   Begin VB.TextBox txtTransmission 
      BackColor       =   &H00C6F5F7&
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
      Height          =   1365
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   4680
   End
   Begin CoolButtonControl.CoolButton cbTP 
      Height          =   1440
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   2540
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
   Begin VB.TextBox txtCurrency 
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6495
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1185
      Width           =   1860
   End
   Begin VB.TextBox txtCurrencyRates 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00706034&
      Height          =   300
      Left            =   6390
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1125
      Width           =   4035
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
      Height          =   315
      Left            =   8955
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   1770
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   3345
      Left            =   30
      OleObjectBlob   =   "frmPOPreview3.frx":1028
      TabIndex        =   5
      Top             =   1485
      Width           =   10725
   End
   Begin CoolButtonControl.CoolButton cbDelto 
      Height          =   1425
      Left            =   3855
      TabIndex        =   24
      Top             =   15
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
   Begin CoolButtonControl.CoolButton cmdDispatchMode 
      Height          =   435
      Left            =   6420
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   0
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   767
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
   Begin VB.Label lblDispatchMode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   6525
      TabIndex        =   30
      Top             =   60
      Width           =   1770
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
      Left            =   4170
      TabIndex        =   26
      Top             =   30
      Width           =   1050
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
      Height          =   945
      Left            =   4080
      TabIndex        =   25
      Top             =   300
      Width           =   2055
   End
   Begin VB.Label lblTPName 
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
      Height          =   1065
      Left            =   255
      TabIndex        =   23
      Top             =   195
      Width           =   3240
   End
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   9780
      X2              =   11250
      Y1              =   -120
      Y2              =   705
   End
   Begin VB.Label lblTotalCaption 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
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
      Height          =   1140
      Left            =   6390
      TabIndex        =   2
      Top             =   4935
      Width           =   2610
   End
   Begin VB.Label lblTotalValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
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
      Height          =   1140
      Left            =   9090
      TabIndex        =   1
      Top             =   4920
      Width           =   1845
   End
End
Attribute VB_Name = "frmPOPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim WithEvents oPO As ro_PO
Dim WithEvents oPO As a_PO
Attribute oPO.VB_VarHelpID = -1
Dim dblTotal As Double
Dim XA As XArrayDB
Dim lngID As Long
Dim oSM As z_StockManager
Dim tmpWidth_4 As Long
Dim tmpWidth_5 As Long
Dim tmpWidth_7 As Long
Dim tmpWidth_8 As Long
Dim bMemoExpanded As Boolean
Dim mbShowMemo As Boolean
Dim mbShowTransmission As Boolean
Dim POLSOS As ADODB.Recordset
Dim POLActions As ADODB.Recordset
Dim iDMIdx As Integer
Dim iDelIdx As Integer

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
Public Sub component(PID As Long)
    On Error GoTo errHandler
    lngID = PID
  '  Set oPO = New a_PO
    Set oPO = New a_PO
    oPO.Load lngID, True
    
    Me.Caption = IIf(oPO.OrderType = "NS", "Subscription", "Purchase") & " order  " & oPO.DOCCode & "    " & oPO.DOCDate
    Me.Caption = Me.Caption & "  " & oPO.StaffNameB
    If Not (Day(oPO.DOCDate) = Day(oPO.ProcessingDate) And Month(oPO.DOCDate) = Month(oPO.ProcessingDate) And Year(oPO.DOCDate) = Year(oPO.ProcessingDate)) Then
        Me.Caption = Me.Caption & "  issued at:" & oPO.ProcessingDateFF
    End If
    
    LoadControls
    SetMenu
    lblOS.BackColor = COLOR_PALEYELLOW
    lblFUL.BackColor = COLOR_FULFILLED
    lblCAN.BackColor = COLOR_CANCELLED
    
    mbShowTransmission = False
    mbShowMemo = False
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.Component(pID)", pID
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.Component(pID)", PID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.component(PID)", PID
End Sub

Private Sub cmdEDI_Click()
    On Error GoTo errHandler
    cmdPrint_Click
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdEDI_Click", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdEDI_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdEDI_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cbDelTo_Click()
    On Error GoTo errHandler
Dim i As Long

'    If iDelIdx = 0 Then iDelIdx = setCurrentAddressIndex
'    iDelIdx = iDelIdx + 1
'    If iDelIdx > oPC.Configuration.Stores.Count Then
'        iDelIdx = 1
'    End If
'    oPO.setDelToStoreIDImmediate oPC.Configuration.Stores(iDelIdx).ID
'    Me.lblDelToAddress.Caption = oPC.Configuration.Stores(iDelIdx).DelAddress
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cbDelTo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Function setCurrentAddressIndex() As Integer
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To oPC.Configuration.Stores.Count
        If oPO.DELTOStoreID = oPC.Configuration.Stores.Item(i).ID Then
            setCurrentAddressIndex = i
            Exit For
        End If
    Next
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.setCurrentAddressIndex"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.setCurrentAddressIndex"
End Function




Private Sub cmdDispatchMode_Click()
    On Error GoTo errHandler
Dim i As Integer
START:
    If oPO.Supplier.ID = 0 Then Exit Sub
    i = iDMIdx + 1
    If i > oPO.Supplier.DispatchModes.Count Then
        i = 1
    End If
    lblDispatchMode.Caption = oPO.Supplier.DispatchModes.ItemByOrdinalIndex(i)
    oPO.SetDispatchModeID oPO.Supplier.DispatchModes.Key(oPO.Supplier.DispatchModes.ItemByOrdinalIndex(i))
    iDMIdx = i
Dim oSM As New z_StockManager
    oSM.SetDispatchModeID oPO.DispatchModeID, oPO.TRID
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdDispatchMode_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdDispatchMode_Click", , EA_NORERAISE
    HandleError
End Sub


'Private Sub cbHeader_Click()
'    On Error GoTo errHandler
'Dim frm As New frmHeader_PO
'Dim strRef As String
'Dim strMemo As String
'
'    frm.Component False, oPO.Memo, oPO.TRID
'    frm.Show vbModal
'    Unload frm
'    oPO.Reload
'    LoadControls
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cbHeader_Click"
'End Sub


Private Sub cmdMemo_Click()
    On Error GoTo errHandler
    ShowMemo Not mbShowMemo
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdMemo_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdMemo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub ShowMemo(bON As Boolean)
    On Error GoTo errHandler
        mbShowMemo = bON
        txtTPMemo.Visible = bON
        If bON Then txtTPMemo.SetFocus
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.ShowMemo(bOn)", bOn
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.ShowMemo(bOn)", bON
End Sub

Private Sub cmdOrderStatus_Click()
    On Error GoTo errHandler
Dim frm As New frmOrderStatusReport
Dim strNote As String
Dim strDiarize As String
Dim oSM As New z_StockManager

    frm.Show vbModal
    strNote = frm.Note
    strDiarize = frm.Diarize
    Unload frm
    If strDiarize > "" Or strNote > "" Then
        oSM.ActionODPO_NoteandDiary oPO.POLines, strDiarize, strNote
    Else
        MsgBox "No action taken.", vbInformation + vbOKOnly, "Warning"
    End If
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdOrderStatus_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdOrderStatus_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdTransmission_Click()
    On Error GoTo errHandler
    mbShowTransmission = Not mbShowTransmission
    If mbShowTransmission Then
        txtTransmission = oPO.Log
        txtTransmission.Visible = True
    Else
        txtTransmission.Visible = False
    End If

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdTransmission_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdTransmission_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdUP_Click()
    On Error GoTo errHandler
Dim i As Long
    If G1.Bookmark > 1 Then
        Screen.MousePointer = vbHourglass
        i = G1.Bookmark
        oPO.BeginEdit
        oPO.POLines.swap FNS(XA.Value(G1.Bookmark, 17)), FNS(XA.Value(G1.Bookmark - 1, 17))
        oPO.ApplyEdit
        LoadGrid
        Screen.MousePointer = vbDefault
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdUP_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdUP_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdDown_Click()
    On Error GoTo errHandler
Dim i As Long
    If G1.Bookmark < XA.UpperBound(1) Then
        Screen.MousePointer = vbHourglass
        i = G1.Bookmark
        oPO.BeginEdit
        oPO.POLines.swap FNS(XA.Value(G1.Bookmark, 17)), FNS(XA.Value(G1.Bookmark + 1, 17))
        oPO.ApplyEdit
        LoadGrid
        Screen.MousePointer = vbDefault
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdDown_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdDown_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.Form_Activate", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.Form_Activate"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.Form_Activate", , EA_NORERAISE
    HandleError
End Sub
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.G1, Me.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.mnuSaveLayout"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuSaveLayout"
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.Form_Deactivate", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.Form_Deactivate"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Public Sub mnuCopyLines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oLine As a_POL
Dim fs As New FileSystemObject

    oPC.PrepareLinesClipboard
    Set rs = oPC.LinesClipboard
    rs.open
    For Each oLine In oPO.POLines
        rs.AddNew
        rs.fields("GUID") = CreateGUID
        rs.fields("PID") = oLine.PID
        rs.fields("Qty") = oLine.QtyFirm + oLine.QtySS
        rs.fields("QtyFirm") = oLine.QtyFirm
        rs.fields("QtySS") = oLine.QtySS
        rs.fields("Price") = oLine.Price(False)
        rs.fields("DISCOUNTRATE") = oLine.Discount
        rs.fields("CODEF") = oLine.ProductCodeF
        rs.fields("EANF") = oLine.EAN
        rs.fields("TITLE") = oLine.Title
        rs.fields("VATRATE") = oPC.Configuration.VATRate
        rs.fields("REF") = oLine.Ref
        rs.fields("ETA") = oLine.ETA
'        rs.Fields("EXTRACHARGEPID") = oLine.ExtraPID
'        rs.Fields("EXTRACHARGEVALUE") = oLine.ExtraCharge
        rs.Update
    Next
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
        fs.CreateFolder (oPC.SharedFolderRoot & "\TEMP")
        If Err <> 0 Then
            MsgBox "Cannot create folder for Papyrus clipboard", vbInformation + vbOKOnly, "Can't do this"
        End If
    End If
    If fs.FileExists(oPC.SharedFolderRoot & "\TEMP\Clipboard.rs") Then
        fs.DeleteFile oPC.SharedFolderRoot & "\TEMP\Clipboard.rs"
    Else
        If fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
            rs.Save oPC.SharedFolderRoot & "\TEMP\Clipboard.rs"
        End If
    End If

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.mnuCopyLines"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuCopyLines"
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oPO.StatusF = "IN PROCESS")
    Forms(0).mnuCancel.Enabled = (oPO.StatusF = "ISSUED") And oPO.CanCancel = True
    Forms(0).mnuCancelLine.Enabled = (oPO.StatusF = "ISSUED")
    Forms(0).mnuCancelINactive.Enabled = (oPO.StatusF = "ISSUED") And oPO.CanCancel = False
    Forms(0).mnuFulfil.Enabled = (oPO.StatusF = "ISSUED") 'And oPO.CanCancel = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Forms(0).mnuCopyLines.Enabled = True
    Forms(0).mnuPastelines.Enabled = True
    If (oPO.Status = stISSUED Or oPO.Status = stCOMPLETE) Then
        If Not oPO.Supplier.OrderToAddress Is Nothing Then
            If (oPC.EDIEnabled And oPO.Supplier.GFXNumber > "" And oPO.Supplier.DispatchMethod = "E") Then
                Forms(0).mnuEmail.Enabled = False
                Forms(0).mnuOutlook.Enabled = False
                Forms(0).mnuEDI.Enabled = oPC.EDIEnabled
            Else
                If (oPC.EmailPO And oPO.Supplier.DispatchMethod = "M" And oPO.Supplier.OrderToAddress.EMail > "") Then
                    Forms(0).mnuEmail.Enabled = Not oPC.UsesOutlookForPOEmail
                    Forms(0).mnuOutlook.Enabled = oPC.UsesOutlookForPOEmail
                    Forms(0).mnuEDI.Enabled = False
                Else
                    Forms(0).mnuEmail.Enabled = False
                    Forms(0).mnuOutlook.Enabled = False
                    Forms(0).mnuEDI.Enabled = False
                End If
            End If
        Else
                Forms(0).mnuEmail.Enabled = False
                Forms(0).mnuOutlook.Enabled = False
        End If
    Else
        Forms(0).mnuEmail.Enabled = False
        Forms(0).mnuOutlook.Enabled = False
    End If

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.SetMenu"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.SetMenu"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.SetMenu"
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
    Screen.MousePointer = vbHourglass
    
    With oPO
        Me.txtStatus = .StatusF
        CancelLine.Visible = (.Status = stCANCELLED Or .Status = stVOID)
        If oPC.GetProperty("CanEditPOs") = "TRUE" Then
            If .Status = stInProcess Or .Status = stISSUED Then
                cmdEdit.Enabled = True
            Else
                cmdEdit.Enabled = False
            End If
        Else
            If .Status = stInProcess Then
                cmdEdit.Enabled = True
            Else
                cmdEdit.Enabled = False
            End If
        End If
     '   Me.imgEmail.Visible = (.supplier.OrderToAddress.EMail > "" And .supplier.DispatchMethod = "M")
        
        Me.lblTPName.Caption = .Supplier.NameAndCode(20)
        If Not .Supplier.BillTOAddress Is Nothing Then
            lblTPName.Caption = lblTPName.Caption & vbCrLf & .Supplier.OrderToAddress.Phone & vbCrLf & .Supplier.OrderToAddress.Fax
        End If
        Me.txtTPMemo = FNS(.Memo)
    '    txtTPMemo.Visible = (txtTPMemo > "")
        Me.lblDelToAddress.Caption = .DeliverToAddress
        .DisplayTotals strTotalCaption, strTotalValues, oPO.ISForeignCurrency
        lblTotalCaption.Caption = strTotalCaption
        lblTotalValues.Caption = strTotalValues
        If oPO.CaptureCurrency.ID <> oPC.Configuration.DefaultCurrencyID Then
            txtCurrency = oPO.CaptureCurrency.Description
        End If
    End With
    LoadGrid
    Screen.MousePointer = vbDefault
    mSetfocus cmdClose
EXIT_Handler:
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.LoadControls"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.LoadControls"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.LoadControls"
End Sub



Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdClose_Click", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdClose_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo errHandler
   oPO.PrintPO_Display (oPO.ISForeignCurrency)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdPreview_Click", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdPreview_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdPreview_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_PO
Dim Res As Boolean
Dim lPO As a_PO
Dim oDOC As a_DocumentControl
Dim qtyLinesToPrint As Integer
Dim Dummy As String

    Screen.MousePointer = vbHourglass
    If PrintCommandButtonCTRLDown Then
        PrintCommandButtonCTRLDown = False

        Screen.MousePointer = vbHourglass
        oPO.POLines.SortPOLines enSequence, True

        Set oDOC = oPC.Configuration.DocumentControls.FindDC(oPO.constDOCCODE)
        If oDOC Is Nothing Then
            qtyLinesToPrint = 1
        Else
            qtyLinesToPrint = oPC.Configuration.DocumentControls.FindDC(oPO.constDOCCODE).QtyCopies
        End If

       If oPO.ExportToXML(oPO.ISForeignCurrency, enView, Dummy, , , qtyLinesToPrint, True) = False Then
           Screen.MousePointer = vbDefault
           MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
       End If
       Screen.MousePointer = vbDefault
    Else
        Set lPO = New a_PO
        lPO.Load oPO.TRID, True
        Set frm = New frmPrintingOptions_PO
        frm.ComponentObject lPO, IIf(oPO.Supplier.GFXNumber > "" And oPC.EDIEnabled, enEDI, enPrint)
            
        Screen.MousePointer = vbDefault
        frm.Show vbModal
        oPO.Log = lPO.Log
    End If

EXIT_Handler:

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim blnEdit As Boolean
Dim frm As frmPO
Dim bCancel As Boolean
Dim lPO As New a_PO
Dim strPreviousStatusBarCaption As String
    strPreviousStatusBarCaption = Forms(0).SB1.Panels(2).text
    Forms(0).SB1.Panels(2).text = "LOADING . . ."
    lPO.Load oPO.TRID, False
    Set frm = New frmPO
    blnEdit = True
    frm.component lPO.OrderType, bCancel, lPO
    Unload Me
    
    frm.Show
    Forms(0).SB1.Panels(2).text = strPreviousStatusBarCaption

EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Integer
Dim currDeposit As Currency
Dim currPrice As Currency
Dim dblVAT As Double
Dim strSummaryDescription As String
Dim strSummary As String
Dim lngTotal As Long
Dim lngDepositTotal As Long
Dim tmp
    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, oPO.POLines.Count, 1, 20
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", "frmPOPreview", CStr(i), G1.Columns(i - 1).Width)
    Next
    For i = 1 To oPO.POLines.Count
        With oPO.POLines(i)
                XA(i, 14) = .POLID & "k"
                XA(i, 15) = .Fulfilled
                XA(i, 1) = .CodeF & IIf(.Replacementfor > 0, "*", "")
                XA(i, 2) = .TitleAuthor
                XA(i, 3) = .Ref
                XA(i, 4) = .QtyFirm
                XA(i, 5) = .QtySS
                XA(i, 6) = .QtyReceivedSoFar
                XA(i, 7) = .PriceF(oPO.ISForeignCurrency)
                XA(i, 8) = .DiscountF
                XA(i, 9) = .PLessDiscExtF(oPO.ISForeignCurrency)
                XA(i, 10) = .ETAF
                XA(i, 11) = .LastActionAndDate
                XA(i, 12) = .ProductCode
                XA(i, 13) = .PID
                XA(i, 16) = .EAN
                XA(i, 17) = .Key
                XA(i, 18) = .Sequence
                XA(i, 19) = .COQty
                XA(i, 20) = .TitleAuthor
                If FNN(XA(i, 19)) > 0 And oPC.MarkCustomerOrderLinesOnPOLines > "" Then
                    XA(i, 2) = oPC.MarkCustomerOrderLinesOnPOLines & XA(i, 2)
                End If
        End With
    Next i
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 18, XORDER_ASCEND, XTYPE_STRING
    G1.Array = XA
    G1.ReBind
    
EXIT_Handler:
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.LoadGrid"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.LoadGrid"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.LoadGrid"
End Sub

Private Sub chklLastActions_Click()
    On Error GoTo errHandler
'
    If chklLastActions = 1 Then
        tmpWidth_4 = G1.Columns(3).Width
        tmpWidth_5 = G1.Columns(4).Width
        tmpWidth_7 = G1.Columns(6).Width
        tmpWidth_8 = G1.Columns(7).Width
        
        G1.Columns(3).Width = 100
        G1.Columns(4).Width = 100
        G1.Columns(6).Width = 100
        G1.Columns(7).Width = 100
        G1.Columns(10).Width = 2000
    Else
        G1.Columns(3).Width = tmpWidth_4
        G1.Columns(4).Width = tmpWidth_5
        G1.Columns(6).Width = tmpWidth_7
        G1.Columns(7).Width = tmpWidth_8
        G1.Columns(10).Width = 100
    End If
    
    
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.chklLastActions_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.chklLastActions_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Me.TOP = 50
        Me.Left = 50
        Me.Height = 6500
        Me.Width = 11500
    End If
    Me.lblDispatchMode.Caption = oPO.Supplier.DispatchModes.Item(oPO.DispatchModeID)
    iDMIdx = oPO.Supplier.DispatchModes.FindIndexByKey(oPO.DispatchModeID)
    iDelIdx = oPC.Configuration.Stores.FindStoreIdxByID(oPO.DELTOStoreID)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    G1.Width = NonNegative_Lng(Me.Width - (G1.Left + 400))
    lngDiff = G1.Height
    G1.Height = NonNegative_Lng(Me.Height - (G1.TOP + 1900))
    lngDiff = (G1.Height - lngDiff)
    txtTransmission.TOP = txtTransmission.TOP + lngDiff
    txtTPMemo.TOP = txtTPMemo.TOP + lngDiff
    cmdEdit.TOP = cmdEdit.TOP + lngDiff
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    cmdClose.TOP = cmdClose.TOP + lngDiff
    cmdTransmission.TOP = cmdTransmission.TOP + lngDiff
    cmdMemo.TOP = cmdMemo.TOP + lngDiff
    Me.cmdOrderStatus.TOP = cmdOrderStatus.TOP + lngDiff
    Me.Command2.TOP = Command2.TOP + lngDiff
    Frame1.TOP = Frame1.TOP + lngDiff
    chklLastActions.TOP = chklLastActions.TOP + lngDiff
    lblTotalCaption.TOP = lblTotalCaption.TOP + lngDiff
    lblTotalValues.TOP = lblTotalValues.TOP + lngDiff
    cmdUP.TOP = cmdUP.TOP + lngDiff
    cmdDown.TOP = cmdDown.TOP + lngDiff
    cmdUP.Left = NonNegative_Lng(Me.Width - 400)
    cmdDown.Left = NonNegative_Lng(Me.Width - 400)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set oPO = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.Form_Unload(Cancel)", Cancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Private Sub cbTP_Click()
    On Error GoTo errHandler
Dim frm As frmSupplierPreview
    Set frm = New frmSupplierPreview
    frm.component oPO.Supplier
    frm.Show
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cbTP_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cbTP_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub G1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 16)) > "", FNS(XA.Value(G1.Bookmark, 16)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    On Error Resume Next
   
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.G1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.G1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub G1_SelChange(Cancel As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 16)) > "", FNS(XA.Value(G1.Bookmark, 16)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    On Error Resume Next
    
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.G1_SelChange(Cancel)", Cancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.G1_SelChange(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub G1_DblClick()
    On Error GoTo errHandler
Dim frmA As frmProductPrevAQ
Dim frm As frmProductPrev
Dim oP As a_Product
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 13))
    If str = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oP.Load str, 0 'oDel.DeliveryLines.FindLineByID(val(Me.lvw.SelectedItem.Key)).pID, 0
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
        LogSaveToFile "Access violation in frmPOPreview: G1_DblCLick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmPOPreview: G1_DblCLick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.G1_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub G1_Click()
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 16)) > "", FNS(XA.Value(G1.Bookmark, 16)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    On Error Resume Next
    
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.G1_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.G1_Click", , EA_NORERAISE
    HandleError
End Sub


Public Sub mnuCancel()
    On Error GoTo errHandler
Dim oSM As New z_StockManager

    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_PO_SIGN, , "Cancel this purchase order", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    Else
        If oPO.Status = stInProcess Then
            If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
    End If


    Screen.MousePointer = vbHourglass
    oSM.CancelPO oPO
    RefreshData
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuCancel"
End Sub

Public Sub mnuCancelLine()
    On Error GoTo errHandler
    
    
    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_PO_SIGN, , "Cancel the selected line", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    Else
        If oPO.Status = stInProcess Then
            If MsgBox("Do you want to cancel the selected line?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    'If MsgBox("Do you wish to cancel the selected line?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oPO.POLines.FindLineByID(val(XA(G1.Bookmark, 14))).CancelLine
    RefreshData
    Screen.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.mnuCancelLine"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuCancelLine"
End Sub
Public Sub mnuCancelInactiveLines()
    On Error GoTo errHandler
    If MsgBox("Do you wish to cancel all the lines that have not had any products received and mark as fulfilled any lines that are partially received?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oPO.POLines.CancelInactiveLines
    RefreshData
    Screen.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.mnuCancelInactiveLines"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuCancelInactiveLines"
End Sub
Public Sub mnuFulfilLine()
    On Error GoTo errHandler
Dim oPOL As a_POL
    Set oPOL = oPO.POLines.FindLineByID(val(XA(G1.Bookmark, 14)))
    If oPOL.Fulfilled <> "OS" Then
        MsgBox "This line is not outstanding and cannot be marked fulfilled.", vbExclamation + vbOKOnly, "Can't do this"
        Exit Sub
    Else
        If MsgBox("Do you wish to mark the selected line as fulfilled?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
        Screen.MousePointer = vbHourglass
        oPOL.MarkLineasFulfilled
        RefreshData
        Screen.MousePointer = vbDefault
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.mnuFulfilLine"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuFulfilLine"
End Sub

Public Sub mnuVoid()
    On Error GoTo errHandler
    
    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_PO_SIGN, , "Void this purchase order", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    Else
        If oPO.Status = stInProcess Then
            If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    
    
  '  If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oPO.VoidDocument
    RefreshData
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.mnuVoid"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuVoid"
End Sub
Public Sub RefreshData()
    On Error GoTo errHandler
    oPO.Reload
    LoadControls
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.RefreshData"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.RefreshData"
End Sub

Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If XA(Bookmark, 15) = "CAN" Then
        RowStyle.BackColor = COLOR_CANCELLED
    ElseIf XA(Bookmark, 15) = "FUL" Then
        RowStyle.BackColor = COLOR_FULFILLED
    Else
        RowStyle.BackColor = COLOR_PALEYELLOW
    End If
'    If XA(Bookmark, 19) > 0 Then
'        rowstyle.
'    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
'         RowStyle)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub
'Private Sub G1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)
'If Col <> 1 Then Exit Sub
'If oPC.Configuration.COLAllocationStyle <> "R" Then Exit Sub
'    If FNN(XA(Bookmark, 19)) > 0 Then
'        CellStyle.ForegroundPicture = LoadResPicture(102, vbResBitmap)
'      '  CellStyle.ForegroundPicture = LoadPicture(oPC.SharedFolderRoot & "\Templates\CUST.BMP")
'        CellStyle.ForegroundPicturePosition = dbgFPLeft
'    End If
'End Sub

Private Sub G1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    If ColIndex = 1 Then ColIndex = 19
    
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    
    G1.Refresh
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.G1_HeadClick(ColIndex)", ColIndex
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 3, 20
            GetRowType = XTYPE_STRING
        Case Else
            GetRowType = XTYPE_NUMBER
    End Select
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.GetRowType(ColIndex)", ColIndex
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.GetRowType(ColIndex)", ColIndex
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.GetRowType(ColIndex)", ColIndex
End Function

Private Sub Option1_Click()
    On Error GoTo errHandler

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.Option1_Click", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.Option1_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.Option1_Click", , EA_NORERAISE
    HandleError
End Sub



'Private Sub optlog_Click()
'    On Error GoTo errHandler
'    If optlog = 1 Then
'        txtTransmission = oPO.Log
'        txtTransmission.Visible = True
'    Else
'        txtTransmission.Visible = False
'    End If
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.optlog_Click", , EA_NORERAISE
'    HandleError
'End Sub
Private Sub G1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuShowOLHistGrp   ' Display the File menu as a
                        ' pop-up menu.
   End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.G1_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, x, Y), _
'         EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.G1_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.G1_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), EA_NORERAISE
    HandleError
End Sub
Public Sub ShowPreviousOLVersions()
    On Error GoTo errHandler
Dim frm As frmPOLHistory
Dim POLID As Long

    POLID = val(XA.Value(G1.Bookmark, 14))
    Set frm = New frmPOLHistory
    frm.component POLID
    frm.Show vbModal
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.ShowPreviousOLVersions"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.ShowPreviousOLVersions"
End Sub

Public Sub mnuEmail()
    On Error GoTo errHandler
Dim Res As Boolean
Dim lPO As a_PO
Dim strFilename As String
Dim strDestinationEmail As String
Dim strWholeMessage As String
Dim strReference As String

    If oPO.Supplier.DispatchMethod = "M" Then
        Screen.MousePointer = vbHourglass
        Set lPO = New a_PO
        lPO.Load oPO.TRID, True
        Res = lPO.ExportToXML(oPO.ISForeignCurrency, enMail, strFilename, strDestinationEmail, strWholeMessage)
        Screen.MousePointer = vbDefault
    ElseIf oPO.Supplier.DispatchMethod = "E" Then
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuEmail"
End Sub
Public Sub mnuEDI()
    On Error GoTo errHandler
Dim lPO As a_PO
Dim Res As Boolean
    If oPO.Supplier.DispatchMethod = "E" Then
        Screen.MousePointer = vbHourglass
        Set lPO = New a_PO
        lPO.Load oPO.TRID, True
        lPO.CalculateTotals
        If lPO.Supplier.EDIType = "E9" Then
            Res = lPO.ExportToXML_MEC(False, enEDI, "ABC")
        Else
            If lPO.Supplier.EDIType = "SA" Then
                Res = lPO.GenerateSAANAMsg
            End If
        End If
        Screen.MousePointer = vbDefault
    End If

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.mnuEDI"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuEDI"
End Sub

Public Sub mnuOutlook()
    On Error GoTo errHandler
Dim ol As Object
Dim olns As Object
Dim oMI As Object
Dim mfol As Object
Dim fol As Outlook.MAPIFolder
Dim Res As Boolean
Dim lPO As a_PO
Dim fold As Outlook.Folders
Dim pAttachmentfilename As String
Dim strDestinationEmail As String
Dim strWholeMessage As String
Dim strReference As String
Dim tmp As String
Dim fs As New FileSystemObject
Dim PapyrusDraftsFolder As String
Dim OutlookParentFolder As String
    If oPO.Supplier.DispatchMethod = "M" Then
        Screen.MousePointer = vbHourglass
        Set lPO = New a_PO
        lPO.Load oPO.TRID, True
        Res = lPO.ExportToXML(oPO.ISForeignCurrency, enMail, pAttachmentfilename, strDestinationEmail, strWholeMessage)
        Screen.MousePointer = vbDefault
    ElseIf oPO.Supplier.DispatchMethod = "E" Then
    End If

    Set ol = CreateObject("Outlook.Application")
    Set olns = ol.GetNamespace("MAPI")
    OutlookParentFolder = GetIniKeyValue(oPC.LocalFolder & "\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERMAIN", "")
    PapyrusDraftsFolder = GetIniKeyValue(oPC.LocalFolder & "\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERSUB", "")
    
    If PapyrusDraftsFolder > "" Then
        Set fol = olns.Folders(OutlookParentFolder)
        Set fold = fol.Folders
        On Error Resume Next
        fold.Add PapyrusDraftsFolder
        On Error GoTo errHandler
        Set mfol = fold(PapyrusDraftsFolder)
    End If
    
    Set oMI = ol.CreateItem(0)
    If pAttachmentfilename > "" Then
        tmp = fs.GetBaseName(pAttachmentfilename)
        strReference = Right(tmp, Len(tmp) - InStr(1, tmp, "_"))
    Else
        strReference = ""
    End If
    
    With oMI
        If oPC.TestMode Then
            .To = oPC.EmailAddressForTesting
        Else
            .To = oPO.Supplier.OrderToAddress.EMail
        End If
        .Subject = "Purchase order: " & strReference
        If oPC.EmailPOShowHTML Then
            .BodyFormat = 2   'HTML format
            .Body = ""
            .HTMLBody = FNS(strWholeMessage)
        Else
            .BodyFormat = 3
            .Body = "Please open the attached order document." & vbCrLf
        End If
        
        .Attachments.Add (pAttachmentfilename)
        .ReadReceiptRequested = True
        .Close (0)  'save and close
        If PapyrusDraftsFolder > "" Then .Move mfol
    End With
    Set oMI = Nothing
    Set olns = Nothing
    Set ol = Nothing
    Set oSM = New z_StockManager
    oSM.LogTransmission lPO.TRID, "Sent to Outlook: " & Format(Date, "dd/mm/yyyy")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuOutlook", , , , "strErrPos", Array(strErrPos)
End Sub
Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
Dim oSM As New z_StockManager
    
    ofrm.component oPO.Memo
    ofrm.Show vbModal
    oSM.SetMemo ofrm.Memo, oPO.TRID
    
    txtTPMemo.Visible = (ofrm.Memo > "")
    txtTPMemo = "Note: " & ofrm.Memo
    oSM.SetMemo ofrm.Memo, oPO.TRID
    oPO.Memo = ofrm.Memo
    
    Unload ofrm

    Set ofrm = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.mnuMemo"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.mnuMemo"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuMemo"
End Sub

Public Sub LoadOrderTracking()
    On Error GoTo errHandler
Dim frmP As frmODPO
Dim cODPO As c_POLSOS2
Dim x As New XArrayDB
Dim i, j As Integer
Dim XMLArgs As String
Dim xMLDoc As ujXML

    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "doc_POL_LINES"
            .chCreate "MessageType"
                .elText = "doc_POL_LINES"
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
    Set frmP = New frmODPO
    Set cODPO = New c_POLSOS2
    cODPO.LoadRecordsetsPIDSet POLSOS, POLActions, XMLArgs
    frmP.component POLSOS, POLActions
    frmP.Show

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.LoadOrderTracking"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.LoadOrderTracking"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.LoadOrderTracking"
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
    Forms(0).frmTRacking.component XMLArgs, ""
    Forms(0).frmTRacking.Show
    Forms(0).frmTRacking.ZOrder 0
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.BrowseTActions"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.BrowseTActions"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.BrowseTActions"
End Sub
Public Sub mnuPreDelAdv()
    On Error GoTo errHandler
Dim frm As New frmPreDeliveryAdvice
Dim x As New XArrayDB
Dim i, j As Integer
Dim xMLDoc As ujXML
Dim XMLArgs As String
    
    On Error Resume Next
    If XA.Count(1) = 0 Then Exit Sub
    If IsNull(G1.Bookmark) Then Exit Sub
    If Err Then Exit Sub
    
    If G1.SelBookmarks.Count < 1 Then
        MsgBox "You have not selected a line.", vbInformation + vbOKOnly, "Can't do this"
        Exit Sub
    End If
    
    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "doc_PRE_DEL_ADVICE"
            .chCreate "MessageType"
                .elText = "PRE_DEL_ADVICE"
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
    
    
    frm.component XMLArgs, "", "P", ""
    frm.Show vbModal
    Unload frm

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.mnuPreDelAdv"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.mnuPreDelAdv"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuPreDelAdv"
End Sub

Private Sub Command2_Click()
    On Error GoTo errHandler
    If G1.SelBookmarks.Count = 0 Then
        MsgBox "Select one or more rows before choosing this option.", vbInformation + vbOKOnly, "Can't do this"
        Exit Sub
    End If
    LoadOrderTracking
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.Command2_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.Command2_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_Change()
    On Error GoTo errHandler
    txtTPMemo = HandleTextWithBites(txtTPMemo)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.txtTPMemo_Change"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.txtTPMemo_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_DblClick()
    On Error GoTo errHandler
    If bMemoExpanded Then
        txtTPMemo.Height = txtTPMemo.Height - 800
        txtTPMemo.Width = txtTPMemo.Width - 800
        txtTPMemo.TOP = txtTPMemo.TOP + 800
        bMemoExpanded = False
        txtTPMemo.ZOrder 1
    Else
        bMemoExpanded = True
        txtTPMemo.Height = txtTPMemo.Height + 800
        txtTPMemo.Width = txtTPMemo.Width + 800
        txtTPMemo.TOP = txtTPMemo.TOP - 800
        txtTPMemo.ZOrder 0
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.txtTPMemo_DblClick"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.txtTPMemo_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_LostFocus()
    On Error GoTo errHandler
    If bMemoExpanded Then
        txtTPMemo.Height = txtTPMemo.Height - 800
        txtTPMemo.Width = txtTPMemo.Width - 800
        txtTPMemo.TOP = txtTPMemo.TOP + 800
        bMemoExpanded = False
        txtTPMemo.ZOrder 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.txtTPMemo_LostFocus"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.txtTPMemo_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single, State As Integer)
    On Error GoTo errHandler
    If Data.GetFormat(vbCFText) Then
        Effect = vbDropEffectCopy Or vbDropEffectMove
    Else
        Effect = vbDropEffectNone
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.txtTPMemo_OLEDragOver(Data,Effect,Button,Shift,x,Y,State)", Array(Data, _
'         Effect, Button, Shift, x, Y, State)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.txtTPMemo_OLEDragOver(Data,Effect,Button,Shift,x,Y,State)", Array(Data, _
         Effect, Button, Shift, x, Y, State), EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    txtTPMemo = HandleTextWithBites(txtTPMemo)
    
'    If InStr(1, txtTPMemo, Chr(13)) > 0 Then
'        If MsgBox("There are multiple lines in the memo you are saving.", vbExclamation + vbOKCancel, "Warning") = vbCancel Then
'            Cancel = True
'            Exit Sub
'        End If
'    End If
Dim oSM As New z_StockManager
    oSM.SetMemo txtTPMemo, oPO.TRID
    oPO.Memo = txtTPMemo
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.txtTPMemo_Validate(Cancel)", Cancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.txtTPMemo_Validate(Cancel)", Cancel, EA_NORERAISE
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.txtTPMemo_DragOver(Source,x,Y,State)", Array(Source, x, Y, State)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.txtTPMemo_DragOver(Source,x,Y,State)", Array(Source, x, Y, State), _
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
    oSM.SetMemo txtTPMemo, oPO.TRID
    oPO.Memo = txtTPMemo
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.txtTPMemo_DragDrop(Source,x,Y)", Array(Source, x, Y)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.txtTPMemo_DragDrop(Source,x,Y)", Array(Source, x, Y), EA_NORERAISE
    HandleError
End Sub







