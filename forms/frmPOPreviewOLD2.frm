VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmPOPreviewo 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Purchase order preview"
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11340
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmPOPreviewOLD2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Load tracking form for this item"
      Height          =   345
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Close the purchase order"
      Top             =   5880
      Width           =   2580
   End
   Begin VB.CheckBox chklLastActions 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Show last actions"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   3915
      TabIndex        =   29
      Top             =   5535
      Width           =   1665
   End
   Begin VB.CommandButton cmdOrderStatus 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Update tracking notes for this document"
      Height          =   345
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Close the purchase order"
      Top             =   5880
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
      Height          =   615
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPOPreviewOLD2.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Print the invoice"
      Top             =   4830
      Width           =   1000
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
      Height          =   615
      Left            =   1275
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPOPreviewOLD2.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Print the invoice"
      Top             =   4830
      Width           =   1000
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
      Height          =   615
      Left            =   2295
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPOPreviewOLD2.frx":0C9E
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Close the purchase order"
      Top             =   4830
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   270
      TabIndex        =   23
      Top             =   7710
      Width           =   870
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   255
      TabIndex        =   18
      Top             =   5460
      Width           =   1590
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Key:"
         Height          =   300
         Left            =   30
         TabIndex        =   22
         Top             =   15
         Width           =   345
      End
      Begin VB.Label lblCAN 
         Caption         =   "CAN"
         Height          =   300
         Left            =   1140
         TabIndex        =   21
         Top             =   15
         Width           =   345
      End
      Begin VB.Label lblFUL 
         BackColor       =   &H00FEABAD&
         Caption         =   "FUL"
         Height          =   300
         Left            =   765
         TabIndex        =   20
         Top             =   15
         Width           =   345
      End
      Begin VB.Label lblOS 
         Caption         =   "OS"
         Height          =   300
         Left            =   435
         TabIndex        =   19
         Top             =   15
         Width           =   345
      End
   End
   Begin VB.CheckBox optlog 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Show transmission log"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   1950
      TabIndex        =   17
      Top             =   5535
      Width           =   1935
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
      Left            =   315
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   7410
   End
   Begin CoolButtonControl.CoolButton cbTP 
      Height          =   660
      Left            =   4170
      TabIndex        =   15
      Top             =   45
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   1164
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
   Begin VB.TextBox txtDocCOde 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   675
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   135
      Width           =   1545
   End
   Begin VB.TextBox txtDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   135
      Width           =   1545
   End
   Begin VB.TextBox txtIssued 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   300
      Left            =   750
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   435
      Width           =   3195
   End
   Begin VB.TextBox txtTPMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000D&
      Height          =   720
      Left            =   3390
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4755
      Visible         =   0   'False
      Width           =   2820
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
      ForeColor       =   &H00706034&
      Height          =   255
      Left            =   9900
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   465
      Width           =   1155
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
      Height          =   555
      Left            =   3360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5640
      Width           =   2910
   End
   Begin VB.TextBox txtDeliverTo 
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
      Height          =   750
      Left            =   7875
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   15
      Width           =   1980
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
      Left            =   10110
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   960
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   3945
      Left            =   210
      OleObjectBlob   =   "frmPOPreviewOLD2.frx":1028
      TabIndex        =   9
      Top             =   780
      Width           =   10725
   End
   Begin CoolButtonControl.CoolButton cbHeader 
      Height          =   390
      Left            =   45
      TabIndex        =   27
      ToolTipText     =   "Show header information"
      Top             =   75
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   688
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
      Picture         =   "frmPOPreviewOLD2.frx":6DDF
      Style           =   1
   End
   Begin VB.Image imgEmail 
      Height          =   240
      Left            =   4290
      Picture         =   "frmPOPreviewOLD2.frx":6FB9
      Top             =   360
      Width           =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Invoice No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   660
      TabIndex        =   14
      Top             =   285
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   705
      Left            =   540
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   3495
   End
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   1320
      X2              =   2790
      Y1              =   45
      Y2              =   870
   End
   Begin VB.Label txtPhone 
      BackColor       =   &H00E0E0E0&
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
      Height          =   210
      Left            =   4680
      TabIndex        =   6
      Top             =   465
      Width           =   3105
   End
   Begin VB.Label txtSuppname 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4680
      TabIndex        =   5
      Top             =   60
      Width           =   3105
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   4260
      TabIndex        =   4
      Top             =   60
      Width           =   270
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
Attribute VB_Name = "frmPOPreviewo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oPO As ro_PO
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


Public Sub Component(pID As Long)
    On Error GoTo errHandler
    lngID = pID
    Set oPO = New ro_PO
    oPO.Load lngID, True
    Me.Caption = "Purchase order to " & oPO.TPNAME & oPO.StaffNameB & "   " & oPO.DocCode
    LoadControls
    SetMenu
  ' If (oPC.EDIEnabled And oPO.Supplier.GFXNumber > "" And oPO.Supplier.DispatchMethod = "E") Or (oPO.Supplier.DispatchMethod = "M" And oPO.Supplier.OrderToAddress.EMail > "") Then
  '      cmdTransmit.Enabled = True
  '  Else
  '      cmdTransmit.Enabled = False
  '  End If
    lblOS.BackColor = COLOR_PALEYELLOW
    lblFUL.BackColor = COLOR_FULFILLED
    lblCAN.BackColor = COLOR_CANCELLED

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.Component(pID)", pID
End Sub

Private Sub cmdEDI_Click()
    On Error GoTo errHandler
    cmdPrint_Click
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdEDI_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cbHeader_Click()
Dim frm As New frmHeader_PO
Dim strRef As String
Dim strMemo As String

    frm.Component False, oPO.Memo, oPO.TRID
    frm.Show vbModal
    Unload frm
    oPO.Reload
    LoadControls
    
End Sub


Private Sub cmdOrderStatus_Click()
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
    
End Sub



Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.Form_Activate", , EA_NORERAISE
    HandleError
End Sub
Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayout Me.G1, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oPO.statusF = "IN PROCESS")
    Forms(0).mnuCancel.Enabled = (oPO.statusF = "ISSUED") And oPO.CanCancel = True
    Forms(0).mnuCancelLine.Enabled = (oPO.statusF = "ISSUED")
    Forms(0).mnuCancelINactive.Enabled = (oPO.statusF = "ISSUED") And oPO.CanCancel = False
    Forms(0).mnuFulfil.Enabled = (oPO.statusF = "ISSUED") 'And oPO.CanCancel = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    If oPC.EmailPO And (oPO.Status = stISSUED Or oPO.Status = stCOMPLETE) Then
        If (oPC.EDIEnabled And oPO.supplier.GFXNumber > "" And oPO.supplier.DispatchMethod = "E") Or _
        (oPO.supplier.DispatchMethod = "M" And oPO.supplier.OrderToAddress.EMail > "") Then
            Forms(0).mnuEmail.Enabled = Not oPC.UsesOutlookForPOEmail
            Forms(0).mnuOutlook.Enabled = oPC.UsesOutlookForPOEmail
        Else
            Forms(0).mnuEmail.Enabled = False
            Forms(0).mnuOutlook.Enabled = False
        End If
    Else
        Forms(0).mnuEmail.Enabled = False
        Forms(0).mnuOutlook.Enabled = False
    End If

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
        Me.txtDate = .DocDateF
        If DateDiff("d", .DocDate, .issDate) > 1 Then
            Me.txtIssued = "Issued: " & .IssDateF
        Else
            txtIssued = ""
        End If
        Me.txtDocCOde = .DocCode
        Me.txtStatus = .statusF
        CancelLine.Visible = (.Status = stCANCELLED Or .Status = stVOID)
        If oPC.getProperty("CanEditPOs") = "TRUE" Then
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
        Me.imgEmail.Visible = (.supplier.OrderToAddress.EMail > "" And .supplier.DispatchMethod = "M")
        
        Me.txtSuppname = .supplier.NameAndCode(20)
        If Not .supplier.billtoaddress Is Nothing Then
            Me.txtPhone = .supplier.OrderToAddress.PhoneandFax
        End If
        Me.txtTPMemo = FNS(.Memo)
        txtTPMemo.Visible = (txtTPMemo > "")
        txtDeliverTo = .DeliverToAddress
        .DisplayTotals strTotalCaption, strTotalValues, oPO.isFOreignCurrency
        lblTotalCaption.Caption = strTotalCaption
        lblTotalValues.Caption = strTotalValues
        txtCurrency = oPO.CaptureCurrency.Description
    End With
    LoadGrid
    Screen.MousePointer = vbDefault
    mSetfocus cmdClose
EXIT_HANDLER:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.LoadControls"
End Sub



Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo errHandler
   oPO.PrintPO_Display (oPO.isFOreignCurrency)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdPreview_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_PO
Dim res As Boolean
Dim lPO As a_PO
    Screen.MousePointer = vbHourglass
    Set lPO = New a_PO
    lPO.Load oPO.TRID, True
    
    Set frm = New frmPrintingOptions_PO
    frm.ComponentObject lPO, IIf(oPO.supplier.GFXNumber > "" And oPC.EDIEnabled, enEDI, enPrint)
        
    Screen.MousePointer = vbDefault
    frm.Show vbModal
    oPO.Log = lPO.Log
EXIT_HANDLER:
 '   Unload Me
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
    WaitMsg "Loading . . .", True, Me
    lPO.Load oPO.TRID, False
    Set frm = New frmPO
    blnEdit = True
    WaitMsg "", False, Me
    Unload Me
    frm.Component bCancel, lPO ', lngID
    frm.Show

EXIT_HANDLER:
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
    XA.ReDim 1, oPO.POLines.Count, 1, 16
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
                XA(i, 7) = .PriceF(oPO.isFOreignCurrency)
                XA(i, 8) = .DiscountF
                XA(i, 9) = .PLessDiscExtF(oPO.isFOreignCurrency)
                XA(i, 10) = .ETAF
                XA(i, 11) = .lastactionAndDate
                XA(i, 12) = .code
                XA(i, 13) = .pID
                XA(i, 16) = .EAN
        End With
    Next i
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    G1.Array = XA
    G1.ReBind
    
EXIT_HANDLER:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.LoadGrid"
End Sub

Private Sub chklLastActions_Click()
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
    
    
    
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Me.top = 50
        Me.left = 50
        Me.Height = 6500
        Me.Width = 11500
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    G1.Width = Me.Width - (G1.left + 400)
    lngDiff = G1.Height
    G1.Height = Me.Height - (G1.top + 1900)
    lngDiff = G1.Height - lngDiff
    txtTransmission.top = txtTransmission.top + lngDiff
    cmdEdit.top = cmdEdit.top + lngDiff
    cmdPrint.top = cmdPrint.top + lngDiff
    cmdClose.top = cmdClose.top + lngDiff
    Me.cmdOrderStatus.top = cmdOrderStatus.top + lngDiff
    Me.Command2.top = Command2.top + lngDiff
    txtTPMemo.top = txtTPMemo.top + lngDiff
    Frame1.top = Frame1.top + lngDiff
    optlog.top = optlog.top + lngDiff
    chklLastActions.top = chklLastActions.top + lngDiff
    lblTotalCaption.top = lblTotalCaption.top + lngDiff
    lblTotalValues.top = lblTotalValues.top + lngDiff
    Me.Command1.top = Command1.top + lngDiff
    txtCurrencyRates.top = txtCurrencyRates.top + lngDiff
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set oPO = Nothing
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
    frm.Component oPO.supplier
    frm.Show
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
    Clipboard.SetText left(str, ISBNLENGTH)
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
    Clipboard.SetText left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    ErrPreserve
    If Err = 521 Then
        Resume Next
    End If
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
        frmA.Component oP
        frmA.Show
    Else
        Set frm = New frmProductPrev
        frm.Component oP
        frm.Show
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
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
    Clipboard.SetText left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.G1_Click", , EA_NORERAISE
    HandleError
End Sub


Public Sub mnuCancel()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
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
Dim oP As a_Product
    If MsgBox("Do you wish to cancel the selected line?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oPO.POLines.FindLineByID(Val(XA(G1.Bookmark, 14))).CancelLine
    RefreshData
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuCancelLine"
End Sub
Public Sub mnuCancelInactiveLines()
    On Error GoTo errHandler
Dim oP As a_Product
    If MsgBox("Do you wish to cancel all the lines that have not had any products received and mark as fulfilled any lines that are partially received?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oPO.POLines.CancelInactiveLines
    RefreshData
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuCancelInactiveLines"
End Sub
Public Sub mnuFulfilLine()
    On Error GoTo errHandler
Dim oP As a_Product
Dim oPOL As ro_POL
    Set oPOL = oPO.POLines.FindLineByID(Val(XA(G1.Bookmark, 14)))
    If oPOL.Fulfilled <> "OS" Then
        MsgBox "This line is not outstanding and cannot be marked fulfilled.", vbExclamation + vbOKOnly, "Can't do this"
        Exit Sub
    Else
'        If oPOL.QtyReceivedSoFar = 0 Then
'            If MsgBox("This line has not been received at all and should be marked cancelled." & vbCrLf & "Do you want to mark the line cancelled?", vbQuestion + vbYesNo, "Can't do this") = vbNo Then
'               Exit Sub
'            End If
'            Screen.MousePointer = vbHourglass
'            oPOL.CancelLine
'            RefreshData
'            Screen.MousePointer = vbDefault
'        Else
            If MsgBox("Do you wish to mark the selected line as fulfilled?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
            Screen.MousePointer = vbHourglass
            Set oP = New a_Product
            oPO.POLines.FindLineByID(Val(XA(G1.Bookmark, 14))).MarkLineasFulfilled
            RefreshData
            Screen.MousePointer = vbDefault
   '     End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuFulfilLine"
End Sub

Public Sub mnuVoid()
    On Error GoTo errHandler
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oPO.VoidDocument
    RefreshData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuVoid"
End Sub
Public Sub RefreshData()
    On Error GoTo errHandler
    oPO.Reload
    LoadControls
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
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
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    
    G1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 3, 4
            GetRowType = XTYPE_STRING
        Case Else
            GetRowType = XTYPE_NUMBER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.GetRowType(ColIndex)", ColIndex
End Function

Private Sub Option1_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.Option1_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub optlog_Click()
    On Error GoTo errHandler
    If optlog = 1 Then
        txtTransmission = oPO.Log
        txtTransmission.Visible = True
    Else
        txtTransmission.Visible = False
    End If
        
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.optlog_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub G1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuShowOLHistGrp   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.G1_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub

Public Sub ShowPreviousOLVersions()
Dim frm As frmPOLHistory
Dim POLID As Long

    POLID = Val(XA.Value(G1.Bookmark, 14))
    Set frm = New frmPOLHistory
    frm.Component POLID
    frm.Show vbModal
End Sub

Public Sub mnuEmail()
    On Error GoTo errHandler
Dim res As Boolean
Dim lPO As a_PO
Dim strFilename As String
Dim strDestinationEmail As String
Dim strWholeMessage As String
Dim strReference As String

    If oPO.supplier.DispatchMethod = "M" Then
        Screen.MousePointer = vbHourglass
        Set lPO = New a_PO
        lPO.Load oPO.TRID, True
        res = lPO.ExportToXML(False, enMail, strFilename, strDestinationEmail, strWholeMessage)
        Screen.MousePointer = vbDefault
    ElseIf oPO.supplier.DispatchMethod = "E" Then
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.cmdTransmit_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuOutlook()
    On Error GoTo errHandler
Dim ol As Object
Dim olns As Object
Dim oMI As Object
Dim mfol As Object
Dim fol As Outlook.MAPIFolder
Dim res As Boolean
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
p 1
    If oPO.supplier.DispatchMethod = "M" Then
        Screen.MousePointer = vbHourglass
        Set lPO = New a_PO
        lPO.Load oPO.TRID, True
        res = lPO.ExportToXML(False, enMail, pAttachmentfilename, strDestinationEmail, strWholeMessage)
        Screen.MousePointer = vbDefault
    ElseIf oPO.supplier.DispatchMethod = "E" Then
    End If

    Set ol = CreateObject("Outlook.Application")
    Set olns = ol.GetNamespace("MAPI")
 p 2
  '  OutlookParentFolder = oPC.getProperty("OutlookParentOfCustomFolder")
  '  PapyrusDraftsFolder = oPC.getProperty("OutlookCustomFolderForEmails")
    OutlookParentFolder = GetIniKeyValue(oPC.LocalFolder & "\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERMAIN", "")
    PapyrusDraftsFolder = GetIniKeyValue(oPC.LocalFolder & "\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERSUB", "")
    
p 3
    If PapyrusDraftsFolder > "" Then
        On Error Resume Next
        Set fol = olns.Folders(OutlookParentFolder)
p 31
        Set fold = fol.Folders
p 32
        fold.Add PapyrusDraftsFolder
p 33
        Set mfol = fold(PapyrusDraftsFolder)
p 34
        On Error GoTo errHandler
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
            .To = oPO.supplier.OrderToAddress.EMail
        End If
        .Subject = "Purchase order: " & strReference
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
    oSM.LogTransmission lPO.TRID, "Sent to Outlook: " & Format(Date, "dd/mm/yyyy")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.mnuOutlook", , , , "strErrPos", Array(strErrPos, OutlookParentFolder, PapyrusDraftsFolder)
End Sub
Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
Dim oSM As New z_StockManager
    
    ofrm.Component oPO.Memo
    ofrm.Show vbModal
    oSM.setMemo ofrm.Memo, oPO.TRID
    
    txtTPMemo.Visible = (ofrm.Memo > "")
    txtTPMemo = "Note: " & ofrm.Memo
    oSM.setMemo ofrm.Memo, oPO.TRID
    oPO.Memo = ofrm.Memo
    
    Unload ofrm

    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.mnuMemo"
End Sub

Public Sub LoadOrderTracking()
Dim frmP As frmODPO
Dim cODPO As c_POLsOS

    Set frmP = New frmODPO
    Set cODPO = New c_POLsOS
    cODPO.Load , , , (XA(G1.Bookmark, 13))
    frmP.Component cODPO, Now, "", "", ""
    frmP.Show

End Sub
Private Sub Command2_Click()
    LoadOrderTracking
End Sub

Private Sub txtTPMemo_Change()
Dim strArg As String
Dim iStart As Integer
Dim iEnd As Integer
Dim oU As New z_UTIL
Dim strResult As String
Dim f As frmFindTextBite

    iStart = 0
    iEnd = 0
    iStart = InStr(1, txtTPMemo, "?") + 1
    If iStart = 0 Then Exit Sub
    strResult = ""
    iEnd = InStr(iStart, txtTPMemo, "?")
    If iStart > 0 And iEnd > iStart Then
        strArg = Trim(Mid(txtTPMemo, iStart, iEnd - iStart))
        strResult = oU.GetTextBite(strArg)
        If strResult > "" Then
            txtTPMemo = Replace(txtTPMemo, "?" & strArg & "?", strResult)
        End If
    Else
    End If
End Sub

Private Sub txtTPMemo_DblClick()
    If bMemoExpanded Then
        txtTPMemo.Height = txtTPMemo.Height - 800
        txtTPMemo.Width = txtTPMemo.Width - 800
        txtTPMemo.top = txtTPMemo.top + 800
        bMemoExpanded = False
        txtTPMemo.ZOrder 1
    Else
        bMemoExpanded = True
        txtTPMemo.Height = txtTPMemo.Height + 800
        txtTPMemo.Width = txtTPMemo.Width + 800
        txtTPMemo.top = txtTPMemo.top - 800
        txtTPMemo.ZOrder 0
    End If
End Sub

Private Sub txtTPMemo_LostFocus()
    If bMemoExpanded Then
        txtTPMemo.Height = txtTPMemo.Height - 800
        txtTPMemo.Width = txtTPMemo.Width - 800
        txtTPMemo.top = txtTPMemo.top + 800
        bMemoExpanded = False
        txtTPMemo.ZOrder 1
    End If
End Sub

Private Sub txtTPMemo_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single, State As Integer)
    If Data.GetFormat(vbCFText) Then
        Effect = vbDropEffectCopy Or vbDropEffectMove
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub txtTPMemo_Validate(Cancel As Boolean)
    If InStr(1, txtTPMemo, Chr(13)) > 0 Then
        If MsgBox("There are multiple lines in the memo you are saving.", vbExclamation + vbOKCancel, "Warning") = vbCancel Then
            Cancel = True
            Exit Sub
        End If
    End If
Dim oSM As New z_StockManager
    oSM.setMemo txtTPMemo, oPO.TRID
    oPO.Memo = txtTPMemo
End Sub

Private Sub txtTPMemo_DragOver(Source As Control, x As Single, _
    Y As Single, State As Integer)
    Dim picdocument As PictureBox
        ' Optionally move the cursor position so
        ' the user can see where the drop would happen.
        txtTPMemo.SelStart = TextBoxCursorPos(txtTPMemo, x, Y)
        txtTPMemo.SelLength = 0
End Sub

Private Sub txtTPMemo_DragDrop(Source As Control, x As Single, _
    Y As Single)
    txtTPMemo.SelStart = TextBoxCursorPos(txtTPMemo, x, Y)
    txtTPMemo.SelLength = 0
    txtTPMemo.SelText = Source
Dim oSM As New z_StockManager
    oSM.setMemo txtTPMemo, oPO.TRID
    oPO.Memo = txtTPMemo
End Sub

