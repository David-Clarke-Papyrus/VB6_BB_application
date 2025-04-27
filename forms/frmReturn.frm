VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmReturn 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Return"
   ClientHeight    =   6135
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   10950
   ControlBox      =   0   'False
   Icon            =   "frmReturn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   10950
   Begin VB.CommandButton cmdIssue 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Issu&e"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9705
      Picture         =   "frmReturn.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5340
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7470
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmReturn.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5340
      Width           =   1110
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Sa&ve"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8580
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmReturn.frx":0C9E
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5340
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin MSComctlLib.ListView lvwLines 
      Height          =   2760
      Left            =   105
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   420
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4868
      SortKey         =   9
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Doc ref."
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Invoice ref."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Price"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Discount"
         Object.Width           =   1835
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Total"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Key"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.TextBox txtError 
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
      ForeColor       =   &H000000C0&
      Height          =   780
      Left            =   2100
      MultiLine       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5325
      Width           =   3540
   End
   Begin VB.CommandButton cmdNewRows 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5340
      Width           =   870
   End
   Begin VB.TextBox txtRunningTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   250
      Left            =   9435
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3255
      Width           =   1200
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00D3D3CB&
      Height          =   1725
      Left            =   90
      TabIndex        =   12
      Top             =   3495
      Width           =   10710
      Begin EXCOMBOBOXLibCtl.ComboBox cboMatch 
         Height          =   315
         Left            =   1800
         OleObjectBlob   =   "frmReturn.frx":1028
         TabIndex        =   33
         Top             =   480
         Width           =   6135
      End
      Begin VB.TextBox txtSuppRef 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   4005
         TabIndex        =   7
         Top             =   885
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   2640
         TabIndex        =   6
         Top             =   885
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CommandButton cmdCancelMatch 
         BackColor       =   &H00C4BCA4&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   525
         Width           =   255
      End
      Begin VB.CommandButton cmdEnter 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Post"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9570
         MaskColor       =   &H00C4BCA4&
         Picture         =   "frmReturn.frx":23D2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   210
         Width           =   1000
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   585
         TabIndex        =   5
         Top             =   885
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   8595
         TabIndex        =   4
         Top             =   435
         Width           =   735
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   6390
         TabIndex        =   8
         Top             =   885
         Width           =   4170
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
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
         Height          =   330
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1335
         Width           =   5220
      End
      Begin VB.TextBox txtCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   75
         TabIndex        =   2
         Top             =   450
         Width           =   1650
      End
      Begin VB.Label lblReference 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ref."
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   3540
         TabIndex        =   10
         Top             =   915
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disc."
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2115
         TabIndex        =   3
         Top             =   900
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblPrice 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   150
         TabIndex        =   32
         Top             =   900
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblWants 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Supplier invoice details"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1740
         TabIndex        =   27
         Top             =   195
         Width           =   3120
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         Caption         =   "Qty"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   8640
         TabIndex        =   18
         Top             =   180
         Width           =   525
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Note:"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   5850
         TabIndex        =   17
         Top             =   930
         Width           =   570
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   660
         TabIndex        =   14
         Top             =   195
         Width           =   510
      End
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
      Height          =   360
      Left            =   6060
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "IN PROCESS"
      Top             =   5535
      Width           =   1260
   End
   Begin CoolButtonControl.CoolButton cbTP 
      Height          =   375
      Left            =   4260
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   661
      BackColor       =   14737632
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
   Begin CoolButtonControl.CoolButton cbDelTo 
      Height          =   375
      Left            =   9315
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      BackStyle       =   0
   End
   Begin VB.TextBox txtCurrencyRates 
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   250
      Left            =   90
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3075
      Width           =   7230
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   4725
      Picture         =   "frmReturn.frx":275C
      Stretch         =   -1  'True
      Top             =   60
      Width           =   360
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Height          =   240
      Left            =   120
      TabIndex        =   24
      Top             =   60
      Width           =   225
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
      Left            =   495
      TabIndex        =   23
      Top             =   45
      Width           =   4020
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
      Left            =   5280
      TabIndex        =   22
      Top             =   45
      Width           =   1530
   End
   Begin VB.Label txtFax 
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
      Left            =   7605
      TabIndex        =   21
      Top             =   45
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
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
      Height          =   240
      Left            =   7020
      TabIndex        =   20
      Top             =   30
      Width           =   390
   End
End
Attribute VB_Name = "frmReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents ooR As a_R
Attribute ooR.VB_VarHelpID = -1
Dim WithEvents oRLine As a_RL
Attribute oRLine.VB_VarHelpID = -1
Dim oTP As a_Supplier
Dim oProd As a_Product
Dim bValidRET As Boolean
Dim bValidRETLine As Boolean
Dim tlSupplier As z_TextList
Dim lngCurrentExtension As Long
Dim lngCurrentTotal As Long
Dim lngCurrentDepositTotal As Long
Dim lngCurrentVATTotal As Long
Dim lngCurrencyID As Long
Dim lngSelectedRowIndex As String
Dim lngEditingIdx As String
Dim vMode As EnumMode  ' 1:TPExists,Adding row;  2:TPExists, not adding row;  3 TPAbsent,not adding row
Dim bFrameEnabled As Boolean
Dim lngStockBal As Long
Dim curDeposit As Currency
Dim curTotal As Double
Dim curPrice As Currency
Dim dblQty As Double
Dim lngCompanyID As Long
Dim currPrice As Currency
Dim blnReadOnly As Boolean
Dim flgLoading As Boolean
Dim strRETErrMsg As String
Dim strRETLErrMsg As String
Dim cDELLSPerPIDTP As c_DELLSPerPIDTP
Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
    ofrm.component ooR.Memo
    ofrm.Show vbModal
    ooR.SetMemo ofrm.Memo
    Unload ofrm
    Set ofrm = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.mnuMemo"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.mnuMemo"
End Sub




Private Sub cmdCancelMatch_Click()
    On Error GoTo errHandler
    txtPrice.Visible = True
    txtDiscount.Visible = True
    txtSuppRef.Visible = True
    lblReference.Visible = True
    lblPrice.Visible = True
    lblDiscount.Visible = True
    mSetfocus txtQty
    cboMatch.Items.RemoveAllItems
    oRLine.AllowNoMatchingDelivery
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.cmdCancelMatch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.Form_Activate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.Form_Deactivate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (ooR.StatusF = "IN PROCESS" And ooR.IsNew = False)
    Forms(0).mnuCancel.Enabled = (ooR.StatusF = "ISSUED") ' And oDOC.CanCancel = True
    Forms(0).mnuDelLine.Enabled = True
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Forms(0).mnuCopyLines.Enabled = False
    Forms(0).mnuPastelines.Enabled = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.SetMenu"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.SetMenu"
End Sub



Public Sub component(pCancel As Boolean, Optional pR As a_R, Optional PID As Long)
    On Error GoTo errHandler
Dim ar() As String
    pCancel = False
    flgLoading = True
    SetupcboMatch
    If pR Is Nothing Then
        Set ooR = New a_R
        ooR.BeginEdit
        ooR.SetStatus stInProcess
        lvwLines.Enabled = False
        If PID > 0 Then
            ooR.LoadSupplierFromID PID
            ooR.TPID = ooR.Supplier.ID
            ooR.TPNAME = ooR.Supplier.NameAndCode(35)
            ooR.CurrencyID = ooR.Supplier.DefaultCurrency.ID
        End If
        lvwLines.Height = 2200
        cmdNewRows.Caption = "&Stop"
        vMode = enAddingRow
        SetEditFrameEnabled True, vMode
        ooR.GetStatus
        mSetfocus txtCode
        Set oRLine = ooR.RLines.Add
    Else
        Set ooR = pR
        ooR.BeginEdit
        LoadListView
        cmdSave.Enabled = False
        cmdIssue.Enabled = False
        cmdCancel.Caption = "&Close"
        cmdNewRows.Enabled = True
        lvwLines.Enabled = True
        lvwLines.Height = 4850
        vMode = enNotEditing
        ooR.GetStatus
        SetEditFrameEnabled False, enNotEditing
        Me.Caption = "RETURN " & ooR.DOCCode & "   Approval: " & ooR.ApprovalRef & "   Approval termination date: " & ooR.ApprovalTermDateF
    End If
        
        LoadSupplier
    ooR.GetStatus
    SetMenu
    Select Case ooR.Status
    Case stInProcess
        cmdIssue.Caption = "1. Request"
    Case stISSUED
        cmdIssue.Caption = "2. Return"
    End Select
        
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.Component(pCancel,pR,pID)", Array(pCancel, pR, PID)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.component(pCancel,pR,PID)", Array(pCancel, pR, PID)
End Sub

Private Sub cmdSupplier_Click()
    On Error GoTo errHandler
Dim frm As frmSupplierPreview
    If ooR.Supplier.Name = "" Then Exit Sub
    Set frm = New frmSupplierPreview
    frm.component ooR.Supplier
    frm.Show
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.cmdSupplier_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.cmdSupplier_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cbTP_Click()
    On Error GoTo errHandler
Dim frm As New frmSupplierPreview
    If ooR.Supplier.ID > 0 Then
        frm.component ooR.Supplier
        frm.Show
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.cbTP_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.cbTP_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set oTP = Nothing
    Set ooR = Nothing
    Set tlSupplier = Nothing
    Set oRLine = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.Form_Terminate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwLines_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
    Cancel = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.lvwLines_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
'         EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.lvwLines_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub
Private Sub lvwLines_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.lvwLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.lvwLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Public Sub mnuDelLine()
    On Error GoTo errHandler
    RemoveDetailLine
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.mnuDelLine"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.mnuDelLine"
End Sub
Private Sub mnuPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_CN
    Set frm = New frmPrintingOptions_CN
    frm.Show vbModal
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.mnuPrint_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.mnuPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub ooR_Valid(pMsg As String)
    On Error GoTo errHandler
    bValidRET = (pMsg = "")
    cmdIssue.Enabled = (bValidRET And ooR.RLines.Count > 0 And vMode = enNotEditing)
    cmdSave.Enabled = (bValidRET And vMode = enNotEditing)
    strRETErrMsg = pMsg
    If vMode = enNotEditing Then
        txtError = strRETErrMsg
    Else
        txtError = strRETLErrMsg
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.ooR_Valid(pMsg)", pMsg, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.ooR_Valid(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub
Private Sub oRLine_Valid(msg As String)
    On Error GoTo errHandler
    cmdEnter.Enabled = (msg = "")
    strRETLErrMsg = msg
    If vMode = enNotEditing Then
        txtError = strRETErrMsg
    Else
        txtError = strRETLErrMsg
    
    End If
        LogSaveToFile ("oRLine_Valid(msg)" & msg)

    Me.Refresh

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.oRLine_Valid(msg)", msg, EA_NORERAISE
    HandleError
End Sub

Private Sub ooR_TotalChange(strtotal As String, strTotalForeign As String)
    On Error GoTo errHandler
    flgLoading = True
    If ooR.CaptureCurrency Is oPC.Configuration.DefaultCurrency Then
        Me.txtRunningTotal = strtotal
    Else
        Me.txtRunningTotal = strTotalForeign
        txtCurrencyRates = ooR.CurrencyConversionAsText & "     Value is : " & strtotal
    End If
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.ooR_TotalChange(strtotal,strTotalForeign)", Array(strtotal, strTotalForeign), _
'         EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.ooR_TotalChange(strtotal,strTotalForeign)", Array(strtotal, strTotalForeign), _
         EA_NORERAISE
    HandleError
End Sub
Private Sub ooR_Reloadlist()
    On Error GoTo errHandler
    LoadListView
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.ooR_Reloadlist", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.ooR_Reloadlist", , EA_NORERAISE
    HandleError
End Sub
Private Sub ooR_Dirty(pVal As Boolean)
    On Error GoTo errHandler
If pVal = True Then
        Me.cmdSave.Enabled = (True And Not bFrameEnabled)
        Me.cmdCancel.Caption = "&Cancel"
    Else
        Me.cmdSave.Enabled = False
        Me.cmdCancel.Caption = "&Close"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.ooR_Dirty(pVal)", pVal, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.ooR_Dirty(pVal)", pVal, EA_NORERAISE
    HandleError
End Sub



Sub vCanAdd_NobrokenRules()
    On Error GoTo errHandler
    Me.cmdNewRows.Enabled = True
    mSetfocus cmdNewRows
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.vCanAdd_NobrokenRules", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.vCanAdd_NobrokenRules", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Left = 10
        TOP = 10
        Width = 11100
        Height = 6700
    End If
  '  SetLvw
    vMode = enNotEditing
    lvwLines.Height = 4850
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set oProd = Nothing
    UnsetMenu
    If ooR.IsEditing Then ooR.CancelEdit
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Public Function SetSupplier(pTPID As Long) As Boolean
    On Error GoTo errHandler
Dim bSuccess As Boolean
    bSuccess = ooR.Supplier.Load(pTPID)
    SetSupplier = bSuccess
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.SetSupplier(pTPID)", pTPID
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.SetSupplier(pTPID)", pTPID
End Function

Private Sub SetEditFrameEnabled(pYesNo As Boolean, eMode As EnumMode)
    On Error GoTo errHandler
Dim lngColour As Long
    'A is adding, E is editing
    bFrameEnabled = pYesNo   'shared for use in all the form
    If (eMode = enAddingRow Or eMode = enNotEditing) And pYesNo Then
        Me.txtCode.Enabled = True
    Else
        Me.txtCode.Enabled = False
    End If
    txtNote.Enabled = pYesNo
    txtCurrencyRates.Enabled = pYesNo
    txtTitle.Enabled = pYesNo
    txtQty.Enabled = pYesNo
    cboMatch.Enabled = pYesNo
    
    Me.cmdEnter.Enabled = Not pYesNo
    Me.cmdCancel.Enabled = Not pYesNo
    Me.cmdIssue.Enabled = (Not pYesNo) And bValidRET
    Me.cmdSave.Enabled = (Not pYesNo) And bValidRET And ooR.IsDirty
    
    
    If pYesNo Then
        lngColour = &HFFFFFF
    Else
        lngColour = 14416635
    End If
    
    Me.txtCode.BackColor = lngColour
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.SetEditFrameEnabled(pYesNo,eMode)", Array(pYesNo, eMode)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.SetEditFrameEnabled(pYesNo,eMode)", Array(pYesNo, eMode)
End Sub

Private Sub cmdEnter_Click()
    On Error GoTo errHandler
Dim currDeposit As Currency
Dim blnResult As Boolean
Dim strCurrFormat As String
Dim curTotalDeposit As Currency
    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Papyrus Ordering Information"
        If txtCode.Enabled = True Then mSetfocus txtCode
        Exit Sub
    End If
    oRLine.ApplyEdit
    oRLine.BeginEdit
    txtPrice.Visible = False
    txtDiscount.Visible = False
    txtSuppRef.Visible = False
    txtPrice = ""
    txtDiscount = ""
    txtSuppRef = ""
    lblReference.Visible = False
    lblPrice.Visible = False
    lblDiscount.Visible = False

    If vMode = enAddingRow Then
        lvwLines.ListItems.Add 1, oRLine.Key
        LoadListViewLine oRLine.Key, Me.lvwLines.ListItems(1)
        Set oRLine = Nothing
        Set oRLine = ooR.RLines.Add
        lvwLines.Height = 2760
        oRLine.TRID = ooR.TRID
        ClearLineControls
        mSetfocus txtCode
    ElseIf vMode = eneditingrow Then
        LoadListViewLine lngSelectedRowIndex, Me.lvwLines.ListItems(lngSelectedRowIndex)
        ClearLineControls
        vMode = enNotEditing
        SetEditFrameEnabled False, vMode
        lvwLines.Height = 4850
        cmdNewRows.Caption = "&Add"
        fr1.ZOrder 1
        txtCurrencyRates.ZOrder 1
    End If
    ooR.CalculateTotals
    LoadListView
    ooR.GetStatus
    txtRunningTotal = ooR.TotalLessDiscExtF(ooR.ISForeignCurrency)
    lvwLines.Enabled = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.cmdEnter_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.cmdEnter_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdNewRows_Click()
    On Error GoTo errHandler
Dim lr As Long
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
'    If Not oProd Is Nothing Then
'        If oProd.IsEditing Then oProd.CancelEdit
'    End If
    If vMode = eneditingrow Then       'We have finished editing a row
        cmdNewRows.Caption = "&Add"
        vMode = enNotEditing
        lvwLines.Enabled = True
        lvwLines.Height = 4850
        fr1.ZOrder 1
        txtCurrencyRates.ZOrder 1
        SetEditFrameEnabled False, vMode
    ElseIf vMode = enAddingRow Then    'we are stopping adding rows
        cmdNewRows.Caption = "&Add"
        vMode = enNotEditing 'enEditingRow
        SetEditFrameEnabled False, vMode
        lvwLines.Enabled = True
        lvwLines.Height = 4850
        fr1.ZOrder 1
        txtCurrencyRates.ZOrder 1
        ooR.GetStatus
    ElseIf vMode = enNotEditing Then  'we are starting to add rows
        cmdNewRows.Caption = "&Stop"
        vMode = enAddingRow
        SetEditFrameEnabled True, vMode
        lvwLines.Enabled = False
        lvwLines.Height = 2760
        Set oRLine = ooR.RLines.Add
        oRLine.TRID = ooR.TRID
        mSetfocus txtCode
    End If
    ClearLineControls
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.cmdNewRows_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.cmdNewRows_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadListView()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Long
    lvwLines.ListItems.Clear
    For i = 1 To ooR.RLines.Count
        Set lstItem = lvwLines.ListItems.Add
        With ooR.RLines.Item(i)
            lstItem.text = .CodeF
            If lstItem.Key = "" Then lstItem.Key = .Key
            lstItem.SubItems(9) = Format(.Key, "@@@@@@@@@@")
            lstItem.SubItems(1) = .Title
            lstItem.SubItems(2) = .DocRef
            lstItem.SubItems(3) = .QtyRequested '& "," & .QtyApproved & "," & .QtyReturned
            lstItem.SubItems(5) = .DOCDate
            lstItem.SubItems(4) = .SINVRef
            lstItem.SubItems(7) = .DiscountF
            If oPC.Configuration.DefaultCurrency Is ooR.CaptureCurrency Then
                lstItem.SubItems(6) = .PriceF(False)
                lstItem.SubItems(8) = .PLessDiscExtF(False)
            Else
                lstItem.SubItems(6) = .PriceF(True)
                lstItem.SubItems(8) = .PLessDiscExtF(True)
            End If
        End With

    Next i
EXIT_Handler:
    Set lstItem = Nothing
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.LoadListView"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.LoadListView"
End Sub
Private Sub LoadListViewLine(i As String, lstItem As ListItem)
    On Error GoTo errHandler
Dim currPrice As Currency
    With oRLine
        lstItem.text = .CodeF
        If lstItem.Key = "" Then lstItem.Key = i
        lstItem.SubItems(1) = .Title
        lstItem.SubItems(2) = .DocRef
        lstItem.SubItems(3) = .QtyRequested '& "," & .QtyApproved & "," & .QtyReturned
        lstItem.SubItems(4) = .SINVRef
        lstItem.SubItems(5) = .DOCDate
        lstItem.SubItems(7) = .DiscountF
        lstItem.SubItems(9) = Format(.Key, "@@@@@@@@@@")
        If oPC.Configuration.DefaultCurrency Is ooR.CaptureCurrency Then
            lstItem.SubItems(6) = .PriceF(False)
            lstItem.SubItems(8) = .PLessDiscExtF(False)
        Else
            lstItem.SubItems(6) = .PriceF(True)
            lstItem.SubItems(8) = .PLessDiscExtF(True)
        End If
    End With
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.LoadListViewLine(i,lstItem)", Array(i, lstItem)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.LoadListViewLine(i,lstItem)", Array(i, lstItem)
End Sub
Private Sub lvwLines_DblClick()
    On Error GoTo errHandler
'This must load the editing line with the current line's data
    If lvwLines.ListItems.Count = 0 Then Exit Sub
    If lvwLines.SelectedItem.Index < 1 Then Exit Sub
    lngEditingIdx = lvwLines.SelectedItem.Key
    Set oRLine = ooR.RLines.Item(lngEditingIdx)
    lngSelectedRowIndex = lvwLines.SelectedItem.Key
    txtCode = oRLine.CodeF
    txtTitle = oRLine.Title
    txtQty = oRLine.QtyRequested
    txtPrice = oRLine.Price(ooR.ISForeignCurrency)
    txtDiscount = oRLine.Discount
    txtSuppRef = oRLine.SINVRef
    txtNote = oRLine.Note
'    txtTotal = oRLine.PLessDiscExtF(ooR.isFOreignCurrency)
'    tlProductTypes.Item (oRLine.ProductTypeID)
    ReloadMatches oRLine.PID
    LoadMatches
    If oRLine.DELLID > 0 Then cboMatch.Items.SelectItem(cboMatch.Items.FindItem(oRLine.DELLID, 8)) = True
'    AutoSelect txtPrice
    lvwLines.Enabled = False
    SetEditFrameEnabled True, eneditingrow
    vMode = eneditingrow
    lvwLines.Height = 2600
    cmdNewRows.Caption = "&Stop edit"
    oRLine.ValidateObject ""
    oRLine.GetStatus
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.lvwLines_DblClick", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.lvwLines_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboTP_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If ooR.Supplier Is Nothing Then
        MsgBox "Please enter a Supplier before continuing", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Cancel = True
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.cboTP_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.cboTP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtCode_LostFocus()
    On Error GoTo errHandler
    If Trim(txtCode) > "" Then
        Sendkeys "{F4}", True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.txtCode_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not IsNumeric(txtDiscount) Then
        txtDiscount.BackColor = vbRed
        Cancel = True
        Exit Sub
    Else
        txtDiscount.BackColor = vbWhite
    End If
    oRLine.Discount = txtDiscount

    ooR.CalculateTotals
    txtRunningTotal = ooR.TotalLessDiscExtF(ooR.ISForeignCurrency)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.txtDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtNote_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    If flgLoading Then Exit Sub
    txtNote = HandleTextWithBites(txtNote)
    On Error Resume Next
    oRLine.Note = Me.txtNote
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oRLine.Note
      txtNote.SelStart = intPos - 1
    End If
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oRLine.Note = txtNote
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtNote = oRLine.Note
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.txtNote_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuVoid()
    On Error GoTo errHandler
    ooR.SetStatus stVOID
    txtStatus = "Void"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.mnuVoid"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.mnuVoid"
End Sub

Private Sub txtCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim pQty As Integer
Dim pApproID As Long
Dim bOK  As Boolean
Dim oPCode As New z_ProdCode

START:
    If txtCode = "" Then Exit Sub
    If Not (IsISBN13(txtCode) Or IsISBN10(txtCode) Or IsHashCode(txtCode) Or IsPrivateCode(txtCode)) Then
        MsgBox "This is an invalid code, retry.", vbInformation, "Warning"
        Cancel = True
        GoTo EXIT_Handler
    End If
    bOK = oRLine.SetLineProduct("", txtCode)
    If bOK Then
        txtQty = oRLine.QtyRequested
        txtTitle = oRLine.Title
        ReloadMatches oRLine.PID
        If cDELLSPerPIDTP.Count = 0 Then
            MsgBox "This book has never been received or has been returned already from " & ooR.Supplier.Name, , "Status"
            Me.txtPrice.Visible = True
        Else
  '          Me.txtPrice.Visible = False
            LoadMatches
        End If
    Else
        MsgBox "This book has never been received"
    End If
    If Me.cboMatch.Items.ItemCount > 0 Then
        cboMatch.Items.SelectItem(cboMatch.Items(0)) = True
    End If
    oRLine.GetStatus

EXIT_Handler:

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub cboMatch_Validate(Cancel As Boolean)
    On Error GoTo errHandler
        LoadRLFromcboMatch
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.cboMatch_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub LoadRLFromcboMatch()
    On Error GoTo errHandler
    If oRLine Is Nothing Then Exit Sub
    If cboMatch.Items.SelectCount = 0 Then
        oRLine.TRID = 0
        oRLine.DELLID = 0
        oRLine.DocRef = ""
        oRLine.DOCDate = CDate(0)
        oRLine.ForeignPrice = 0
        oRLine.Discount = 0
        oRLine.SINVRef = ""
        Exit Sub
    End If
    oRLine.TRID = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 9)
    oRLine.DELLID = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 8)
    oRLine.DocRef = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 0)
    oRLine.DOCDate = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 1)
    oRLine.LocalPrice = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 10)
    oRLine.ForeignPrice = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 11)
    oRLine.Discount = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 7)
    oRLine.SINVRef = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 5) & "   " & cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 6)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.LoadRLFromcboMatch"
End Sub
Private Sub RemoveDetailLine()
    On Error GoTo errHandler
Dim i As Integer
Dim iMax As Integer
    iMax = lvwLines.ListItems.Count
    For i = iMax To 1 Step -1
        If lvwLines.ListItems(i).Selected Then
            ooR.RLines.Remove lvwLines.ListItems(i).Key
            Exit For
        End If
    Next i
    If i = 0 Then
        MsgBox "Select an item prior to deleting.", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Exit Sub
    End If
    lvwLines.ListItems.Remove i
    lvwLines.Refresh
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.RemoveDetailLine"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.RemoveDetailLine"
End Sub

Private Sub LoadSupplier()
    On Error GoTo errHandler
    With ooR
        txtStatus = .StatusF
        SetIssueButtonCaption
        Me.txtSuppname = .Supplier.NameAndCode(20)
        If Not .Supplier.BillTOAddress Is Nothing Then
            Me.txtPhone = .Supplier.BillTOAddress.Phone
            Me.txtFax = .Supplier.BillTOAddress.Fax
        End If
    End With
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.LoadSupplier"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.LoadSupplier"
End Sub


Private Sub SaveR()
    On Error GoTo errHandler
  '  If ooR.RLines.IsEditing Then ooR.RLines.ApplyEdit
    ooR.Post
    
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
   ' Resume
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.SaveR"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.SaveR"
End Sub

'Public Sub PrintOrder()
'Dim blnDeposit As Boolean
'Dim blnDiscount As Boolean
'Dim blnRoundedUp As Boolean
'Dim blnNoCNLines As Boolean
'Dim blnHideVAT As Boolean
'Dim iCurrency As Integer
'
'    On Error GoTo ERR_Handler
'
'    Me.MousePointer = vbHourglass
'    ooR.Load ooR.TRID, False
'    blnDiscount = False ' TO BE REMOVED ON COMPLETION????
'
'    If blnNoCNLines Then
'        MsgBox "There are no records to print on this invoice.", vbOKOnly + vbInformation, "Papyrus Invoicing Status"
'        GoTo EXIT_Handler
'    End If
'
'EXIT_Handler:
'    Me.MousePointer = vbDefault
'    Exit Sub
'ERR_Handler:
'    Select Case Err
'    Case 5941
'        MsgBox "Book Mark on word document is missing", vbOKOnly + vbInformation, "Papyrus Information"
'        Resume Next
'    Case Else
'        MsgBox Error
'        GoTo EXIT_Handler
'    End Select
'    Resume
'End Sub
Private Sub cmdIssue_Click()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoCNLines As Boolean
Dim iCurrency As Integer
'Dim ViewOrPrint As PreviewPrint
Dim strResult As String
Dim frm As frmReturn3
Dim frmAPP As frmApproval
    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_RETFIN_SIGN, , "Sign this return", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    End If
   ' SaveR
   ' ooR.BeginEdit
    If ooR.Status = stInProcess Then
      '  If MsgBox("Request this return?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
      '      Exit Sub
      '  Else
            ooR.SetStatus stISSUED
      '  End If
    ElseIf ooR.Status = stISSUED Then
        If MsgBox("Issue this return?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
            Exit Sub
        Else
            ooR.SetStatus stCOMPLETE
        End If
    End If
    
    If ooR.Status = stCOMPLETE Then
        Set frmAPP = New frmApproval
        frmAPP.Show vbModal
        ooR.ApprovalRef = frmAPP.ApprovalRef
        ooR.ApprovalTermDate = frmAPP.ApprovalDate
        Unload frmAPP
        Set frmAPP = Nothing
    End If
    
    WaitMsg "Issuing return  . . .", True, Me
    ooR.StaffID = gSTAFFID
    
    
    strResult = ooR.Post
    If strResult = "ERROR" Then
        MsgBox "This action has failed. Contact support"
        Exit Sub
    End If
    Set frm = New frmReturn3
    frm.Component2 ooR.TRID
    frm.Show
    WaitMsg "", False, Me
    Unload Me

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.cmdIssue_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.cmdIssue_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdSave_Click()
    On Error GoTo errHandler
    SaveR
    ooR.BeginEdit
    cmdCancel.Caption = "&Close"
    cmdSave.Enabled = False
    mSetfocus cmdNewRows
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.cmdSave_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
Dim frm As frmReturnPreview
    If cmdCancel.Caption <> "&Close" Then
        If MsgBox("You wish to cancel this return?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
            Exit Sub
        End If
    End If
    ooR.CancelEdit
    If cmdCancel.Caption = "&Close" Then
        Set frm = New frmReturnPreview
        frm.ComponentObject ooR
        frm.Show
    End If
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.cmdCancel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub ClearLineControls()
    On Error GoTo errHandler
    flgLoading = True
    txtCode = ""
    txtTitle = ""
    txtPrice = ""
    txtDiscount = ""
    txtNote = ""
    txtQty = ""
    txtSuppRef = ""
    cboMatch.BeginUpdate
    cboMatch.Items.RemoveAllItems
    cboMatch.EndUpdate
   ' cboMatch.Items.SelectItem(cboMatch.Items(0)) = True
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.ClearLineControls"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.ClearLineControls"
End Sub


Private Sub txtPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    
    If flgLoading Then Exit Sub
    If ooR.ISForeignCurrency Then
        If Not oRLine.SetPrice(txtPrice) Then
            Cancel = True
        End If
    Else
        If Not oRLine.SetPrice(txtPrice) Then
            Cancel = True
        End If
    End If
    ooR.CalculateTotals
    txtRunningTotal = ooR.TotalLessDiscExtF(ooR.ISForeignCurrency)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.txtPrice_Validate"
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

'Private Sub txtPrice_GotFocus()
'    AutoSelect txtPrice
'End Sub
'Private Sub txtPrice_Validate(Cancel As Boolean)
'    If flgLoading Then Exit Sub
'    If Not oRLine.SetPrice(txtPrice) Then
'        Cancel = True
'    End If
'    ooR.CalculateTotals
'    txtTotal = oRLine.PLessDiscExtF(ooR.isFOreignCurrency)
'End Sub
Private Sub txtQty_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtQty
    LoadRLFromcboMatch
Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.txtQty_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oRLine.SetQtyRequested(txtQty) Then
        Cancel = True
    End If
    ooR.CalculateTotals
 '   txtTotal = oRLine.PLessDiscExtF(ooR.isFOreignCurrency)
    txtRunningTotal = ooR.TotalLessDiscExtF(ooR.ISForeignCurrency)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub SetIssueButtonCaption()
    On Error GoTo errHandler
        If ooR.Status = stInProcess Then
            cmdIssue.Caption = "1. Request"
        ElseIf ooR.Status = stISSUED Then
            cmdIssue.Caption = "2. Return"
'        ElseIf ooR.IsDirty Then
'            cmdIssue.Caption = "Save"
        Else: cmdIssue.Enabled = False
        End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.SetIssueButtonCaption"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.SetIssueButtonCaption"
End Sub

Private Sub lvwLines_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    On Error GoTo errHandler
   lvwLines.SortKey = ColumnHeader.Index - 1
    If lvwLines.SortOrder = lvwAscending Then
        lvwLines.SortOrder = lvwDescending
    Else
        lvwLines.SortOrder = lvwAscending
    End If
   lvwLines.Sorted = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.lvwLines_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.lvwLines_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
    HandleError
End Sub
Private Sub SetLvw()
    On Error GoTo errHandler
Dim Style As Long
Dim hHeader As Long
   
  'get the handle to the listview header
   hHeader = SendMessage(lvwLines.hWnd, LVM_GETHEADER, 0, ByVal 0&)
   
  'get the current style attributes for the header
   Style = GetWindowLong(hHeader, GWL_STYLE)
   
  'modify the style by toggling the HDS_BUTTONS style
   Style = Style Xor HDS_BUTTONS
   
  'set the new style and redraw the listview
   If Style Then
      Call SetWindowLong(hHeader, GWL_STYLE, Style)
      Call SetWindowPos(lvwLines.hWnd, Me.hWnd, 0, 0, 0, 0, SWP_FLAGS)
   End If


'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.SetLvw"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.SetLvw"
End Sub
Sub SetupcboMatch()
    On Error GoTo errHandler
    
    
    cboMatch.BeginUpdate
    cboMatch.WidthList = 502
    cboMatch.HeightList = 162
    cboMatch.AllowSizeGrip = False
    cboMatch.AutoDropDown = True
 '   cboMatch.SelForeColor = vbRed
    cboMatch.Columns.Add "Doc. ref."
    cboMatch.Columns.Add "Date"
    cboMatch.Columns.Add "Qty rec."
    cboMatch.Columns.Add "Inv. price"
    cboMatch.Columns.Add "Inv. disc."
    cboMatch.Columns.Add "Supp. invoice"
    cboMatch.Columns.Add "Supp. invoice date"
    cboMatch.Columns.Add ""
    cboMatch.Columns.Add ""
    cboMatch.Columns.Add ""
    cboMatch.Columns.Add ""
    cboMatch.Columns.Add ""
    cboMatch.Columns(0).Width = 80
    cboMatch.Columns(1).Width = 80
    cboMatch.Columns(2).Width = 50
    cboMatch.Columns(3).Width = 65
    cboMatch.Columns(3).Alignment = RightAlignment
    cboMatch.Columns(4).Width = 65
    cboMatch.Columns(4).Alignment = RightAlignment
    cboMatch.Columns(5).Width = 90
    cboMatch.Columns(6).Width = 65
    cboMatch.Columns(7).Width = 0
    cboMatch.Columns(8).Width = 0
    cboMatch.Columns(9).Width = 0
    cboMatch.Columns(10).Width = 0
    cboMatch.Columns(11).Width = 0
    cboMatch.Columns(7).Visible = False
    cboMatch.Columns(8).Visible = False
    cboMatch.Columns(9).Visible = False
    cboMatch.Columns(10).Visible = False
    cboMatch.Columns(11).Visible = False
'    cboMatch.BackColorLock = Me.BackColor
'    cboMatch.BackColor = RGB(220, 220, 220)
'    cboMatch.SelForeColor = vbBlack
'    cboMatch.SelBackColor = RGB(256, 256, 256)
    cboMatch.EndUpdate

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.SetupcboMatch"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.SetupcboMatch"
End Sub

Private Sub LoadMatches()
    On Error GoTo errHandler
Dim oDL As d_DELLSPerPIDTP
Dim i As Integer
Dim ar()
    If cDELLSPerPIDTP.Count = 0 Then Exit Sub
    cboMatch.BeginUpdate
    cboMatch.Items.RemoveAllItems
    i = 0
    ReDim ar(11, cDELLSPerPIDTP.Count - 1)
    For Each oDL In cDELLSPerPIDTP
        ar(0, i) = oDL.DocRef
        ar(1, i) = oDL.DOCDate
        ar(2, i) = oDL.Qty
        ar(3, i) = oDL.PriceF
        ar(4, i) = oDL.DiscountF
        ar(5, i) = oDL.SINVRef
        ar(6, i) = oDL.SINVDate
        ar(7, i) = oDL.Discount
        ar(8, i) = oDL.DELLID
        ar(9, i) = oDL.TRID
        ar(10, i) = oDL.LocalPrice
        ar(11, i) = oDL.ForeignPrice
        i = i + 1
    Next
    cboMatch.PutItems ar
    cboMatch.EndUpdate
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.LoadMatches"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.LoadMatches"
End Sub

Public Function ReloadMatches(pPID As String)
    On Error GoTo errHandler
    Set cDELLSPerPIDTP = Nothing
    Set cDELLSPerPIDTP = New c_DELLSPerPIDTP
    cDELLSPerPIDTP.Load ooR.Supplier.ID, pPID
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn.ReloadMatches(pPID)", pPID
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.ReloadMatches(pPID)", pPID
End Function


Private Sub txtSuppRef_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oRLine.SINVRef = txtSuppRef
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn.txtSuppRef_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
