VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAPPR 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Appro return"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11595
   ControlBox      =   0   'False
   Icon            =   "frmApproRet.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   11595
   Begin MSComctlLib.ListView lvwLines 
      Height          =   3330
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5874
      SortKey         =   4
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
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title / Author / Publisher"
         Object.Width           =   9349
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Appro code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Key"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   0
      EndProperty
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
      Left            =   8625
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmApproRet.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5370
      UseMaskColor    =   -1  'True
      Width           =   1110
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
      Height          =   855
      Left            =   1095
      MultiLine       =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5340
      Width           =   3390
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
      Height          =   690
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5400
      Width           =   780
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
      Left            =   7515
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmApproRet.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5370
      Width           =   1110
   End
   Begin VB.TextBox txtRunningTotal 
      Alignment       =   1  'Right Justify
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
      Height          =   250
      Left            =   9570
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3405
      Width           =   1200
   End
   Begin VB.CommandButton cmdIssue 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Issue"
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
      Left            =   9750
      Picture         =   "frmApproRet.frx":085E
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5370
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00D3D3CB&
      Height          =   1605
      Left            =   120
      TabIndex        =   9
      Top             =   3735
      Width           =   10650
      Begin EXCOMBOBOXLibCtl.ComboBox cboAPPL 
         Height          =   315
         Left            =   1800
         OleObjectBlob   =   "frmApproRet.frx":0BE8
         TabIndex        =   15
         Top             =   465
         Width           =   6735
      End
      Begin VB.TextBox txtLastAt 
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
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1170
         Width           =   3975
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8625
         TabIndex        =   3
         Top             =   465
         Width           =   1035
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
         Height          =   705
         Left            =   9735
         MaskColor       =   &H00C4BCA4&
         Picture         =   "frmApproRet.frx":1F92
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   825
         Width           =   840
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
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   855
         Width           =   3555
      End
      Begin VB.TextBox txtCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   465
         Width           =   1650
      End
      Begin VB.Label Label4 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Qty"
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
         Height          =   300
         Left            =   8985
         TabIndex        =   14
         Top             =   195
         Width           =   390
      End
      Begin VB.Label Label9 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Code"
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
         Height          =   285
         Left            =   660
         TabIndex        =   11
         Top             =   210
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmAPPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oAPPR As a_APPR
Attribute oAPPR.VB_VarHelpID = -1
Dim WithEvents oAPPRLine As a_APPRL
Attribute oAPPRLine.VB_VarHelpID = -1
Dim oCustomer As a_Customer
Dim oProd As a_Product
Dim oCurrentCopy
Dim bValidAPP As Boolean
Dim bValidCOLine As Boolean
Dim tlCustomer As z_TextList
Dim lngCurrentExtension As Long
Dim lngCurrentTotal As Long
Dim lngCurrentDepositTotal As Long
Dim lngCurrentVATTotal As Long

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
Dim cAPPL As New c_APPLsPerTPPID
Dim blnReadOnly As Boolean
Dim flgLoading As Boolean
Dim WithEvents vCanAdd As z_BrokenRules
Attribute vCanAdd.VB_VarHelpID = -1
Dim WithEvents vCanIssue As z_BrokenRules
Attribute vCanIssue.VB_VarHelpID = -1
Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayoutLvw Me.lvwLines, Me.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.mnuSaveLayout"
End Sub

Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
    ofrm.component oAPPR.MEMO
    ofrm.Show vbModal
    oAPPR.SetMemo ofrm.MEMO
    Unload ofrm
    Set ofrm = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.mnuMemo"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.mnuMemo"
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oAPPR.StatusF = "IN PROCESS" And oAPPR.IsNew = False)
    Forms(0).mnuDelLine.Enabled = True
    Forms(0).mnuCancelLine.Enabled = (oAPPR.StatusF = "ISSUED" And oAPPR.IsNew = False)
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmCO.SetMenu"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.SetMenu"
End Sub

Public Sub component(Optional pAPP As a_APPR, Optional pCustID As Long)
    On Error GoTo errHandler
    flgLoading = True
    If pAPP Is Nothing Then
        Set oAPPR = New a_APPR
        oAPPR.BeginEdit
        oAPPR.SetStatus stInProcess
        lvwLines.Enabled = False
        If pCustID > 0 Then
            LoadNewCustomer pCustID
        End If
        Me.Caption = "Appro return for " & oAPPR.Customer.NameAndCode(40)
        lvwLines.Height = 3500
        cmdNewRows.Caption = "&Stop"
        vMode = enAddingRow
        SetEditFrameEnabled True, vMode
        mSetfocus txtCode
        Set oAPPRLine = oAPPR.APPRLines.Add
        oAPPRLine.SetQty 1
        oAPPR.GetStatus
        mSetfocus txtCode
    Else
        Set oAPPR = pAPP
        oAPPR.BeginEdit
        LoadCustomer
        LoadListView
        Me.Caption = "Appro return for " & oAPPR.Customer.NameAndCode(40)
        cmdSave.Enabled = False
        cmdIssue.Enabled = False
        cmdCancel.Caption = "&Close"
        cmdNewRows.Enabled = True
        lvwLines.Enabled = True
        lvwLines.Height = 5200
        SetEditFrameEnabled False, enNotEditing
        vMode = enNotEditing
        ClearLineControls
    End If
    SetMenu
    flgLoading = False

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.Component(pAPP,pCustID)", Array(pAPP, pCustID)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.component(pAPP,pCustID)", Array(pAPP, pCustID)
End Sub




Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Public Sub mnuDelLine()
    On Error GoTo errHandler
    RemoveDetailLine
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.mnuDelLine"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.mnuDelLine"
End Sub
Public Sub mnuVoid()
    On Error GoTo errHandler
Dim strResult As String
    oAPPR.SetStatus stVOID
    oAPPR.ApplyEdit
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.mnuVoid"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.mnuVoid"
End Sub

Private Sub RemoveDetailLine()
    On Error GoTo errHandler
Dim i As Integer
Dim iMax As Integer
    iMax = lvwLines.ListItems.Count
    For i = iMax To 1 Step -1
        If lvwLines.ListItems(i).Selected Then
            oAPPR.APPRLines.Remove lvwLines.ListItems(i).key
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
'    ErrorIn "frmAPPR.RemoveDetailLine"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.RemoveDetailLine"
End Sub

Private Sub lvwLines_Click()
    On Error GoTo errHandler
    
    On Error Resume Next
    Clipboard.Clear
        Clipboard.SetText Left(lvwLines.SelectedItem.SubItems(5), ISBNLENGTH)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.lvwLines_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.lvwLines_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwLines_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.lvwLines_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
'         EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.lvwLines_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub


Private Sub mnuDel_Click()
    On Error GoTo errHandler
    RemoveDetailLine
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.mnuDel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.mnuDel_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_APP
    Set frm = New frmPrintingOptions_APP
    frm.Show vbModal

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.mnuPrint_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.mnuPrint_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub oAPPR_Valid(pMsg As String)
    On Error GoTo errHandler
    bValidAPP = (pMsg = "")
    cmdIssue.Enabled = (bValidAPP And oAPPR.APPRLines.Count > 0 And vMode = enNotEditing)
    cmdSave.Enabled = (bValidAPP And oAPPR.APPRLines.Count > 0 And vMode = enNotEditing)
    Me.txtError = pMsg
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.oAPPR_Valid(pMsg)", pMsg, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.oAPPR_Valid(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub

Sub oAPPRLine_ExtensionChange(lngExtension As Long, strExtension As String)
    On Error GoTo errHandler
    flgLoading = True
   ' Me.txtTotal = strExtension
    flgLoading = False
    lngCurrentExtension = lngExtension
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.oAPPRLine_ExtensionChange(lngExtension,strExtension)", Array(lngExtension, _
'         strExtension), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.oAPPRLine_ExtensionChange(lngExtension,strExtension)", Array(lngExtension, _
         strExtension), EA_NORERAISE
    HandleError
End Sub

Private Sub oAPPRLine_Valid(msg As String)
    On Error GoTo errHandler
    cmdEnter.Enabled = (msg = "")
    txtError = msg
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.oAPPRLine_Valid(Msg)", msg, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.oAPPRLine_Valid(msg)", msg, EA_NORERAISE
    HandleError
End Sub

Private Sub oAPPR_TotalChange(lngTotal As Long, strtotal As String, lngTotalDeposit As Long, strTotalDeposit As String, lngTotalVAT As Long, strTotalVAT As String)
    On Error GoTo errHandler
    flgLoading = True
    Me.txtRunningTotal = strtotal
    lngCurrentTotal = lngTotal
'    Me.txtRunningDeposit = strTotalDeposit
    lngCurrentDepositTotal = lngTotalDeposit
    lngCurrentVATTotal = lngTotalVAT
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.oAPPR_TotalChange(lngTotal,strtotal,lngTotalDeposit,strTotalDeposit,lngTotalVAT," & _
'        "strTotalVAT)", Array(lngTotal, strtotal, lngTotalDeposit, strTotalDeposit, lngTotalVAT, strTotalVAT), _
'         EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.oAPPR_TotalChange(lngTotal,strtotal,lngTotalDeposit,strTotalDeposit,lngTotalVAT," & _
        "strTotalVAT)", Array(lngTotal, strtotal, lngTotalDeposit, strTotalDeposit, lngTotalVAT, strTotalVAT), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub oAPPR_Reloadlist()
    On Error GoTo errHandler
    LoadListView
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.oAPPR_Reloadlist", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.oAPPR_Reloadlist", , EA_NORERAISE
    HandleError
End Sub
Private Sub oAPPR_Dirty(pVal As Boolean)
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
'    ErrorIn "frmAPPR.oAPPR_Dirty(pVal)", pVal, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.oAPPR_Dirty(pVal)", pVal, EA_NORERAISE
    HandleError
End Sub

Sub vCanAdd_NobrokenRules()
    On Error GoTo errHandler
    Me.cmdNewRows.Enabled = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.vCanAdd_NobrokenRules", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.vCanAdd_NobrokenRules", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
Dim strAddress As String
    If Me.WindowState <> 2 Then
        Left = 10
        top = 10
        Width = 11100
        Height = 6700
    End If
    flgLoading = True
    oAPPR.GetStatus
    SetLvw
    SetEditFrameEnabled False, enNotEditing
    vMode = enNotEditing
    SetupcboAPPL
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Initialize()
    On Error GoTo errHandler
    Set vCanAdd = New z_BrokenRules
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.Form_Initialize", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If oAPPR.IsEditing Then oAPPR.CancelEdit
    UnsetMenu
    
    Set oCustomer = Nothing
    Set oCurrentCopy = Nothing
    Set oAPPR = Nothing
    Set tlCustomer = Nothing
    Set oAPPRLine = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

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
    Me.txtTitle.Enabled = pYesNo
    Me.txtQty.Enabled = pYesNo
    
    Me.cmdEnter.Enabled = Not pYesNo
    Me.cmdCancel.Enabled = Not pYesNo
    Me.cmdIssue.Enabled = (Not pYesNo) And bValidAPP
    Me.cmdSave.Enabled = (Not pYesNo) And bValidAPP And oAPPR.IsDirty
    
    If pYesNo Then
        lngColour = &HFFFFFF
    Else
        lngColour = 14416635
    End If
    
    Me.txtCode.BackColor = lngColour
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.SetEditFrameEnabled(pYesNo,eMode)", Array(pYesNo, eMode)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.SetEditFrameEnabled(pYesNo,eMode)", Array(pYesNo, eMode)
End Sub

Private Sub cmdEnter_Click()
    On Error GoTo errHandler
Dim currDeposit As Currency
Dim blnResult As Boolean
Dim strCurrFormat As String
Dim curTotalDeposit As Currency
    
    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Papyrus appro Information"
        ClearLineControls
        Exit Sub
    End If
    oAPPRLine.ApplyEdit
    oAPPRLine.BeginEdit

    If vMode = enAddingRow Then
        lvwLines.ListItems.Add key:=oAPPRLine.key
        LoadListViewLine lvwLines.ListItems(lvwLines.ListItems.Count), oAPPRLine
        lvwLines.Refresh
        
'        Set oAPPRLine = Nothing
'        Set oAPPRLine = oAPPR.APPRLines.Add
'        oAPPRLine.SetQty 1
'        oAPPRLine.TRID = oAPPR.TRID
        
        ChangeState enAddingRow
        mSetfocus txtCode
    ElseIf vMode = eneditingrow Then
        LoadListViewLine lvwLines.ListItems(lngSelectedRowIndex), oAPPRLine
        ChangeState enNotEditing
'        LoadListViewLine lvwLines.ListItems(lngSelectedRowIndex)
'        ClearLineControls
'        lvwLines.Enabled = True
'        lvwLines.Height = 4020
'        vMode = enNotEditing
'        SetEditFrameEnabled False, vMode
'        cmdNewRows.Caption = "&Add"
'        fr1.ZOrder 1
'        mSetfocus cmdNewRows
    End If
    LoadListView
    
    oAPPR.GetStatus
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.cmdEnter_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.cmdEnter_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub ChangeState(pToMode As EnumMode)
    On Error GoTo errHandler
Dim lngColour As Long
    vMode = pToMode

    Select Case pToMode
    Case eneditingrow
        fr1.Visible = True
        txtCode.Enabled = True
        txtTitle.Enabled = True
 '       txtTotal.Enabled = True
        txtQty.Enabled = True
 '       cboRef.Visible = False
        cmdEnter.Enabled = False
        cmdCancel.Enabled = False
        cmdIssue.Enabled = False
        cmdSave.Enabled = False
        cmdNewRows.Caption = "&Stop"
        cmdNewRows.Enabled = (oAPPR.APPRLines.Count > 0)
        lvwLines.Enabled = False
        lvwLines.Height = 2200
        UnsetMenu
        fr1.ZOrder 1
    Case enAddingRow
        fr1.Visible = True
        txtCode.Enabled = True
        txtTitle.Enabled = True
  '      txtTotal.Enabled = True
        txtQty.Enabled = True
        txtError = ""
        flgLoading = True
 '       txtRef = ""
        flgLoading = False
        cmdEnter.Enabled = False
        cmdCancel.Enabled = True
        cmdIssue.Enabled = False
        cmdSave.Enabled = False
        cmdNewRows.Enabled = (oAPPR.APPRLines.Count > 0)
        cmdNewRows.Caption = "&Stop"
        lvwLines.Enabled = False
        lvwLines.Height = 2200
        ClearLineControls
        fr1.ZOrder 1
        mSetfocus txtCode
        Set oAPPRLine = oAPPR.APPRLines.Add
        oAPPRLine.TRID = oAPPR.TRID
        oAPPRLine.SetQty 1
        oAPPRLine.TRID = oAPPR.TRID
        
        UnsetMenu
    Case enNotEditing
        flgLoading = True
        fr1.Visible = False
        txtError = ""
        flgLoading = False
        cmdEnter.Enabled = False
        cmdCancel.Enabled = True
        cmdIssue.Enabled = True
        cmdSave.Enabled = True
        cmdNewRows.Enabled = True  '(oInvoice.InvoiceLines.Count > 0)
        cmdNewRows.Caption = "&Add"
        lvwLines.Enabled = True
        lvwLines.Height = 4000
        SetMenu
        fr1.ZOrder 1
    End Select
    If Not oAPPR.IsDirty Then
        cmdCancel.Caption = "&Close"
    Else
        cmdCancel.Caption = "&Cancel"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoice.ChangeState(pToMode)", pToMode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.ChangeState(pToMode)", pToMode
End Sub

Private Sub LoadNewCustomer(plngTPID As Long)
    On Error GoTo errHandler
    If oAPPR.SetCustomer(plngTPID) Then
        With oAPPR.Customer
'            txtPhone = .phone
'            txtCustName = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
            oAPPR.TPID = plngTPID
        End With
        vCanAdd.RuleBroken "TP", False
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.LoadNewCustomer(plngTPID)", plngTPID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.LoadNewCustomer(plngTPID)", plngTPID
End Sub


Private Sub cmdNewRows_Click()
    On Error GoTo errHandler
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
    If vMode = eneditingrow Then
        LogSaveToFile "Invoice New row button:enEditingRow"
        ChangeState enNotEditing
    ElseIf vMode = enAddingRow Then
        LogSaveToFile "Invoice New row button:enAddingRow"
        ChangeState enNotEditing
    ElseIf vMode = enNotEditing Then
        LogSaveToFile "Invoice New row button:enNotEditing"
        ChangeState enAddingRow
    End If

    ClearLineControls
   
'    If vMode = enEditingRow Then
'        cmdNewRows.Caption = "&Add"
'        vMode = enNotEditing
'        Me.lvwLines.Enabled = True
'        lvwLines.Height = 5200
'        Me.fr1.ZOrder 1
'        SetEditFrameEnabled False, vMode
'    ElseIf vMode = enAddingRow Then
'        vMode = enNotEditing
'        cmdNewRows.Caption = "&Add"
'        Me.lvwLines.Enabled = True
'        lvwLines.Height = 5200
'        Me.fr1.ZOrder 1
'        SetEditFrameEnabled False, vMode
'        oAPPR.GetStatus
'    ElseIf vMode = enNotEditing Then
'        vMode = enAddingRow
'        Set oAPPRLine = oAPPR.APPRLines.Add
'        oAPPRLine.TRID = oAPPR.TRID
'        cmdNewRows.Caption = "&Stop"
'        Me.lvwLines.Enabled = False
'        lvwLines.Height = 3500
'        SetEditFrameEnabled True, vMode
'        mSetfocus txtCode
'    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.cmdNewRows_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.cmdNewRows_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadListView()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Long
    For i = 1 To lvwLines.ColumnHeaders.Count
        lvwLines.ColumnHeaders(i).Width = GetSetting("PBKS", Me.Name, CStr(i), lvwLines.ColumnHeaders(i).Width)
    Next
    lvwLines.ListItems.Clear
    For i = 1 To oAPPR.APPRLines.Count
        Set lstItem = lvwLines.ListItems.Add
        LoadListViewLine lstItem, oAPPR.APPRLines(i)
'        lstItem.SubItems(4) = oAPPR.APPRLines(i).Key
'        Set oAPPRLine = oAPPR.APPRLines(i)
'        LoadListViewLine i & "k", lstItem
    Next i
EXIT_Handler:
    Set lstItem = Nothing
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.LoadListView"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.LoadListView"
End Sub
Private Sub LoadListViewLine(lstItem As ListItem, oAPPRLine As a_APPRL)
    On Error GoTo errHandler
Dim currPrice As Currency
    With oAPPRLine
        lstItem.Text = .CodeF
        lstItem.key = .key
        lstItem.SubItems(1) = .Title
        lstItem.SubItems(2) = .qty & " (" & .QtyIssued & ")"
        lstItem.SubItems(3) = .ApproCode
        lstItem.SubItems(4) = Format(.key, "@@@@@@@@@@")
        lstItem.SubItems(5) = .EAN
    End With
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.LoadListViewLine(lstItem,oAPPRLine)", Array(lstItem, oAPPRLine)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.LoadListViewLine(lstItem,oAPPRLine)", Array(lstItem, oAPPRLine)
End Sub
Private Sub lvwLines_DblClick()
    On Error GoTo errHandler
'This must load the editing line with the current line's data
    If lvwLines.SelectedItem.Index < 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    If lvwLines.ListItems.Count = 0 Then Exit Sub
    
    lngEditingIdx = lvwLines.SelectedItem.key
    Set oAPPRLine = Nothing
    Set oAPPRLine = oAPPR.APPRLines(lvwLines.SelectedItem.key)
    cAPPL.Load oAPPR.Customer.ID, oAPPRLine.PID, ""
    LoadcboAPPL True
    cboAPPL.Items.SelectItem(cboAPPL.Items.FindItem(oAPPRLine.APPLID, 5)) = True
    lngSelectedRowIndex = lvwLines.SelectedItem.key
    
    ChangeState eneditingrow
    
    Me.txtCode = oAPPRLine.EAN
    Me.txtTitle = oAPPRLine.Title
    Me.txtQty = oAPPRLine.qty
    
    mSetfocus txtQty
    oAPPRLine.GetStatus
    
    Screen.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.lvwLines_DblClick", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.lvwLines_DblClick", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuEditNote_Click()
    On Error GoTo errHandler
Dim ofrm As New frmNote
    ofrm.component oAPPR
    ofrm.Show vbModal
    Unload ofrm
    Set ofrm = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.mnuEditNote_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.mnuEditNote_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuFileCancel_Click()
    On Error GoTo errHandler
    If oAPPR.IsDirty Then
        oAPPR.CancelEdit
    End If
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.mnuFileCancel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.mnuFileCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuFileExit_Click()
    On Error GoTo errHandler
    oAPPR.CancelEdit
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.mnuFileExit_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.mnuFileExit_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuFilePrint_Click()
    On Error GoTo errHandler
    cmdIssue_Click
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.mnuFilePrint_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.mnuFilePrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuFileVoid_Click()
    On Error GoTo errHandler
Dim strResult As String
    oAPPR.SetStatus stVOID
    oAPPR.ApplyEdit
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.mnuFileVoid_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.mnuFileVoid_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim pQty As Integer
Dim pApproID As Long
Dim bOK  As Boolean
Dim oP  As New a_Product
Dim strPID As String
Dim strEAN As String
Dim strCode As String

    If txtCode = "" Or vMode = eneditingrow Then Exit Sub
    bOK = oP.Load("", 0, Trim$(txtCode), , True) <> 99

    If Not bOK Then Exit Sub
    cAPPL.Load oAPPR.Customer.ID, oP.PID, ""
    LoadcboAPPL False
    If cboAPPL.Items.ItemCount > 0 Then
        cboAPPL.Items.SelectItem(cboAPPL.Items(0)) = True
        oAPPRLine.APPLID = cboAPPL.Items.CellCaption(cboAPPL.Items.SelectedItem, 5)
        oAPPRLine.qty = cboAPPL.Items.CellCaption(cboAPPL.Items.SelectedItem, 2)
        oAPPRLine.Title = cboAPPL.Items.CellCaption(cboAPPL.Items.SelectedItem, 4)
        oAPPRLine.CodeF = cboAPPL.Items.CellCaption(cboAPPL.Items.SelectedItem, 6)
        oAPPRLine.EAN = oP.EAN
        oAPPRLine.ApproCode = cboAPPL.Items.CellCaption(cboAPPL.Items.SelectedItem, 0)
    Else
        oAPPRLine.APPLID = 0
    End If
    txtQty = oAPPRLine.qty
EXIT_Handler:
    Set oProd = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtCode_LostFocus()
    On Error GoTo errHandler
        If txtCode > "" Then SendKeys "{DOWN}", True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.txtCode_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.txtCode_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboAPPL_SelectionChanged()
    On Error GoTo errHandler
    oAPPRLine.APPLID = cboAPPL.Items.CellCaption(cboAPPL.Items.SelectedItem, 5)
    oAPPRLine.QtyIssued = cboAPPL.Items.CellCaption(cboAPPL.Items.SelectedItem, 2)
    oAPPRLine.QtyReturned = cboAPPL.Items.CellCaption(cboAPPL.Items.SelectedItem, 3)
    oAPPRLine.qty = oAPPRLine.QtyIssued - oAPPRLine.QtyReturned
    oAPPRLine.Title = cboAPPL.Items.CellCaption(cboAPPL.Items.SelectedItem, 4)
    oAPPRLine.code = cboAPPL.Items.CellCaption(cboAPPL.Items.SelectedItem, 6)
    oAPPRLine.ApproCode = cboAPPL.Items.CellCaption(cboAPPL.Items.SelectedItem, 0)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.cboAPPL_SelectionChanged", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.cboAPPL_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadCustomer()
    On Error GoTo errHandler
Dim strAddress As String
    With oAPPR
        SetIssueButtonCaption
    End With
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.LoadCustomer"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.LoadCustomer"
End Sub


Private Sub SaveAPPR()
    On Error GoTo errHandler
    
    oAPPR.Post
    
EXIT_Handler:
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.SaveAPPR"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.SaveAPPR"
End Sub

Public Sub PrintApproR()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoCOLines As Boolean
Dim blnHideVAT As Boolean
Dim iCurrency As Integer

    
    Me.MousePointer = vbHourglass
    oAPPR.Load oAPPR.TRID, False
    blnDiscount = False ' TO BE REMOVED ON COMPLETION????
    
    If blnNoCOLines Then
        MsgBox "There are no records to print on this invoice.", vbOKOnly + vbInformation, "Papyrus Invoicing Status"
        GoTo EXIT_Handler
    End If
    
EXIT_Handler:
    Me.MousePointer = vbDefault
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.PrintApproR"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.PrintApproR"
End Sub
Private Sub cmdIssue_Click()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoCOLines As Boolean
Dim iCurrency As Integer
Dim ar As arCOLSOS
Dim rs As ADODB.Recordset
Dim strResult As String
Dim frm As frmAPPRPreview
Dim tmpID As Long
Dim OpenResult As Integer

    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_APPR_SIGN, , "Sign this appro", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    Else
        If oAPPR.STATUS = stInProcess Then
            If MsgBox("Issue this appro?.  Confirm.", vbYesNo + vbQuestion, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
    End If

    WaitMsg "Issuing appro return  . . .", True, Me
    oAPPR.SetStatus stISSUED
    oAPPR.StaffID = gSTAFFID
    oAPPR.ApplyEdit
    oAPPR.Post
    tmpID = oAPPR.TRID
    Set frm = New frmAPPRPreview
    frm.component tmpID
    frm.Show
    WaitMsg "", False, Me
    
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.Open "Select * from vGetOSCOLSForAPPR WHERE APPRID = " & tmpID, oPC.COShort, adOpenStatic, adLockOptimistic
    If Not (rs.eof And rs.BOF) Then  'there are COLs awaiting
        Set ar = New arCOLSOS
        ar.component rs, "Customer orders outstanding for items returned"
        LogSaveToFile "Warning about " & FNS(rs.Fields("Title")) & " on document " & FNS(rs.Fields("DOCCODE")) & " outstanding for in ApproReturn ID " & CStr(tmpID)
        ar.Show
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.cmdIssue_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdSave_Click()
    On Error GoTo errHandler
    oAPPR.SetStatus stInProcess
    oAPPR.ApplyEdit
    oAPPR.BeginEdit
    Set oAPPRLine = oAPPR.APPRLines.Add
    cmdCancel.Caption = "&Close"
    cmdSave.Enabled = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.cmdSave_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
Dim frm As frmAPPRPreview
    If cmdCancel.Caption <> "&Close" Then
        If MsgBox("You wish to cancel this appro return?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
            Exit Sub
        End If
    End If
    oAPPR.CancelEdit
    If cmdCancel.Caption = "&Close" Then
        Set frm = New frmAPPRPreview
        frm.ComponentObject oAPPR
        frm.Show
    End If
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.cmdCancel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub ClearLineControls()
    On Error GoTo errHandler
    flgLoading = True
    Me.txtCode = ""
    Me.txtTitle = ""
    cboAPPL.Items.RemoveAllItems
'    cboAPPL
'    cboAPPL.Items.SelectItem(cboAPPL.Items(0)) = True
    Me.txtQty = ""
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.ClearLineControls"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.ClearLineControls"
End Sub

Private Sub lvwLines_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.lvwLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.lvwLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtQty")
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.txtQty_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.txtQty_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oAPPRLine.SetQty(txtQty) Then
        Cancel = True
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_LostFocus()
    On Error GoTo errHandler
  '  txtQty = oAPPRLine.QtyF
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.txtQty_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.txtQty_LostFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub SetIssueButtonCaption()
    On Error GoTo errHandler
        If oAPPR.StatusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oAPPR.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Caption = "Print"
        End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.SetIssueButtonCaption"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.SetIssueButtonCaption"
End Sub


'Private Sub lvwLines_ColumnClick(ByVal ColumnHeader As ColumnHeader)
'   ' When a ColumnHeader object is clicked, the ListView control is
'   ' sorted by the subitems of that column.
'   ' Set the SortKey to the Index of the ColumnHeader - 1
'   lvwLines.SortKey = ColumnHeader.Index - 1
'   ' Set Sorted to True to sort the list.
'    If lvwLines.SortOrder = lvwAscending Then
'        lvwLines.SortOrder = lvwDescending
'    Else
'        lvwLines.SortOrder = lvwAscending
'    End If
'   lvwLines.Sorted = True
'End Sub
Private Sub SetLvw()
    On Error GoTo errHandler
Dim style As Long
Dim hHeader As Long
   
  'get the handle to the listview header
   hHeader = SendMessage(lvwLines.hWnd, LVM_GETHEADER, 0, ByVal 0&)
   
  'get the current style attributes for the header
   style = GetWindowLong(hHeader, GWL_STYLE)
   
  'modify the style by toggling the HDS_BUTTONS style
   style = style Xor HDS_BUTTONS
   
  'set the new style and redraw the listview
   If style Then
      Call SetWindowLong(hHeader, GWL_STYLE, style)
      Call SetWindowPos(lvwLines.hWnd, Me.hWnd, 0, 0, 0, 0, SWP_FLAGS)
   End If


'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.SetLvw"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.SetLvw"
End Sub
Sub SetupcboAPPL()
    On Error GoTo errHandler

    cboAPPL.BeginUpdate
    cboAPPL.WidthList = 400
    cboAPPL.HeightList = 162
    cboAPPL.AllowSizeGrip = True
    cboAPPL.AutoDropDown = True
    cboAPPL.Columns.Add "Date"
    cboAPPL.Columns.Add "Doc.ref."
    cboAPPL.Columns.Add "Qty"
    cboAPPL.Columns.Add "Qty returned"
    cboAPPL.Columns.Add "Title"
    cboAPPL.Columns.Add "APPLID"
    cboAPPL.Columns.Add "Code"
    cboAPPL.Columns(0).Width = 70
    cboAPPL.Columns(1).Width = 70
    cboAPPL.Columns(2).Width = 40
    cboAPPL.Columns(3).Width = 40
    cboAPPL.Columns(4).Width = 100
    cboAPPL.Columns(5).Width = 0
    cboAPPL.Columns(6).Width = 0
    cboAPPL.BackColorLock = Me.BackColor
    cboAPPL.EndUpdate
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.SetupcboAPPL"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.SetupcboAPPL"
End Sub
Private Sub LoadcboAPPL(pLoadAll As Boolean)
    On Error GoTo errHandler
Dim i As Integer
Dim oD As d_APPL
Dim ar()
    cboAPPL.Items.RemoveAllItems
    i = 0
    If cAPPL.Count < 1 Then
        Exit Sub
    End If
    For Each oD In cAPPL
        If pLoadAll Then
            i = i + 1
        ElseIf Not InCurrentList(oD.CodeF) Then
            i = i + 1
        End If
    Next
    If i = 0 Then Exit Sub
    ReDim ar(6, i - 1)
    i = 0
    cboAPPL.BeginUpdate
    For Each oD In cAPPL
            If pLoadAll Or Not InCurrentList(oD.CodeF) Then
            ar(0, i) = oD.DOCCode
            ar(1, i) = oD.TRDateF
            ar(2, i) = oD.qty
            ar(3, i) = oD.QtyReturned
            ar(4, i) = oD.Title
            ar(5, i) = oD.APPLID
            ar(6, i) = oD.CodeF
            i = i + 1
        End If
    Next
    cboAPPL.PutItems ar
    cboAPPL.EndUpdate
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.LoadcboAPPL(pLoadAll)", pLoadAll
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.LoadcboAPPL(pLoadAll)", pLoadAll
End Sub

Private Function InCurrentList(pCode As String)
    On Error GoTo errHandler
    InCurrentList = Not lvwLines.FindItem(pCode) Is Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPR.InCurrentList(pCode)", pCODE
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPR.InCurrentList(pCODE)", pCode
End Function
