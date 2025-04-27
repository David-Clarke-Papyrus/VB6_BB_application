VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "COOLBU~1.OCX"
Begin VB.Form frmTF 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Transfer"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11130
   ControlBox      =   0   'False
   Icon            =   "frmTF.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   11130
   Begin VB.CommandButton cmdScannerOnly 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Use scanner only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5640
      Visible         =   0   'False
      Width           =   2310
   End
   Begin CoolButtonControl.CoolButton CBInOut 
      Height          =   930
      Left            =   255
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1640
      BackColor       =   14737632
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Transfer - IN"
      Style           =   6
      BackStyle       =   0
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
      Left            =   8610
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmTF.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5370
      UseMaskColor    =   -1  'True
      Width           =   1110
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
      Height          =   765
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5385
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
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
      Picture         =   "frmTF.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5370
      Width           =   1110
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00D3D3CB&
      Height          =   1290
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   10740
      Begin VB.TextBox txtStock 
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
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   870
         Width           =   4170
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4440
         TabIndex        =   6
         Top             =   480
         Width           =   1560
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2760
         TabIndex        =   5
         Top             =   480
         Width           =   1560
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdEnter 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Post"
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
         Left            =   9810
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
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
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   870
         Width           =   5730
      End
      Begin VB.TextBox txtCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   285
         TabIndex        =   3
         Top             =   480
         Width           =   1560
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Discount"
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
         Height          =   225
         Left            =   4485
         TabIndex        =   21
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Price"
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
         Height          =   225
         Left            =   2805
         TabIndex        =   19
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         Caption         =   "Qty"
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
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label9 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Code"
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
         Height          =   225
         Left            =   315
         TabIndex        =   15
         Top             =   270
         Width           =   1065
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
      Left            =   9510
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "IN PROCESS"
      Top             =   135
      Width           =   1260
   End
   Begin VB.CommandButton cmdIssue 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Issue"
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
      Left            =   9750
      Picture         =   "frmTF.frx":0A5E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5370
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin MSComctlLib.ListView lvwLines 
      Height          =   2250
      Left            =   135
      TabIndex        =   2
      Top             =   1215
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   3969
      SortKey         =   7
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483635
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title / Author / Publisher"
         Object.Width           =   7232
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Discount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Ext."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Excl .VAT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Key"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   0
      EndProperty
   End
   Begin CoolButtonControl.CoolButton cbDelTo 
      Height          =   465
      Left            =   3915
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   405
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   820
      BackColor       =   14737632
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      BackStyle       =   0
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
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
      Height          =   300
      Left            =   9435
      TabIndex        =   20
      Top             =   3660
      Width           =   1395
   End
   Begin VB.Label lblToFrom 
      Alignment       =   2  'Center
      BackColor       =   &H00D3D3CB&
      Caption         =   "FROM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   2865
      TabIndex        =   16
      Top             =   465
      Width           =   795
   End
End
Attribute VB_Name = "frmTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oTF As a_TF
Attribute oTF.VB_VarHelpID = -1
Dim WithEvents oTFLine As a_TFL
Attribute oTFLine.VB_VarHelpID = -1
Dim oInv As c_Invoices
Private tlSections As z_TextList

Dim strSelectedRowIndex As String
Dim lngEditingIdx As String
Dim vMode As EnumMode  ' 1:TPExists,Adding row;  2:TPExists, not adding row;  3 TPAbsent,not adding row
Dim bFrameEnabled As Boolean
Dim bValidTF As Boolean
Dim blnReadOnly As Boolean
Dim flgLoading As Boolean
Dim WithEvents vCanAdd As z_BrokenRules
Attribute vCanAdd.VB_VarHelpID = -1
Dim WithEvents vCanIssue As z_BrokenRules
Attribute vCanIssue.VB_VarHelpID = -1
Dim WithEvents oValidTF As z_BrokenRules
Attribute oValidTF.VB_VarHelpID = -1

Private Sub oTF_ValuesChange(pTotalExtF As String)
    On Error GoTo errHandler
    lblTotal.Caption = pTotalExtF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.oTF_ValuesChange(pTotalExtF)", pTotalExtF, EA_NORERAISE
    HandleError
End Sub

Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
    ofrm.component oTF.Memo
    ofrm.Show vbModal
    oTF.setMemo ofrm.Memo
    Unload ofrm
    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.mnuMemo"
End Sub

Private Sub cmdScannerOnly_Click()
    On Error GoTo errHandler
    If txtQty.Visible Then
        txtQty.Visible = False
        txtPrice.Visible = False
    Else
        txtQty.Visible = True
        txtPrice.Visible = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.cmdScannerOnly_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oTF.statusF = "IN PROCESS" And oTF.IsNew = False)
    Forms(0).mnuDelLine.Enabled = True
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.SetMenu"
End Sub


Private Sub CBInOut_Click()
    On Error GoTo errHandler
    Select Case CBInOut.Caption
    Case "Transfer - IN"
        CBInOut.Caption = "Transfer - OUT"
        lblToFrom.Caption = "TO"
        oTF.InOut = "OUT"
    Case "Transfer - OUT"
        CBInOut.Caption = "Transfer - IN"
        lblToFrom.Caption = "FROM"
        oTF.InOut = "IN"
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.CBInOut_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub lvwLines_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.lvwLines_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), EA_NORERAISE
    HandleError
End Sub
Private Sub lvwLines_Click()
    On Error GoTo errHandler
    Clipboard.Clear
    Clipboard.SetText left(lvwLines.SelectedItem.Text, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.lvwLines_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuDel_Click()
    On Error GoTo errHandler
    RemoveDetailLine
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.mnuDel_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_CN
    Set frm = New frmPrintingOptions_CN
    frm.Show vbModal

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.mnuPrint_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub oTF_Reloadlist()
    On Error GoTo errHandler
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.oTF_Reloadlist", , EA_NORERAISE
    HandleError
End Sub
Private Sub oTF_Dirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal = True Then
        Me.cmdSave.Enabled = (True And Not bFrameEnabled)
        Me.cmdCancel.Caption = "&Cancel"
    Else
        Me.cmdSave.Enabled = False
        Me.cmdCancel.Caption = "&Close"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.oTF_Dirty(pVal)", pVal, EA_NORERAISE
    HandleError
End Sub
Private Sub oTF_CurrRowStatus(pMsg As String)
    On Error GoTo errHandler
    MsgBox "CurrentRow Status = " & pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.oTF_CurrRowStatus(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub



Private Sub oTF_Valid(pMsg As String)
    On Error GoTo errHandler
    bValidTF = (pMsg = "")
    cmdIssue.Enabled = (bValidTF And oTF.TFLines.Count > 0 And vMode = enNotEditing)
    cmdSave.Enabled = (bValidTF And oTF.TFLines.Count > 0 And vMode = enNotEditing)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.oTF_Valid(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub

Sub vCanAdd_NobrokenRules()
    On Error GoTo errHandler
    Me.cmdNewRows.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.vCanAdd_NobrokenRules", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
Dim curTotalDeposit As Currency
    left = 10
    top = 10
    Width = 11100
    Height = 6700
    flgLoading = True
    SetLvw
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Initialize()
    On Error GoTo errHandler
    Set vCanAdd = New z_BrokenRules
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If oTF.IsEditing Then oTF.CancelEdit
    
    Set oTF = Nothing
    Set oTFLine = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Public Sub component(Optional pTF As a_TF)
    On Error GoTo errHandler
    flgLoading = True
    If pTF Is Nothing Then
        Set oTF = New a_TF
        oTF.BeginEdit
        oTF.SetStatus stInProcess
        Me.lvwLines.Enabled = False
        cmdNewRows.Caption = "&Stop"
        vMode = enAddingRow
        SetEditFrameEnabled True, vMode
        Set oTFLine = oTF.TFLines.Add
        oTFLine.SetQty 1
        ClearLineControls
        txtQty = 1
        fr1.ZOrder 0
        oTF.DestID = oPC.Configuration.DefaultStoreID
        cbDelTo_Click
        mSetfocus txtCode
    Else
        Set oTF = pTF
        oTF.BeginEdit
        LoadListView
        cmdSave.Enabled = False
        cmdIssue.Enabled = False
        cmdCancel.Caption = "&Close"
        cmdNewRows.Enabled = True
        Me.lvwLines.Enabled = True
        lvwLines.Height = 4100
        Me.cbDelTo.Caption = oPC.Configuration.Stores.FindStoreByID(oTF.DestID).Description
        If oTF.InOut = "IN" Then
            CBInOut.Caption = "Transfer - IN"
            lblToFrom.Caption = "FROM"
        ElseIf oTF.InOut = "OUT" Then
            CBInOut.Caption = "Transfer - OUT"
            lblToFrom.Caption = "TO"
        End If
        vMode = enNotEditing
        SetEditFrameEnabled False, enNotEditing
        fr1.ZOrder 1
        mSetfocus lvwLines
    End If
    flgLoading = False
    SetMenu
    oTF.GetStatus

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.Component(pTF)", pTF
End Sub
Private Sub SetEditFrameEnabled(pYesNo As Boolean, eMode As EnumMode)
    On Error GoTo errHandler
Dim lngColour As Long
    'A is adding, E is editing
    bFrameEnabled = pYesNo   'shared for use in all the form
    
    If (eMode = enAddingRow Or eMode = enNotEditing) And pYesNo Then
        Me.txtCode.Enabled = True
        Me.txtQty.Enabled = True
        Me.txtPrice.Enabled = True
    Else
        Me.txtCode.Enabled = False
        Me.txtQty.Enabled = True
        Me.txtPrice.Enabled = True
    End If
    Me.txtTitle.Enabled = pYesNo
    
    Me.cmdEnter.Enabled = pYesNo
    Me.cmdCancel.Enabled = Not pYesNo
    Me.cmdIssue.Enabled = (Not pYesNo) And bValidTF
    Me.cmdSave.Enabled = (Not pYesNo) And bValidTF And oTF.IsDirty
    
    If pYesNo Then
        lngColour = &HFFFFFF
    Else
        lngColour = 14416635
    End If
    
    Me.txtCode.BackColor = lngColour
    Me.txtQty.BackColor = lngColour
    Me.txtPrice.BackColor = lngColour
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.SetEditFrameEnabled(pYesNo,eMode)", Array(pYesNo, eMode)
End Sub
Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayoutLvw Me.lvwLines, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEnter_Click()
    On Error GoTo errHandler
Dim currDeposit As Currency
Dim blnResult As Boolean
Dim strCurrFormat As String
Dim curTotalDeposit As Currency
Dim strPos As String

    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Can't post."
        If txtCode.Enabled Then mSetfocus txtCode
        Exit Sub
    End If
    If txtQty = "" Or txtQty = "0" Then
        MsgBox "Enter a qty", vbOKOnly + vbInformation, "Can't post."
        If txtQty.Enabled Then mSetfocus txtQty
        Exit Sub
    End If
    If txtPrice = "" Or txtPrice = "0" Then
        MsgBox "Enter a price", vbOKOnly + vbInformation, "Can't post."
        If txtPrice.Enabled Then mSetfocus txtPrice
        Exit Sub
    End If
    oTFLine.ApplyEdit
    oTFLine.BeginEdit
    oTFLine.RecalculateLine

    If vMode = enAddingRow Then
        lvwLines.ListItems.Add Key:=oTFLine.Key
        LoadListViewLine Me.lvwLines.ListItems(lvwLines.ListItems.Count)
        Set oTFLine = Nothing
        lvwLines.Refresh
        Set oTFLine = oTF.TFLines.Add
        oTFLine.SetQty 1
        oTFLine.trid = oTF.trid
        mSetfocus txtCode
    ElseIf vMode = enEditingRow Then
        LoadListViewLine Me.lvwLines.ListItems(strSelectedRowIndex)
       ' LoadListViewLine strSelectedRowIndex, Me.lvwLines.ListItems(strSelectedRowIndex)
        lvwLines.Enabled = True
        lvwLines.Height = 4100
        vMode = enNotEditing
        SetEditFrameEnabled False, vMode
        cmdNewRows.Caption = "&Add"
        fr1.ZOrder 1
    End If
    
    ClearLineControls
    txtQty = 1
    oTF.CalculateTotal
    LoadListView
    oTF.GetStatus
    lblTotal.Caption = oTF.TotalPayableExTaxF
    lvwLines.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.cmdEnter_Click", , EA_NORERAISE, , strPos, Array(strPos)
    HandleError
End Sub
Private Sub LoadListViewLine(lstItem As ListItem)
    On Error GoTo errHandler
Dim currPrice As Currency
    With oTFLine
        lstItem.Text = .CodeF
        lstItem.SubItems(1) = .Title
        lstItem.SubItems(2) = .Qty
        lstItem.SubItems(3) = .PriceF
        lstItem.SubItems(4) = .DiscountF
        lstItem.SubItems(5) = .ExtF
        lstItem.SubItems(6) = .Ext_NetVATF
        lstItem.SubItems(7) = Format(.Key, "@@@@@@@@@@")
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.LoadListViewLine(lstItem)", Array(lstItem)
End Sub
Private Sub LoadListView()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Long
    For i = 1 To lvwLines.ColumnHeaders.Count
        lvwLines.ColumnHeaders(i).Width = GetSetting("PBKS", Me.Name, CStr(i), lvwLines.ColumnHeaders(i).Width)
    Next

    lvwLines.ListItems.Clear
    For i = 1 To oTF.TFLines.Count
        lvwLines.ListItems.Add Key:=oTF.TFLines(i).Key
        Set lstItem = lvwLines.ListItems(lvwLines.ListItems.Count)
        With oTF.TFLines(i)
            lstItem.Text = .CodeF
            lstItem.SubItems(1) = .Title
            lstItem.SubItems(2) = .Qty
            lstItem.SubItems(3) = .PriceF
            lstItem.SubItems(4) = .DiscountF
            lstItem.SubItems(5) = .ExtF
            lstItem.SubItems(6) = .Ext_NetVATF
            lstItem.SubItems(7) = Format(.Key, "@@@@@@@@@@")
        End With
    Next i
EXIT_HANDLER:
    Set lstItem = Nothing
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.LoadListView"
End Sub


Private Sub cmdNewRows_Click()
    On Error GoTo errHandler
Dim lr As Long
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
    
    If vMode = enEditingRow Then       'We have finished editing a row
        cmdNewRows.Caption = "&Add"
        vMode = enNotEditing
        SetEditFrameEnabled False, vMode
        lvwLines.Height = 4100
        fr1.ZOrder 1
        Me.lvwLines.Enabled = True
    ElseIf vMode = enAddingRow Then    'we are stopping adding rows
        vMode = enNotEditing
        cmdNewRows.Caption = "&Add"
        lvwLines.Height = 4100
        fr1.ZOrder 1
        SetEditFrameEnabled False, vMode
        lvwLines.Enabled = True
        oTF.GetStatus
    ElseIf vMode = enNotEditing Then  'we are starting to add rows
        vMode = enAddingRow
        Set oTFLine = oTF.TFLines.Add
        oTFLine.SetQty 1
        oTFLine.trid = oTF.trid
        cmdNewRows.Caption = "&Stop"
        Me.lvwLines.Enabled = False
        lvwLines.Height = 2400
        SetEditFrameEnabled True, vMode
        mSetfocus txtCode
    End If

    ClearLineControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.cmdNewRows_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub lvwLines_DblClick()
    On Error GoTo errHandler
'This must load the editing line with the current line's data
    If lvwLines.ListItems.Count = 0 Then Exit Sub
    If lvwLines.SelectedItem.Index < 1 Then Exit Sub
    lngEditingIdx = lvwLines.SelectedItem.Key
    Set oTFLine = oTF.TFLines(lngEditingIdx)
    strSelectedRowIndex = lvwLines.SelectedItem.Key
    
    txtTitle = oTFLine.Title
    txtQty = Abs(oTFLine.Qty)
    txtPrice = Abs(oTFLine.Price)
    txtCode = oTFLine.code
    
    SetEditFrameEnabled True, enEditingRow
    vMode = enEditingRow
    fr1.ZOrder 0
    lvwLines.Height = 2400
    cmdNewRows.Caption = "&Stop edit"
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.lvwLines_DblClick", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim pQty As Integer
Dim pApproID As Long
Dim bOK  As Boolean
Dim oPCode As New z_ProdCode

    
    If txtCode = "" Or vMode = enEditingRow Then Exit Sub
START:
    If txtCode = "" Then Exit Sub
    If Not (IsISBN13(txtCode) Or IsISBN10(txtCode) Or IsHashCode(txtCode) Or IsPrivateCode(txtCode)) Then
        MsgBox "This is an invalid code, retry.", vbInformation, "Warning"
        Cancel = True
        GoTo EXIT_HANDLER
    End If
    bOK = oTFLine.SetLineProduct("", txtCode)
    If bOK = False Then   'Book in database
        Dim frmAdHoc As frmAdHocProduct
        Set frmAdHoc = New frmAdHocProduct
        frmAdHoc.component txtCode
        frmAdHoc.Show vbModal
        txtCode = frmAdHoc.code
        Unload frmAdHoc
        Set frmAdHoc = Nothing
        Cancel = True
        GoTo START
    Else
        txtTitle = oTFLine.Title
        txtPrice = oTFLine.Price
        txtQty = oTFLine.Qty
        txtDiscount = oTFLine.Discount
        If oTFLine.Product.QtyOnHand > 0 Then
            txtStock = oTFLine.Product.QtyOnHandF & " @ " & oTFLine.Product.SPF
            txtStock.Visible = True
        Else
            txtStock.Visible = False
        End If
        
    End If
    oTFLine.GetStatus
    If txtQty.Visible = False Then
        cmdEnter_Click
    End If
EXIT_HANDLER:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub cboMatch_Click()
    On Error GoTo errHandler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.cboMatch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub RemoveDetailLine()
    On Error GoTo errHandler
Dim i As Integer
Dim iMax As Integer
    iMax = lvwLines.ListItems.Count
    For i = iMax To 1 Step -1
        If lvwLines.ListItems(i).Selected Then
            oTF.TFLines.Remove lvwLines.ListItems(i).Key
            Exit For
        End If
    Next i
    If i = 0 Then
        MsgBox "Select an item prior to deleting.", vbOKOnly + vbInformation, "Can't delete"
        Exit Sub
    End If
    lvwLines.ListItems.Remove i
    lvwLines.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.RemoveDetailLine"
End Sub


Private Sub SaveTF()
    On Error GoTo errHandler
    
    oTF.post
    
EXIT_HANDLER:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.SaveTF"
End Sub

Public Sub PrintTransfer()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoTFLines As Boolean
Dim blnHideVAT As Boolean
Dim iCurrency As Integer

    
    Me.MousePointer = vbHourglass
    oTF.Load oTF.trid
    blnDiscount = False ' TO BE REMOVED ON COMPLETION????
    
    If blnNoTFLines Then
        MsgBox "There are no records to print on this transfer.", vbOKOnly + vbInformation, "Can't print"
        GoTo EXIT_HANDLER
    End If
    
EXIT_HANDLER:
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.PrintTransfer"
End Sub
Private Sub cmdIssue_Click()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoTFLines As Boolean
Dim iCurrency As Integer
'Dim ViewOrPrint As PreviewPrint
Dim strResult As String
Dim frm As frmTFPreview

    If oPC.Configuration.Signtransactions = True Then
        If SecurityControl(enSECURITY_TFR_SIGN, , "Sign this transfer", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    Else
        If oTF.Status = stInProcess Then
            If MsgBox("Issue this transfer?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    WaitMsg "Issuing transfer  . . .", True, Me
    oTF.SetStatus stISSUED
    oTF.StaffID = gSTAFFID
    
    strResult = oTF.post
    oTF.CalculateTotal
    Set frm = New frmTFPreview
    frm.ComponentObject oTF
    frm.Show
    WaitMsg "", False, Me
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.cmdIssue_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdSave_Click()
    On Error GoTo errHandler
    oTF.SetStatus stInProcess
    SaveTF
    oTF.BeginEdit
    Set oTFLine = oTF.TFLines.Add
    cmdCancel.Caption = "&Close"
    cmdSave.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oTF.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub ClearLineControls()
    On Error GoTo errHandler
    flgLoading = True
    Me.txtCode = ""
    Me.txtTitle = ""
    Me.txtQty = ""
    txtDiscount = ""
    txtPrice = ""
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.ClearLineControls"
End Sub

Private Sub lvwLines_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.lvwLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtQty")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.txtQty_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oTFLine.SetQty(txtQty) Then
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtPrice")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.txtPrice_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oTFLine.SetPrice(txtPrice) Then
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtDiscount_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtDiscount")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.txtDiscount_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oTFLine.SetDiscount(txtDiscount) Then
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.txtDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub SetIssueButtonCaption()
    On Error GoTo errHandler
        If oTF.statusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oTF.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Caption = "Print"
        End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.SetIssueButtonCaption"
End Sub

Private Sub lvwLines_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    On Error GoTo errHandler
   ' When a ColumnHeader object is clicked, the ListView control is
   ' sorted by the subitems of that column.
   ' Set the SortKey to the Index of the ColumnHeader - 1
   lvwLines.SortKey = ColumnHeader.Index - 1
   ' Set Sorted to True to sort the list.
    If lvwLines.SortOrder = lvwAscending Then
        lvwLines.SortOrder = lvwDescending
    Else
        lvwLines.SortOrder = lvwAscending
    End If
   lvwLines.Sorted = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.lvwLines_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
    HandleError
End Sub
Private Sub SetLvw()
    On Error GoTo errHandler
Dim style As Long
Dim hHeader As Long
   
  'get the handle to the listview header
   hHeader = SendMessage(lvwLines.hwnd, LVM_GETHEADER, 0, ByVal 0&)
   
  'get the current style attributes for the header
   style = GetWindowLong(hHeader, GWL_STYLE)
   
  'modify the style by toggling the HDS_BUTTONS style
   style = style Xor HDS_BUTTONS
   
  'set the new style and redraw the listview
   If style Then
      Call SetWindowLong(hHeader, GWL_STYLE, style)
      Call SetWindowPos(lvwLines.hwnd, Me.hwnd, 0, 0, 0, 0, SWP_FLAGS)
   End If


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.SetLvw"
End Sub

Private Sub cbDelTo_Click()
    On Error GoTo errHandler
Static i As Long
    i = oPC.Configuration.OptionLoopStores(GetMax(i, 1), True)
    oTF.DestID = oPC.Configuration.Stores(i).ID
    oTF.DestinationName = oPC.Configuration.Stores(i).Description
    Me.cbDelTo.Caption = oPC.Configuration.Stores(i).Description
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTF.cbDelTo_Click", , EA_NORERAISE
    HandleError
End Sub
Public Sub mnuDelLine()
    On Error GoTo errHandler
    RemoveDetailLine
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.mnuDelLine"
End Sub
'Private Sub RemoveDetailLine()
'    On Error GoTo errHandler
'Dim i As Integer
'Dim iMax As Integer
'    iMax = lvwLines.ListItems.Count
'    For i = iMax To 1 Step -1
'        If lvwLines.ListItems(i).Selected Then
'            oTF.transferLines.Remove lvwLines.ListItems(i).Key
'            Exit For
'        End If
'    Next i
'    If i = 0 Then
'        MsgBox "Select an item prior to deleting.", vbOKOnly + vbInformation, "Can't do this"
'        Exit Sub
'    End If
'    lvwLines.ListItems.Remove i
'    lvwLines.Refresh
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.RemoveDetailLine"
'End Sub

