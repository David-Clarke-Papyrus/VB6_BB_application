VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmAPP 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Appro"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11595
   ControlBox      =   0   'False
   Icon            =   "frmAPP.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   11595
   Begin MSComctlLib.ListView lvwLines 
      Height          =   2430
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1200
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4286
      SortKey         =   6
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title / Author / Publisher"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ref"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Price"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Discount"
         Object.Width           =   1835
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
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
      Left            =   8595
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmAPP.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5760
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
      Height          =   1050
      Left            =   1065
      MultiLine       =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5670
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5790
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
      Left            =   7500
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmAPP.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1110
   End
   Begin VB.TextBox txtRunningTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   16
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
      Left            =   9720
      Picture         =   "frmAPP.frx":085E
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin CoolButtonControl.CoolButton cmdBill 
      Height          =   1095
      Left            =   7500
      TabIndex        =   23
      Top             =   45
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   1931
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
   Begin CoolButtonControl.CoolButton cmdSelectCustomer 
      Height          =   1080
      Left            =   4365
      TabIndex        =   24
      Top             =   75
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   1905
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
   Begin CoolButtonControl.CoolButton cbComp 
      Height          =   360
      Left            =   765
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   105
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   635
      BackColor       =   14737632
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      BackStyle       =   0
   End
   Begin VB.Frame fr1 
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
      Height          =   1995
      Left            =   105
      TabIndex        =   10
      Top             =   3660
      Width           =   10695
      Begin VB.TextBox txtRef 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4830
         TabIndex        =   6
         Top             =   1530
         Width           =   1875
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
         Height          =   840
         Left            =   75
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6840
         TabIndex        =   2
         Top             =   465
         Width           =   1000
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4830
         TabIndex        =   5
         Top             =   1140
         Width           =   4815
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8895
         TabIndex        =   4
         Top             =   465
         Width           =   735
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
         Picture         =   "frmAPP.frx":0BE8
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   495
         Width           =   5070
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7875
         TabIndex        =   3
         Top             =   465
         Width           =   990
      End
      Begin VB.TextBox txtCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   30
         TabIndex        =   1
         Top             =   465
         Width           =   1650
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         Caption         =   "Ref:"
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
         Height          =   300
         Left            =   4215
         TabIndex        =   33
         Top             =   1545
         Width           =   585
      End
      Begin VB.Label Label5 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Last sent to . . ."
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
         Height          =   255
         Left            =   75
         TabIndex        =   32
         Top             =   855
         Width           =   1350
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         Caption         =   "Note:"
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
         Height          =   300
         Left            =   4230
         TabIndex        =   22
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Qty"
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
         Height          =   300
         Left            =   7200
         TabIndex        =   21
         Top             =   195
         Width           =   360
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         Caption         =   "Disc."
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
         Height          =   300
         Left            =   8955
         TabIndex        =   14
         Top             =   195
         Width           =   585
      End
      Begin VB.Label Label9 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Code"
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
         Height          =   285
         Left            =   660
         TabIndex        =   13
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label6 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Price"
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
         Height          =   300
         Left            =   8085
         TabIndex        =   12
         Top             =   195
         Width           =   555
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Del."
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
      Left            =   3960
      TabIndex        =   31
      Top             =   105
      Width           =   300
   End
   Begin VB.Label lblFax 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   4515
      TabIndex        =   29
      Top             =   795
      Width           =   2100
   End
   Begin VB.Label Label1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "From"
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
      Left            =   165
      TabIndex        =   28
      Top             =   135
      Width           =   555
   End
   Begin VB.Label txtPhone 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4500
      TabIndex        =   26
      Top             =   510
      Width           =   2100
   End
   Begin VB.Label txtCustName 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   4530
      TabIndex        =   25
      Top             =   150
      Width           =   2100
   End
   Begin VB.Label lblb 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Bill"
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
      Left            =   7170
      TabIndex        =   20
      Top             =   60
      Width           =   300
   End
   Begin VB.Label lblAddBill 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   810
      Left            =   7545
      TabIndex        =   19
      Top             =   90
      Width           =   1920
   End
End
Attribute VB_Name = "frmAPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oAPP As a_APP
Attribute oAPP.VB_VarHelpID = -1
Dim WithEvents oAPPLine As a_APPL
Attribute oAPPLine.VB_VarHelpID = -1
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
Dim strAPPErrMsg As String
Dim strAPPLErrMsg As String
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
Dim WithEvents vCanAdd As z_BrokenRules
Attribute vCanAdd.VB_VarHelpID = -1
Dim WithEvents vCanIssue As z_BrokenRules
Attribute vCanIssue.VB_VarHelpID = -1
Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
    ofrm.component oAPP.Memo
    ofrm.Show vbModal
    oAPP.SetMemo ofrm.Memo
    Unload ofrm
    Set ofrm = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuMemo"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.mnuMemo"
End Sub
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayoutLvw Me.lvwLines, Me.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.mnuSaveLayout"
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oAPP.StatusF = "IN PROCESS" And Not oAPP.IsNew)
    Forms(0).mnuDelLine.Enabled = True
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.SetMenu"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.SetMenu"
End Sub



Public Sub component(Optional pAPP As a_APP, Optional pCustID As Long)
    On Error GoTo errHandler
    flgLoading = True
    If pAPP Is Nothing Then
        Set oAPP = New a_APP
        oAPP.BeginEdit
        oAPP.SetStatus stInProcess
        oAPP.COMPID = oPC.Configuration.DefaultCOMPID
        lvwLines.Enabled = False
        SetControlsForNew
        If pCustID > 0 Then
            LoadNewCustomer pCustID
        End If
        lvwLines.Height = 2200
        cmdNewRows.Caption = "&Stop"
        vMode = enAddingRow
        SetEditFrameEnabled True, vMode
        oAPP.GetStatus
        mSetfocus txtCode
        Set oAPPLine = oAPP.ApproLines.Add
        oAPPLine.SetQty 1
        Me.Caption = "Appro (new) for " & oAPP.Customer.Fullname
    Else
        Set oAPP = pAPP
        oAPP.BeginEdit
        LoadCustomer
        LoadListView
        cmdSave.Enabled = False
        cmdIssue.Enabled = False
        cmdCancel.Caption = "&Close"
        cmdNewRows.Enabled = True
        lvwLines.Enabled = True
        lvwLines.Height = 4050
        vMode = enNotEditing
        SetEditFrameEnabled False, vMode
        Me.Caption = "Appro (edit) for " & oAPP.Customer.Fullname
    End If
    flgLoading = False
    SetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.Component(pAPP,pCustID)", Array(pAPP, pCustID)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.component(pAPP,pCustID)", Array(pAPP, pCustID)
End Sub

Private Sub cmdSelectCustomer_Click()
    On Error GoTo errHandler
Dim frm As New frmCustomerPreview
    
    If oAPP.Customer.ID > 0 Then
        frm.component oAPP.Customer
        frm.Show
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.cmdSelectCustomer_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.cmdSelectCustomer_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdBill_Click()
    On Error GoTo errHandler
Static iBillIdx As Integer
Dim i As Integer
START:
    If oAPP Is Nothing Then Exit Sub
    i = iBillIdx + 1
    If i > oAPP.Customer.Addresses.Count Then
        i = 1
    End If
    If oAPP.Customer.Addresses.Count > 0 Then
        Me.lblAddBill.Caption = oAPP.Customer.Addresses(i).AddressMailing & vbCrLf & oAPP.Customer.Addresses(i).EMail
        oAPP.SetApproToAddress oAPP.Customer.Addresses(i)
    End If
    iBillIdx = i
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.cmdBill_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.cmdBill_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.Form_Activate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.Form_Activate", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.Form_Deactivate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwLines_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
    Cancel = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.lvwLines_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
'         EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.lvwLines_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub
Private Sub lvwLines_Click()
    On Error GoTo errHandler
    If lvwLines Is Nothing Then Exit Sub
    If lvwLines.SelectedItem Is Nothing Then Exit Sub
    If lvwLines.SelectedItem.Index > 0 Then
        On Error Resume Next
        Clipboard.Clear
        Clipboard.SetText Left(lvwLines.SelectedItem.SubItems(7), ISBNLENGTH)
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.lvwLines_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.lvwLines_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuAddresses_Click()
    On Error GoTo errHandler
Dim frm As frmInvAddr
    Set frm = New frmInvAddr
    frm.component oAPP
    frm.Show vbModal
    lblAddBill.Caption = oAPP.ApproToAddress.AddressShort
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuAddresses_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.mnuAddresses_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuChangeCustomer_Click()
    On Error GoTo errHandler
Dim lngTPID As Long
Dim frm As frmBrowseCustomers2
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    If oAPP.SetCustomer(lngTPID) Then
        With oAPP.Customer
            txtPhone = .Phone
            txtCustName = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
            lblAddBill.Caption = .BillTOAddress.AddressShort
        End With
        vCanAdd.RuleBroken "TP", False
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuChangeCustomer_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.mnuChangeCustomer_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuDelLine()
    On Error GoTo errHandler
    RemoveDetailLine
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuDel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.mnuDelLine"
End Sub


Private Sub mnuPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_APP
    Set frm = New frmPrintingOptions_APP
    frm.Show vbModal

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuPrint_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.mnuPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oAPP_Valid(pMsg As String)
    On Error GoTo errHandler
    bValidAPP = (pMsg = "")
    cmdIssue.Enabled = (bValidAPP And oAPP.ApproLines.Count > 0 And vMode = enNotEditing)
    cmdSave.Enabled = (bValidAPP And oAPP.ApproLines.Count > 0 And vMode = enNotEditing)
    strAPPErrMsg = pMsg
    If vMode = enNotEditing Then
        txtError = strAPPErrMsg
    Else
        txtError = strAPPLErrMsg
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.oAPP_Valid(pMsg)", pMsg, EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.oAPP_Valid(pMsg)", pMsg, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.oAPP_Valid(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub

Sub oAPPLine_ExtensionChange(lngExtension As Long, strExtension As String)
    On Error GoTo errHandler
    flgLoading = True
   ' Me.txtTotal = strExtension
    flgLoading = False
    lngCurrentExtension = lngExtension
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.oAPPLine_ExtensionChange(lngExtension,strExtension)", Array(lngExtension, _
'         strExtension), EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.oAPPLine_ExtensionChange(lngExtension,strExtension)", Array(lngExtension, _
'         strExtension), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.oAPPLine_ExtensionChange(lngExtension,strExtension)", Array(lngExtension, _
         strExtension), EA_NORERAISE
    HandleError
End Sub

Private Sub oAPPLine_Valid(msg As String)
    On Error GoTo errHandler
    cmdEnter.Enabled = (msg = "")
    strAPPLErrMsg = msg
    If vMode = enNotEditing Then
        txtError = strAPPErrMsg
    Else
        txtError = strAPPLErrMsg
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.oAPPLine_Valid(Msg)", Msg, EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.oAPPLine_Valid(Msg)", msg, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.oAPPLine_Valid(msg)", msg, EA_NORERAISE
    HandleError
End Sub

Private Sub oAPP_TotalChange(lngTotal As Long, strtotal As String, lngTotalDeposit As Long, strTotalDeposit As String, lngTotalVAT As Long, strTotalVAT As String)
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
'    ErrorIn "frmAPP.oAPP_TotalChange(lngTotal,strtotal,lngTotalDeposit,strTotalDeposit,lngTotalVAT," & _
'        "strTotalVAT)", Array(lngTotal, strtotal, lngTotalDeposit, strTotalDeposit, lngTotalVAT, strTotalVAT), _
'         EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.oAPP_TotalChange(lngTotal,strtotal,lngTotalDeposit,strTotalDeposit,lngTotalVAT," & _
'        "strTotalVAT)", Array(lngTotal, strtotal, lngTotalDeposit, strTotalDeposit, lngTotalVAT, strTotalVAT), _
'         EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.oAPP_TotalChange(lngTotal,strtotal,lngTotalDeposit,strTotalDeposit,lngTotalVAT," & _
        "strTotalVAT)", Array(lngTotal, strtotal, lngTotalDeposit, strTotalDeposit, lngTotalVAT, strTotalVAT), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub oAPP_Reloadlist()
    On Error GoTo errHandler
    LoadListView
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.oAPP_Reloadlist", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.oAPP_Reloadlist", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.oAPP_Reloadlist", , EA_NORERAISE
    HandleError
End Sub
Private Sub oAPP_Dirty(pVal As Boolean)
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
'    ErrorIn "frmAPP.oAPP_Dirty(pVal)", pVal, EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.oAPP_Dirty(pVal)", pVal, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.oAPP_Dirty(pVal)", pVal, EA_NORERAISE
    HandleError
End Sub
Private Sub oAPP_CurrRowStatus(pMsg As String)
    On Error GoTo errHandler
  '  MsgBox "CurrentRow Status = " & pMsg
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.oAPP_CurrRowStatus(pMsg)", pMsg, EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.oAPP_CurrRowStatus(pMsg)", pMsg, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.oAPP_CurrRowStatus(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub



Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    txtNote = HandleTextWithBites(txtNote)
    oAPPLine.Note = txtNote
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtRef_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    txtRef = HandleTextWithBites(txtRef)
    oAPPLine.Ref = txtRef
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtRef_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.txtRef_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Sub vCanAdd_NobrokenRules()
    On Error GoTo errHandler
    Me.cmdNewRows.Enabled = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.vCanAdd_NobrokenRules", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.vCanAdd_NobrokenRules", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
Dim strAddress As String
    If Me.WindowState <> 2 Then
       Left = 10
        top = 10
        Width = 11100
        Height = 7000
    End If
  '  SetLvw
    
    flgLoading = True
    oAPP.GetStatus
    SetEditFrameEnabled False, enNotEditing
    vMode = enNotEditing
    LoadComps
    If oAPP.APPROTOID > 0 Then
        strAddress = oAPP.ApproToAddress.AddressMailing
    End If
    Me.lblAddBill.Caption = IIf(strAddress > "", strAddress, "unknown")
    oAPP.GetStatus
    
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadComps()
    On Error GoTo errHandler
Dim oComp As a_Company
Dim oItem As ListItem
Dim i As Integer
    If oAPP.COMPID > 0 Then
        cbComp.Caption = oPC.Configuration.Companies(CStr(oAPP.COMPID)).CompanyName
    Else
        cbComp.Caption = oPC.Configuration.DefaultCompany.CompanyName
        oAPP.COMPID = oPC.Configuration.DefaultCOMPID
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.LoadComps"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.LoadComps"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.LoadComps"
End Sub
Private Sub Form_Initialize()
    On Error GoTo errHandler
    Set vCanAdd = New z_BrokenRules
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.Form_Initialize", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.Form_Initialize", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If oAPP.IsEditing Then oAPP.CancelEdit
    
    Set oCustomer = Nothing
    Set oCurrentCopy = Nothing
    Set oAPP = Nothing
    Set tlCustomer = Nothing
    Set oAPPLine = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.Form_Unload(Cancel)", Cancel, EA_NORERAISE
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
    Me.txtNote.Enabled = pYesNo
    Me.txtRef.Enabled = pYesNo
    Me.txtDiscount.Enabled = pYesNo
    Me.txtPrice.Enabled = pYesNo
    Me.txtTitle.Enabled = pYesNo
    Me.txtQty.Enabled = pYesNo
    Me.cmdEnter.Enabled = pYesNo
    Me.cmdCancel.Enabled = Not pYesNo
    Me.cmdIssue.Enabled = (Not pYesNo) And bValidAPP
    Me.cmdSave.Enabled = (Not pYesNo) And bValidAPP And oAPP.IsDirty
    
    If pYesNo Then
        lngColour = &HFFFFFF
    Else
        lngColour = 14416635
    End If
    
    Me.txtCode.BackColor = lngColour
    Me.txtPrice.BackColor = lngColour
    Me.txtDiscount.BackColor = lngColour
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.SetEditFrameEnabled(pYesNo,eMode)", Array(pYesNo, eMode)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.SetEditFrameEnabled(pYesNo,eMode)", Array(pYesNo, eMode)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.SetEditFrameEnabled(pYesNo,eMode)", Array(pYesNo, eMode)
End Sub
Private Sub SetControlsForNew()
    On Error GoTo errHandler
    txtPhone = ""
    lblFax = ""
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.SetControlsForNew"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.SetControlsForNew"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.SetControlsForNew"
End Sub

Private Sub cmdEnter_Click()
    On Error GoTo errHandler
Dim currDeposit As Currency
Dim blnResult As Boolean
Dim strCurrFormat As String
Dim curTotalDeposit As Currency
    
    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Papyrus appro Information"
        If txtCode.Enabled = True Then mSetfocus txtCode
        Exit Sub
    End If
    oAPPLine.ApplyEdit
    oAPPLine.BeginEdit

    If vMode = enAddingRow Then
        lvwLines.ListItems.Add Key:=oAPPLine.Key
        LoadListViewLine lvwLines.ListItems(lvwLines.ListItems.Count), oAPPLine   'oAPPLine.Key, Me.lvwLines.ListItems(1)
        lvwLines.Refresh
        ChangeState enAddingRow
'        Set oAPPLine = Nothing
'        Set oAPPLine = oAPP.ApproLines.Add
'        oAPPLine.SetQty 1
'        oAPPLine.trid = oAPP.trid
'        ClearLineControls
        mSetfocus txtCode
    ElseIf vMode = eneditingrow Then
        LoadListViewLine lvwLines.ListItems(lngSelectedRowIndex), oAPPLine             'lngSelectedRowIndex, Me.lvwLines.ListItems(lngSelectedRowIndex)
        ChangeState enNotEditing
'        ClearLineControls
'        vMode = enNotEditing
'        SetEditFrameEnabled False, vMode
'        lvwLines.Height = 4120
'        cmdNewRows.Caption = "&Add"
'        fr1.ZOrder 1
    End If
    
 '   oAPP.CalculateTotal
 '   LoadListView
    oAPP.GetStatus
 '   lvwLines.Enabled = True
'errHandler:
'    ErrPreserve
'
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.cmdEnter_Click", , EA_NORERAISE, , "oAPPLine.Key, vMode,lvwLines.ListItems.count", Array(oAPPLine.key, vMode, lvwLines.ListItems)
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.cmdEnter_Click", , EA_NORERAISE
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
        txtNote.Enabled = True
        txtRef.Enabled = True
        txtDiscount.Enabled = True
        txtPrice.Enabled = True
        txtTitle.Enabled = True
       ' txtTotal.Enabled = True
        txtQty.Enabled = True
       ' cboRef.Visible = False
        cmdEnter.Enabled = False
        cmdCancel.Enabled = False
        cmdIssue.Enabled = False
        cmdSave.Enabled = False
        cmdNewRows.Caption = "&Stop"
        cmdNewRows.Enabled = (oAPP.ApproLines.Count > 0)
        cmdCancel.Caption = "&Close"
        lvwLines.Enabled = False
        lvwLines.Height = 2200
        fr1.ZOrder 1
        
    Case enAddingRow
        fr1.Visible = True
        txtCode.Enabled = True
        txtNote.Enabled = True
        txtRef.Enabled = True
        txtDiscount.Enabled = True
        txtPrice.Enabled = True
        txtTitle.Enabled = True
      '  txtTotal.Enabled = True
        txtQty.Enabled = True
        txtError = ""
        cmdEnter.Enabled = False
        cmdCancel.Enabled = True
        cmdIssue.Enabled = False
        cmdSave.Enabled = False
        cmdNewRows.Enabled = (oAPP.ApproLines.Count > 0)
        cmdNewRows.Caption = "&Stop"
        
     '   lblTPPhone.Caption = ""
        lvwLines.Enabled = False
        lvwLines.Height = 2200
        ClearLineControls
        fr1.ZOrder 1
        
        mSetfocus txtCode
        
        Set oAPPLine = oAPP.ApproLines.Add
        oAPPLine.TRID = oAPP.TRID
        oAPPLine.SetQty 1
        
    Case enNotEditing
'        cmdNewRows.Caption = "&Add"
        flgLoading = True
        fr1.Visible = False
        txtError = ""
    '   txtRef = ""
        flgLoading = False
        cmdEnter.Enabled = False
        cmdCancel.Enabled = True
        cmdIssue.Enabled = True
        cmdSave.Enabled = True
        cmdNewRows.Enabled = (oAPP.ApproLines.Count > 0)
        cmdNewRows.Caption = "&Add"
        
        lvwLines.Enabled = True
        lvwLines.Height = 4000
        
        fr1.ZOrder 1
    
    
    End Select
'    lblAppro.Caption = ""
'    cboRef.Visible = False

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.ChangeState(pToMode)", pToMode
End Sub


Private Sub LoadNewCustomer(plngTPID As Long)
    On Error GoTo errHandler
    If oAPP.SetCustomer(plngTPID) Then
        With oAPP.Customer
            txtPhone = .Phone
            txtCustName = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
            oAPP.TPID = plngTPID
            If Not .ApproAddress Is Nothing Then
                oAPP.SetApproToAddress .ApproAddress
                lblAddBill.Caption = .ApproAddress.AddressShort
            End If
        End With
        vCanAdd.RuleBroken "TP", False
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.LoadNewCustomer(plngTPID)", plngTPID
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.LoadNewCustomer(plngTPID)", plngTPID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.LoadNewCustomer(plngTPID)", plngTPID
End Sub


Private Sub cmdNewRows_Click()
    On Error GoTo errHandler
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
    
    If vMode = eneditingrow Then
        ChangeState enNotEditing
    ElseIf vMode = enAddingRow Then
        ChangeState enNotEditing
    ElseIf vMode = enNotEditing Then
        ChangeState enAddingRow
    End If
    ClearLineControls
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.cmdNewRows_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.cmdNewRows_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadListView()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Long
    For i = 1 To 6
        lvwLines.ColumnHeaders(i).Width = GetSetting("PBKS", Me.Name, CStr(i), lvwLines.ColumnHeaders(i).Width)
    Next
    lvwLines.ListItems.Clear
    
    For i = 1 To oAPP.ApproLines.Count
        Set lstItem = lvwLines.ListItems.Add
        LoadListViewLine lstItem, oAPP.ApproLines(i)
    Next i
    
EXIT_Handler:
    Set lstItem = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.LoadListView"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.LoadListView"
End Sub
Private Sub LoadListViewLine(lstItem As ListItem, oLine As a_APPL)
    On Error GoTo errHandler
Dim currPrice As Currency
    
    With oLine
        lstItem.text = .CodeF
        lstItem.Key = .Key
        lstItem.SubItems(1) = .Title
        lstItem.SubItems(2) = .Qty
        lstItem.SubItems(3) = .Ref
        lstItem.SubItems(4) = .PriceF
        lstItem.SubItems(5) = .DiscountF
        lstItem.SubItems(6) = .ExtensionNetF
        lstItem.SubItems(7) = Format(.Key, "@@@@@@@@@@")
        lstItem.SubItems(8) = .EAN
    End With
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.LoadListViewLine(lstItem,oLine)", Array(lstItem, oLine)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.LoadListViewLine(lstItem,oLine)", Array(lstItem, oLine)
End Sub
Private Sub lvwLines_DblClick()
    On Error GoTo errHandler
'This must load the editing line with the current line's data
    If lvwLines.ListItems.Count = 0 Then Exit Sub
    If lvwLines.SelectedItem.Index < 1 Then Exit Sub
    lngEditingIdx = lvwLines.SelectedItem.Key
    Set oAPPLine = oAPP.ApproLines(lngEditingIdx)
    
    lngSelectedRowIndex = lvwLines.SelectedItem.Key
    Me.txtCode = CStr(oAPPLine.EAN)
    Me.txtTitle = oAPPLine.Title
    Me.txtPrice = CStr(oAPPLine.Price)
    Me.txtDiscount = CStr(oAPPLine.Discount)
    Me.txtQty = oAPPLine.Qty
    txtNote = oAPPLine.Note
    txtRef = oAPPLine.Ref
    SetEditFrameEnabled True, eneditingrow
    fr1.Visible = True
    vMode = eneditingrow
    mSetfocus txtPrice
    lvwLines.Height = 2500
    cmdNewRows.Caption = "&Stop edit"
    oAPPLine.GetStatus
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.lvwLines_DblClick", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.lvwLines_DblClick", , EA_NORERAISE
    HandleError
End Sub

'---------Companies code
'Private Sub LoadComps()
'Dim oComp As a_APPmpany
'Dim oItem As ListItem
'Dim i As Integer
'    If oAPP.CompanyID > 0 Then
'        txtComp = oPC.Configuration.Companies(CStr(oAPP.CompanyID)).CompanyName
'    Else
'        txtComp = oPC.Configuration.DefaultCompany.CompanyName
'        oAPP.CompanyID = oPC.Configuration.DefaultCompanyID
'    End If
'End Sub

'Private Sub txtOrdernum_Validate(Cancel As Boolean)
'Dim intPos As Integer
'    If flgLoading Then Exit Sub
'    On Error Resume Next
'    oAPPLine.Ref = txtOrdernum
'    If Err Then
'      Beep
'      intPos = txtOrdernum.SelStart
'      txtOrdernum = oAPPLine.Ref
'      txtOrdernum.SelStart = intPos - 1
'    End If
'
'End Sub

'Private Sub txtNote_Change()
'Dim intPos As Integer
'    If flgLoading Then Exit Sub
'    On Error Resume Next
'    oAPPLine.setnote (txtNote)
'    If Err Then
'      Beep
'      intPos = txtNote.SelStart
'      txtNote = oAPPLine.Note
'      txtNote.SelStart = intPos - 1
'    End If
'End Sub
'Private Sub txtNote_Validate(Cancel As Boolean)
'    Cancel = Not oAPPLine.setnote(txtNote)
'End Sub
'Private Sub txtNote_LostFocus()
'    If flgLoading Then Exit Sub
'    txtNote = oAPPLine.Note
'End Sub

Private Sub mnuEditNote_Click()
    On Error GoTo errHandler
Dim ofrm As New frmNote
    ofrm.component oAPP
    ofrm.Show vbModal
    Unload ofrm
    Set ofrm = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuEditNote_Click", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuEditNote_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.mnuEditNote_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuFileCancel_Click()
    On Error GoTo errHandler
    If oAPP.IsDirty Then
        oAPP.CancelEdit
    End If
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuFileCancel_Click", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuFileCancel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.mnuFileCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuFileExit_Click()
    On Error GoTo errHandler
    oAPP.CancelEdit
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuFileExit_Click", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuFileExit_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.mnuFileExit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuFileOK_Click()
    On Error GoTo errHandler
'    cmdOK_Click
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuFileOK_Click", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuFileOK_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.mnuFileOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuFilePrint_Click()
    On Error GoTo errHandler
    cmdIssue_Click
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuFilePrint_Click", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuFilePrint_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.mnuFilePrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuFileVoid_Click()
    On Error GoTo errHandler
Dim strResult As String
Dim bMadeChanges As Boolean

    oAPP.SetStatus stVOID
    oAPP.ApplyEdit bMadeChanges
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuFileVoid_Click", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuFileVoid_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.mnuFileVoid_Click", , EA_NORERAISE
    HandleError
End Sub
'Private Sub txtAccNum_Validate(Cancel As Boolean)
'Dim lngCustID As Long
'Dim bResult As Boolean
'    If Len(txtAccnum) > 0 Then
'        bResult = oAPP.SetCustomerFromAccNum(txtAccnum)
'        If bResult Then
'            With oAPP.Customer
'                txtCustName = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
'                txtPhone = .Phone
'                lblAddBill.Caption = .BillToADdress.AddressShort
'                lblAddDel.Caption = .BillToADdress.AddressShort
'            End With
'            vCanAdd.RuleBroken "TP", False
'        Else
'            MsgBox "No such account number", , "Can't fetch customer"
'            txtAccnum = ""
'            Set oCustomer = Nothing
'            Cancel = True
'        End If
'    End If
'End Sub
'Private Sub txtComp_DblClick()
'Dim iCompIdx As Integer
'Dim i As Integer
'Start:
'    i = iCompIdx + 1
'    If i > oPC.Configuration.Companies.Count Then
'        i = 1
'    End If
'    If lngCompanyID = oPC.Configuration.Companies(i).ID Then
'        GoTo Start
'    End If
'    txtComp = oPC.Configuration.Companies(i).CompanyName
'    oAPP.CompanyID = oPC.Configuration.Companies(i).ID
'    iCompIdx = i
'End Sub
Private Sub cbComp_Click()
    On Error GoTo errHandler
    oAPP.COMPID = OptionLoop(oAPP.COMPID, oPC.Configuration.Companies.Count)
    cbComp.Caption = oPC.Configuration.Companies(oAPP.COMPID).CompanyName
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.cbComp_Click", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.cbComp_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.cbComp_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim pQty As Integer
Dim pApproID As Long
Dim bOK  As Boolean

    txtCode = FNS(txtCode)
    If txtCode = "" Or vMode = eneditingrow Then Exit Sub
    If Not (IsISBN13(txtCode) Or IsISBN10(txtCode) Or IsHashCode(txtCode) Or IsPrivateCode(txtCode)) Then
        MsgBox "This is an invalid code, retry.", vbInformation, "Warning"
        Cancel = True
        GoTo EXIT_Handler
    End If
    
    bOK = oAPPLine.SetLineProduct("", txtCode)
    If bOK Then
        txtTitle = oAPPLine.Title
        txtPrice = oAPPLine.Price
        txtQty = oAPPLine.Qty
        txtDiscount = oAPPLine.Discount
        mSetfocus txtPrice
        txtCode = oAPPLine.EAN
        txtLastAt = oAPPLine.LastApproto
        AutoSelect txtPrice
    Else
        MsgBox "Cannot find book on database.", vbOKOnly + vbInformation, "Error"
        Cancel = True
        GoTo EXIT_Handler
    End If

EXIT_Handler:
    Set oProd = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub RemoveDetailLine()
    On Error GoTo errHandler
Dim i As Integer
Dim iMax As Integer
    iMax = lvwLines.ListItems.Count
    For i = iMax To 1 Step -1
        If lvwLines.ListItems(i).Selected Then
            oAPP.ApproLines.Remove lvwLines.ListItems(i).Key
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
'    ErrorIn "frmAPP.RemoveDetailLine"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.RemoveDetailLine"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.RemoveDetailLine"
End Sub

Private Sub LoadCustomer()
    On Error GoTo errHandler
Dim strAddress As String
    With oAPP
        SetIssueButtonCaption
        txtCustName = .Customer.Fullname
        If Not .Customer.BillTOAddress Is Nothing Then
            txtPhone = .Customer.BillTOAddress.Phone
            lblFax = .Customer.BillTOAddress.Fax
        End If
        If oAPP.APPROTOID > 0 Then
            strAddress = oAPP.ApproToAddress.AddressMailing
        End If
        Me.lblAddBill.Caption = IIf(strAddress > "", strAddress, "unknown")
    End With
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.LoadCustomer"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.LoadCustomer"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.LoadCustomer"
End Sub


Private Sub SaveCO()
    On Error GoTo errHandler
    
    oAPP.Post
    
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
  '  Resume
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.SaveCO"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.SaveCO"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.SaveCO"
End Sub

Public Sub PrintAppro()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoCOLines As Boolean
Dim blnHideVAT As Boolean
Dim iCurrency As Integer

    
    Me.MousePointer = vbHourglass
    oAPP.Load oAPP.TRID, False
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
'    ErrorIn "frmAPP.PrintAppro"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.PrintAppro"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.PrintAppro"
End Sub
Private Sub cmdIssue_Click()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoCOLines As Boolean
Dim iCurrency As Integer
Dim strResult As String
Dim frm As frmAPPPreview

    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_APP_SIGN, , "Sign this appro", DOCAPPROVAL, , , gSTAFFID) = False Then
               Exit Sub
        End If
    Else
        If oAPP.Status = stInProcess Then
            If MsgBox("Issue this appro?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
    End If


    WaitMsg "Issuing appro  . . .", True, Me
    oAPP.SetStatus stISSUED
    oAPP.StaffID = gSTAFFID
    
    oAPP.Post
    Set frm = New frmAPPPreview
    frm.ComponentObject oAPP
    frm.Show
    WaitMsg "", False, Me
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.cmdIssue_Click", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.cmdIssue_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.cmdIssue_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdSave_Click()
    On Error GoTo errHandler
Dim bMadeChanges As Boolean
Dim strPos As String

strPos = "1"
    oAPP.SetStatus stInProcess
strPos = "2"
    oAPP.ApplyEdit bMadeChanges
    oAPP.Reload
strPos = "3"
    LoadListView
strPos = "4"
    oAPP.BeginEdit
strPos = "5"
    Set oAPPLine = oAPP.ApproLines.Add
strPos = "6"
strPos = "7"
    cmdCancel.Caption = "&Close"
strPos = "8"
    cmdSave.Enabled = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.cmdSave_Click", , EA_NORERAISE, , strPos, Array(strPos)
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
Dim frm As frmAPPPreview
    If cmdCancel.Caption <> "&Close" Then
        If MsgBox("You wish to cancel this appro?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
            Exit Sub
        End If
    End If
    oAPP.CancelEdit
  '  LoadListView
    If cmdCancel.Caption = "&Close" Then
        Set frm = New frmAPPPreview
        frm.ComponentObject oAPP
        frm.Show
    End If
    
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.cmdCancel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub ClearLineControls()
    On Error GoTo errHandler
    flgLoading = True
    Me.txtCode = ""
    Me.txtDiscount = ""
    Me.txtPrice = ""
    Me.txtTitle = ""
    Me.txtNote = ""
    Me.txtRef = ""
    Me.txtQty = ""
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.ClearLineControls"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.ClearLineControls"
End Sub

Private Sub lvwLines_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.lvwLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.lvwLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtPrice")
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtPrice_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.txtPrice_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oAPPLine.SetPrice(txtPrice) Then
        Cancel = True
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_LostFocus()
    On Error GoTo errHandler
  '  txtPrice = oAPPLine.PriceF
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtPrice_LostFocus", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtPrice_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.txtPrice_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtQty")
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtQty_GotFocus", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtQty_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.txtQty_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oAPPLine.SetQty(txtQty) Then
        Cancel = True
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_LostFocus()
    On Error GoTo errHandler
  '  txtQty = oAPPLine.QtyF
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtQty_LostFocus", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtQty_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.txtQty_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oAPPLine.SetDiscount(txtDiscount) Then
        Cancel = True
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.txtDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtDiscount_LostFocus()
    On Error GoTo errHandler
  '  txtDiscount = oAPPLine.DiscountPercentF
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtDiscount_LostFocus", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtDiscount_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.txtDiscount_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDiscount_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtDiscount")
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtDiscount_GotFocus", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.txtDiscount_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.txtDiscount_GotFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub SetIssueButtonCaption()
    On Error GoTo errHandler
        If oAPP.StatusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oAPP.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Caption = "Print"
        End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.SetIssueButtonCaption"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.SetIssueButtonCaption"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.SetIssueButtonCaption"
End Sub
'Private Sub txtAccNum_LostFocus()
'    txtAccnum = UCase(txtAccnum)
'End Sub


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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.lvwLines_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.lvwLines_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.lvwLines_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
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
'    ErrorIn "frmAPP.SetLvw"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.SetLvw"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.SetLvw"
End Sub

