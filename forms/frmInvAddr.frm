VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmInvAddr 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Addresses and notes"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   Icon            =   "frmInvAddr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4065
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3225
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2640
      Left            =   285
      TabIndex        =   1
      Top             =   345
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4657
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   13882315
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Addresses"
      TabPicture(0)   =   "frmInvAddr.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cboBilling"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtBilling"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboDelivery"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDelivery"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "frmInvAddr.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtInvMemo"
      Tab(1).ControlCount=   1
      Begin VB.TextBox txtDelivery 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   1140
         Left            =   2775
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1350
         Width           =   2355
      End
      Begin VB.ComboBox cboDelivery 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2775
         TabIndex        =   6
         Top             =   870
         Width           =   2370
      End
      Begin VB.TextBox txtBilling 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   1140
         Left            =   195
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1350
         Width           =   2355
      End
      Begin VB.ComboBox cboBilling 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   195
         TabIndex        =   3
         Top             =   870
         Width           =   2370
      End
      Begin VB.TextBox txtInvMemo 
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
         Height          =   2040
         Left            =   -74910
         TabIndex        =   2
         Top             =   480
         Width           =   5085
      End
      Begin VB.Label Label2 
         Caption         =   "Delivery"
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
         Height          =   270
         Left            =   2760
         TabIndex        =   7
         Top             =   540
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Billing"
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
         Height          =   270
         Left            =   180
         TabIndex        =   4
         Top             =   540
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmInvAddr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oInv As a_Invoice
Dim flgLoading As Boolean

Public Sub component(pINV As a_Invoice)
    On Error GoTo errHandler
Dim oAddr As a_Address
    Set oInv = pINV
    flgLoading = True
    With oInv.Customer
        For Each oAddr In .Addresses
            Me.cboBilling.AddItem oAddr.Description
        Next
      '  cboBilling = IIf(oInv.BillToADdress.Description > "", oInv.BillToADdress.Description, .DefaultAddress.Description)
      '  txtBilling = IIf(oInv.BillToADdress.AddressMailing > "", oInv.BillToADdress.AddressMailing, .DefaultAddress.AddressMailing)
        cboBilling = oInv.BillTOAddress.Description
        txtBilling = oInv.BillTOAddress.AddressMailing
        
        For Each oAddr In .Addresses
            Me.cboDelivery.AddItem oAddr.Description
        Next
    '    cboDelivery = IIf(oInv.DelToAddress.Description > "", oInv.DelToAddress.Description, .DefaultAddress.Description)
    '    txtDelivery = IIf(oInv.DelToAddress.AddressMailing > "", oInv.DelToAddress.AddressMailing, .DefaultAddress.AddressMailing)
        cboDelivery = oInv.DelToAddress.Description
        txtDelivery = oInv.DelToAddress.AddressMailing
    End With
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvAddr.component(pINV)", pINV
End Sub

'Private Sub cboBilling_Change()
'    If flgLoading Then Exit Sub
'    Set oInv.InvDocAddress = oInv.customer.Addresses.FindByDescription(cboBilling)
'    Me.txtBilling = oInv.InvDocAddress.AddressMailing
'End Sub

Private Sub cboBilling_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oInv.SetBillToAddress oInv.Customer.Addresses.FindByDescription(cboBilling)
    Me.txtBilling = oInv.BillTOAddress.AddressMailing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvAddr.cboBilling_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cboDelivery_Change()
'    If flgLoading Then Exit Sub
'    Set oInv.InvGoodsAddress = oInv.customer.Addresses.FindByDescription(cboDelivery)
'    Me.txtDelivery = oInv.InvGoodsAddress.AddressMailing
'End Sub

Private Sub cboDelivery_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oInv.setDelToAddress oInv.Customer.Addresses.FindByDescription(cboDelivery)
    Me.txtDelivery = oInv.DelToAddress.AddressMailing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvAddr.cboDelivery_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvAddr.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        TOP = 50
        Left = 50
    End If
    txtInvMemo = oInv.Memo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvAddr.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    oInv.SetMemo txtInvMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvAddr.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub
