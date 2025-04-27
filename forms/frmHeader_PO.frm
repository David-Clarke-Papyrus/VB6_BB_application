VERSION 5.00
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmHeader_PO 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Purchase order header"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   5400
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5400
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6195
      Width           =   135
   End
   Begin VB.CommandButton cmdClose 
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
      Left            =   4155
      Picture         =   "frmHeader_PO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1665
      Width           =   1000
   End
   Begin VB.TextBox txtMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   75
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1470
      Width           =   4005
   End
   Begin CoolButtonControl.CoolButton cbDelTo 
      Height          =   1305
      Left            =   1095
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   45
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   2302
      BackColor       =   14737632
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin VB.Label lblDelTo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1020
      Left            =   1290
      TabIndex        =   7
      Top             =   240
      Width           =   2805
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deliver to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   105
      TabIndex        =   6
      Top             =   45
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "(Click ESC to cancel)"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   30
      TabIndex        =   4
      Top             =   2415
      Width           =   1800
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Memo"
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
      Height          =   240
      Left            =   150
      TabIndex        =   2
      Top             =   1245
      Width           =   1965
   End
End
Attribute VB_Name = "frmHeader_PO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strOrderNumber As String
Dim strOrderDate As String
Dim strMemo As String
Dim flgLoading As Boolean
Dim oPO As a_PO
Dim bCancel As Boolean
Dim Blocked As Boolean
Dim iOpt As Long

Public Property Get Cancelled() As Boolean
    Cancelled = bCancel
End Property

Public Sub component(pLocked As Boolean, pMemo As String, pTRID As Long)
    On Error GoTo errHandler
    Set oPO = New a_PO
    oPO.Load pTRID, False
    oPO.BeginEdit
'    Me.lblOrderNumber.Caption = pCaptionOrdernum
'    Me.lblOrderDate.Caption = pCaptionOrderDate
'    Me.txtOrderNumber = pOrderNumber
'    Me.txtOrderDate = pOrderDate
    Me.txtMemo = pMemo
    Blocked = pLocked
    If Me.WindowState <> 2 Then
        TOP = 3000
        Left = 1000
    End If
'    txtOrderNumber.Locked = pLocked
'    txtOrderDate.Locked = pLocked
    txtMemo.Locked = pLocked
    If pLocked Then
'        txtOrderNumber.BackColor = &HDBFAFB
'        txtOrderDate.BackColor = &HDBFAFB
        txtMemo.BackColor = &HDBFAFB
    Else
'        txtOrderNumber.BackColor = &HFFFFFF
'        txtOrderDate.BackColor = &HFFFFFF
        txtMemo.BackColor = &HFFFFFF
    End If
    iOpt = oPC.Configuration.Stores.FindStoreIdxByID(oPO.DELTOStoreID)

    Me.lblDelTo.Caption = oPC.Configuration.Stores.FindStoreByID(oPO.DELTOStoreID).DescriptionandDelAddress
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_PO.component(pLocked,pMemo,pTRID)", Array(pLocked, pMemo, pTRID)
End Sub

Private Sub cbDelTo_Click()
    On Error GoTo errHandler
    iOpt = OptionLoop(GetMax(iOpt, 1), oPC.Configuration.Stores.Count)
    oPO.setDelToStoreID oPC.Configuration.Stores(iOpt).ID
    Me.lblDelTo.Caption = oPC.Configuration.Stores(iOpt).DescriptionandDelAddress
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_PO.cbDelTo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancel = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_PO.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    oPO.SetMemo Memo
    If oPO.IsDirty Then
        If oPO.IsEditing Then
            oPO.ApplyEdit
        End If
    Else
        oPO.CancelEdit
    End If
    Set oPO = Nothing
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_PO.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    bCancel = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_PO.Form_Load", , EA_NORERAISE
    HandleError
End Sub

'Private Sub txtOrderDate_GotFocus()
'    AutoSelect txtOrderDate
'End Sub
'
'Private Sub txtOrderDate_LostFocus()
'    Me.cmdClose.Enabled = (IsDate(txtOrderDate) Or txtOrderDate = "")
'End Sub
'
'Private Sub txtOrderNumber_GotFocus()
'    AutoSelect txtOrderNumber
'End Sub
'
'Public Property Get OrderNumber() As String
'    OrderNumber = txtOrderNumber
'End Property
'
'Public Property Get OrderDate() As String
'    OrderDate = txtOrderDate
'End Property
Public Property Get Memo() As String
    On Error GoTo errHandler
    Memo = txtMemo
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_PO.Memo"
End Property


