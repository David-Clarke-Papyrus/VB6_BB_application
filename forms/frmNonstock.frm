VERSION 5.00
Begin VB.Form frmServiceItem 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Service item"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   5115
   Begin VB.TextBox txtCharge 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   1500
      TabIndex        =   12
      Top             =   1785
      Width           =   1380
   End
   Begin VB.CheckBox chkExcludeFromSales 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Exclude from sales"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   105
      TabIndex        =   11
      Top             =   2595
      Width           =   1620
   End
   Begin VB.ComboBox cboProductType 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1515
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1380
      Width           =   2115
   End
   Begin VB.TextBox txtVAT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   1515
      TabIndex        =   7
      Top             =   1005
      Width           =   1380
   End
   Begin VB.TextBox txtErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   945
      Left            =   195
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   6465
      Width           =   4350
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
      Height          =   615
      Left            =   3015
      Picture         =   "frmNonstock.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2490
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
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
      Left            =   4050
      Picture         =   "frmNonstock.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2490
      Width           =   1000
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1515
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   615
      Width           =   2415
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1515
      TabIndex        =   0
      Top             =   240
      Width           =   1680
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Default charge"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   360
      TabIndex        =   13
      Top             =   1815
      Width           =   1080
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Product type"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   420
      TabIndex        =   10
      Top             =   1410
      Width           =   1035
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "V.A.T. Rate"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   375
      TabIndex        =   8
      Top             =   1035
      Width           =   1080
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   300
      TabIndex        =   5
      Top             =   660
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   795
      TabIndex        =   4
      Top             =   285
      Width           =   660
   End
End
Attribute VB_Name = "frmServiceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oProd As a_Product
Attribute oProd.VB_VarHelpID = -1
Dim flgLoading As Boolean

Sub component(pProduct As a_Product)
    On Error GoTo errHandler
        Set oProd = pProduct
        oProd.BeginEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.component(pProduct)", pProduct
End Sub



Private Sub cmdDelete_Click()
    On Error GoTo errHandler

    If MsgBox("Confirm deletion of product: " & oProd.Title & vbCrLf & "This is permanent!", vbOKCancel + vbExclamation, "Confirm") = vbOK Then
        oProd.Delete
        oProd.ApplyEdit
        Unload Me
    End If


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdSetDefault_Click()
    On Error GoTo errHandler
    Me.txtVAT = oPC.Configuration.VATRate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.cmdSetDefault_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub chkExcludeFromSales_Click()
    On Error GoTo errHandler
    oProd.ExcludeFromSales = IIf(Me.chkExcludeFromSales = 1, True, False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.chkExcludeFromSales_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oProd_Valid(strMsg As String)
    On Error GoTo errHandler
    Me.txtErrors = strMsg
    Me.cmdOK.Enabled = (strMsg = "")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.oProd_Valid(strMsg)", strMsg, EA_NORERAISE
    HandleError
End Sub
Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oProd.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long
    oProd.SetServiceItem
    oProd.ApplyEdit lngResult
    If lngResult = 99 Then
        MsgBox "Invalid values - check that the code is has not been already used", , "Save failed"
    Else
        Unload Me
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        TOP = 50
        Left = 50
        Width = 5700
        Height = 3650
    End If
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    Me.txtCode = oProd.code
    Me.txtTitle = oProd.Title
    Me.txtVAT = oProd.VATRateToUseF
    Me.txtCharge = oProd.SPF
    Me.chkExcludeFromSales = IIf(oProd.ExcludeFromSales = True, 1, 0)
    LoadCombo cboProductType, oPC.Configuration.ProductTypes_Short
    cboProductType = oPC.Configuration.ProductTypes.Item(oProd.ProductTypeID)
    flgLoading = False
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.LoadControls"
End Sub




Private Sub txtCharge_GotFocus()
    On Error GoTo errHandler
    txtCharge = oProd.SP
    AutoSelect txtCharge

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.txtCharge_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCharge_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oProd.SetSP(txtCharge) Then
        Cancel = True
    End If
    txtCharge = oProd.SPF

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.txtCharge_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oProd.SetCode txtCode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtTitle_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtTitle = oProd.Title
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.txtTitle_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oProd.SetTitle(txtTitle)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.txtTitle_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oProd.SetTitle (txtTitle)
    If Err Then
      Beep
      intPos = txtTitle.SelStart
      txtTitle = oProd.Title
      txtTitle.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.txtTitle_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtVAT_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtVAT
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.txtVAT_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtVAT_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtVAT = oProd.VATRateToUse
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.txtVAT_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtVAT_Validate(Cancel As Boolean)
    On Error GoTo errHandler
   If flgLoading Then Exit Sub
   Cancel = Not oProd.SetVAT(txtVAT)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.txtVAT_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub cboProductType_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oProd.SetProductTypeID oPC.Configuration.ProductTypes.Key(cboProductType)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmServiceItem.cboProductType_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceItem.cboProductType_Click", , EA_NORERAISE
    HandleError
End Sub

