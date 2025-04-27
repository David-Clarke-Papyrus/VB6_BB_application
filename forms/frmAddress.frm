VERSION 5.00
Begin VB.Form frmAddress 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Address"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkGetsCat 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Gets catalogue"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   4260
      TabIndex        =   32
      Top             =   1605
      Width           =   1410
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
      Height          =   615
      Left            =   5130
      Picture         =   "frmAddress.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3315
      Width           =   945
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Catalogue postage"
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
      Height          =   1260
      Left            =   3495
      TabIndex        =   14
      Top             =   1965
      Width           =   2760
      Begin VB.PictureBox Picture 
         BackColor       =   &H00D3D3CB&
         Height          =   900
         Left            =   1275
         ScaleHeight     =   840
         ScaleWidth      =   1395
         TabIndex        =   36
         Top             =   225
         Width           =   1455
         Begin VB.OptionButton optAir 
            BackColor       =   &H00D3D3CB&
            Caption         =   "air"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Left            =   120
            TabIndex        =   38
            Top             =   -15
            Width           =   960
         End
         Begin VB.OptionButton optSurface 
            BackColor       =   &H00D3D3CB&
            Caption         =   "surface"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   120
            TabIndex        =   37
            Top             =   390
            Width           =   855
         End
      End
      Begin VB.CheckBox chkFormailing 
         BackColor       =   &H00D3D3CB&
         Caption         =   "For mailing"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   180
         TabIndex        =   33
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Height          =   1500
      Left            =   3480
      TabIndex        =   23
      Top             =   0
      Width           =   3525
      Begin VB.TextBox txtBusphone 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   705
         TabIndex        =   9
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   705
         TabIndex        =   8
         Top             =   180
         Width           =   1935
      End
      Begin VB.TextBox txtFax 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   705
         TabIndex        =   10
         Top             =   780
         Width           =   1935
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   705
         TabIndex        =   11
         Top             =   1080
         Width           =   2760
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   -105
         TabIndex        =   27
         Top             =   195
         Width           =   750
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Phone 2"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   45
         TabIndex        =   26
         Top             =   510
         Width           =   600
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "fax"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   210
         TabIndex        =   25
         Top             =   795
         Width           =   405
      End
      Begin VB.Label Label11 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   255
         TabIndex        =   24
         Top             =   1110
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Height          =   3225
      Left            =   30
      TabIndex        =   16
      Top             =   0
      Width           =   3435
      Begin VB.ComboBox cboCNTRY 
         Height          =   315
         Left            =   885
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   2445
         Width           =   1905
      End
      Begin VB.TextBox txtPCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   885
         TabIndex        =   7
         Top             =   2775
         Width           =   1230
      End
      Begin VB.TextBox txtAddressee 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   210
         TabIndex        =   0
         Top             =   345
         Width           =   3120
      End
      Begin VB.TextBox txtL6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   885
         TabIndex        =   6
         Top             =   2145
         Width           =   2445
      End
      Begin VB.TextBox txtL5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   885
         TabIndex        =   5
         Top             =   1845
         Width           =   2445
      End
      Begin VB.TextBox txtL4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   885
         TabIndex        =   4
         Top             =   1545
         Width           =   2445
      End
      Begin VB.TextBox txtL3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   885
         TabIndex        =   3
         Top             =   1245
         Width           =   2445
      End
      Begin VB.TextBox txtL2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   885
         TabIndex        =   2
         Top             =   945
         Width           =   2445
      End
      Begin VB.TextBox txtL1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   885
         TabIndex        =   1
         Top             =   645
         Width           =   2445
      End
      Begin VB.Label Label10 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Post code"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   120
         TabIndex        =   30
         Top             =   2805
         Width           =   720
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Country"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   150
         TabIndex        =   29
         Top             =   2490
         Width           =   690
      End
      Begin VB.Label Label8 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Addressee"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   300
         TabIndex        =   28
         Top             =   135
         Width           =   1050
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Province"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   75
         TabIndex        =   22
         Top             =   2175
         Width           =   765
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Town"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   75
         TabIndex        =   21
         Top             =   1875
         Width           =   765
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Line 4"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   75
         TabIndex        =   20
         Top             =   1590
         Width           =   765
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Line 3"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   75
         TabIndex        =   19
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Line 2"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   75
         TabIndex        =   18
         Top             =   1005
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Line 1"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   75
         TabIndex        =   17
         Top             =   705
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
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
      Left            =   6075
      Picture         =   "frmAddress.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3315
      Width           =   945
   End
   Begin VB.TextBox txtDescription 
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
      Left            =   2250
      TabIndex        =   35
      Top             =   3435
      Width           =   2715
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H00D3D3CB&
      ForeColor       =   &H000000FF&
      Height          =   810
      Left            =   2250
      TabIndex        =   31
      Top             =   4170
      Width           =   2910
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00D3D3CB&
      Caption         =   "Dispatch mode to this address"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   45
      TabIndex        =   15
      Top             =   3480
      Width           =   2145
   End
End
Attribute VB_Name = "frmAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oAdd As a_Address
Attribute oAdd.VB_VarHelpID = -1
Dim flgLoading As Boolean

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
'    If oAdd.isediting Then
'        MsgBox "You have not saved the changes you have made." & vbCrLf & "This action has been cancelled.", vbInformation, "Can't do this now"
'        Cancel = True
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Private Sub oADD_Valid(pErrors As String, pValid As Boolean)
    On Error GoTo errHandler
    Me.lblErrors = pErrors
    Me.cmdOK.Enabled = pValid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.oADD_Valid(pErrors,pValid)", Array(pErrors, pValid), EA_NORERAISE
    HandleError
End Sub
Sub component(pAdd As a_Address)
    On Error GoTo errHandler
    Set oAdd = pAdd
    oAdd.BeginEdit
    If oAdd.Addressee = "<Double-click to enter new details>" Then
        oAdd.SetAddressee ""
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.Component(pAdd)", pAdd
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oAdd.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    oAdd.ApplyEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        TOP = 1800
        Left = 150
        Height = 4395
    End If
    LoadControls
    oAdd.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Sub LoadControls()
    On Error GoTo errHandler
Dim bResult As Boolean
Dim strErrors As String

    flgLoading = True
    txtDescription = oAdd.Description
'    If oAdd.cu.initials > "" Then
'        txtAddressee = oCust.Title & " " & oCust.initials & " " & oCust.Name
'    Else
        txtAddressee = oAdd.Addressee
'    End If

    txtL1 = oAdd.Line1
    txtL2 = oAdd.Line2
    txtL3 = oAdd.Line3
    txtL4 = oAdd.Line4
    txtL5 = oAdd.Line5
    txtL6 = oAdd.Line6
    chkFormailing = IIf(oAdd.ForMailing, 1, 0)
    If oPC.SupportsCatalogue = True Then
        Me.chkGetsCat = IIf(oAdd.GetsCatalogue, 1, 0)
    Else
        Me.chkGetsCat.Visible = False
    End If

    txtPhone = oAdd.Phone
    txtFax = oAdd.Fax
    txtBusphone = oAdd.BusPhone
    txtPCode = oAdd.pCode
    txtEmail = oAdd.EMail
    optAir = (oAdd.PostageType = 1)
    optSurface = (oAdd.PostageType = 2)
    LoadCombo cboCNTRY, oAdd.Countries
    If oAdd.CountryID > 0 Then
        On Error Resume Next   'a drop down list can only assume a value it 'knows'
        cboCNTRY.text = oAdd.Countries(CStr(oAdd.CountryID))
        On Error GoTo errHandler
    ElseIf oPC.Configuration.LocalCountryID Then
        On Error Resume Next   'a drop down list can only assume a value it 'knows'
        cboCNTRY.text = oAdd.Countries(CStr(oPC.Configuration.LocalCountryID))
        oAdd.CountryID = oPC.Configuration.LocalCountryID
        On Error GoTo errHandler
    End If
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.LoadControls"
End Sub



Private Sub optAir_Click()
    On Error GoTo errHandler
    oAdd.PostageType = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.optAir_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optSurface_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oAdd.PostageType = 2
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.optSurface_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub optBillTo_Click()
    On Error GoTo errHandler
    oAdd.Category = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.optBillTo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optDeliverTo_Click()
    On Error GoTo errHandler
    oAdd.Category = 2
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.optDeliverTo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub optOther_Click()
    On Error GoTo errHandler
    oAdd.Category = 3
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.optOther_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtDescription_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oAdd.SetDescription (txtDescription)
    If Err Then
      Beep
      intPos = txtDescription.SelStart
      txtDescription = oAdd.Description
      txtDescription.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtDescription_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDescription_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtDescription = oAdd.Description
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtDescription_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDescription_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim bOK As Boolean
  '  Cancel = Not oAdd.SetDescription(txtDescription)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtDescription_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtAddressee_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtAddressee = oAdd.Addressee
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtAddressee_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAddressee_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    On Error Resume Next
    Cancel = Not oAdd.SetAddressee(txtAddressee)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtAddressee_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtAddressee_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    If flgLoading Then Exit Sub
    oAdd.SetAddressee (txtAddressee)
    If Err Then
      Beep
      intPos = txtAddressee.SelStart
      txtAddressee = oAdd.Addressee
      txtAddressee.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtAddressee_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtL1_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtL1 = oAdd.Line1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL1_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtL1_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    If flgLoading Then Exit Sub
    oAdd.SetLine1 (txtL1)
    If Err Then
      Beep
      intPos = txtL1.SelStart
      txtL1 = oAdd.Line1
      txtL1.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL1_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtL1_Validate(Cancel As Boolean)
    On Error Resume Next
    Cancel = Not oAdd.SetLine1(txtL1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL1_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtL2_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtL2 = oAdd.Line2
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL2_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtL2_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    If flgLoading Then Exit Sub
    oAdd.SetLine2 (txtL2)
    If Err Then
      Beep
      intPos = txtL2.SelStart
      txtL2 = oAdd.Line2
      txtL2.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL2_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtL2_Validate(Cancel As Boolean)
    On Error Resume Next
    Cancel = Not oAdd.SetLine2(txtL2)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL2_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtL3_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtL3 = oAdd.Line3
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL3_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtL3_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    If flgLoading Then Exit Sub
    oAdd.SetLine3 (txtL3)
    If Err Then
      Beep
      intPos = txtL3.SelStart
      txtL3 = oAdd.Line3
      txtL3.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL3_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtL3_Validate(Cancel As Boolean)
    On Error Resume Next
    Cancel = Not oAdd.SetLine3(txtL3)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL3_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtL4_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtL4 = oAdd.Line4
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL4_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtL4_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    If flgLoading Then Exit Sub
    oAdd.SetLine4 (txtL4)
    If Err Then
      Beep
      intPos = txtL4.SelStart
      txtL4 = oAdd.Line4
      txtL4.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL4_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtL4_Validate(Cancel As Boolean)
    On Error Resume Next
    Cancel = Not oAdd.SetLine4(txtL4)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL4_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtL5_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtL5 = oAdd.Line5
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL5_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtL5_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    If flgLoading Then Exit Sub
    oAdd.SetLine5 (txtL5)
    If Err Then
      Beep
      intPos = txtL5.SelStart
      txtL5 = oAdd.Line5
      txtL5.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL5_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtL5_Validate(Cancel As Boolean)
    On Error Resume Next
    Cancel = Not oAdd.SetLine5(txtL5)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL5_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtL6_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtL6 = oAdd.Line6
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL6_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtL6_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    If flgLoading Then Exit Sub
    oAdd.SetLine6 (txtL6)
    If Err Then
      Beep
      intPos = txtL6.SelStart
      txtL6 = oAdd.Line6
      txtL6.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL6_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtL6_Validate(Cancel As Boolean)
    On Error Resume Next
    Cancel = Not oAdd.SetLine6(txtL6)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtL6_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub cboCNTRY_Validate(Cancel As Boolean)
    On Error Resume Next
    If flgLoading Then Exit Sub
    If Not cboCNTRY.ListIndex = -1 Then
        oAdd.CountryID = oAdd.Countries.Key(cboCNTRY)
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.cboCNTRY_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtPCode_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtPCode = oAdd.pCode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtPCode_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPCode_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    If flgLoading Then Exit Sub
    oAdd.SetPCode (txtPCode)
    If Err Then
      Beep
      intPos = txtPCode.SelStart
      txtPCode = oAdd.pCode
      txtPCode.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtPCode_Change", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtPhone_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtPhone = oAdd.Phonef
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtPhone_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPhone_Change()
Dim intPos As Integer
    On Error Resume Next
    If flgLoading Then Exit Sub
    oAdd.SetPhone (txtPhone)
    If Err Then
      Beep
      intPos = txtPhone.SelStart
      txtPhone = oAdd.Phone
      txtPhone.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtPhone_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPhone_Validate(Cancel As Boolean)
    On Error Resume Next
    If txtPhone = "" Then Exit Sub
    txtPhone = PhoneFormat(txtPhone, oPC.DefaultAreaCode)
    Cancel = Not oAdd.SetPhone(txtPhone)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtPhone_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtBusPhone_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtBusphone = oAdd.BusPhone
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtBusPhone_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtBusPhone_Change()
    On Error Resume Next
Dim intPos As Integer
    If flgLoading Then Exit Sub
    oAdd.SetBusPhone (txtBusphone)
    If Err Then
      Beep
      intPos = txtBusphone.SelStart
      txtBusphone = oAdd.BusPhone
      txtBusphone.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtBusPhone_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtBusPhone_Validate(Cancel As Boolean)
    On Error Resume Next
    txtBusphone = PhoneFormat(txtBusphone, oPC.DefaultAreaCode)
    Cancel = Not oAdd.SetBusPhone(txtBusphone)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtBusPhone_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtFax_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtFax = PhoneFormat(txtFax, oPC.DefaultAreaCode)
    txtFax = oAdd.Fax
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtFax_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFax_Change()
    On Error Resume Next
Dim intPos As Integer
    If flgLoading Then Exit Sub
    oAdd.SetFax (txtFax)
    If Err Then
      Beep
      intPos = txtFax.SelStart
      txtFax = oAdd.Fax
      txtFax.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtFax_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFax_Validate(Cancel As Boolean)
    On Error Resume Next
    txtFax = PhoneFormat(txtFax, oPC.DefaultAreaCode)
    Cancel = Not oAdd.SetFax(txtFax)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtFax_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtEMail_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtEmail = oAdd.EMail
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtEMail_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtEMail_Validate(Cancel As Boolean)
Dim intPos As Integer
    On Error Resume Next
    Cancel = Not oAdd.SetEmail(txtEmail)
    If Err Then
      Beep
      intPos = txtFax.SelStart
      txtFax = oAdd.Fax
      txtFax.SelStart = intPos - 1
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.txtEMail_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub chkGetsCat_Click()
    On Error GoTo errHandler
    oAdd.GetsCatalogue = (chkGetsCat = 1)
   ' chkGetsCatalogue = oAdd.GetsCatalogue = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.chkGetsCat_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkFormailing_Click()
    On Error GoTo errHandler
If flgLoading Then Exit Sub
    oAdd.ForMailing = (chkFormailing = 1)
    Me.optAir.Enabled = oAdd.ForMailing
    Me.optSurface.Enabled = oAdd.ForMailing
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAddress.chkFormailing_Click", , EA_NORERAISE
    HandleError
End Sub

