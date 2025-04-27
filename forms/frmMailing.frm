VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmMailing 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse mailing addresses"
   ClientHeight    =   7545
   ClientLeft      =   750
   ClientTop       =   345
   ClientWidth     =   8160
   Icon            =   "frmMailing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Labels"
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
      Height          =   900
      Left            =   75
      TabIndex        =   18
      Top             =   6600
      Width           =   6735
      Begin VB.CommandButton cmdList 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Print label preview"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4680
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Click to find all customers matching the retrictions selected."
         Top             =   330
         UseMaskColor    =   -1  'True
         Width           =   1830
      End
      Begin VB.CheckBox chkCataloguePrinting 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Note this mailing on cust. rec."
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
         Height          =   315
         Left            =   1680
         TabIndex        =   20
         Top             =   405
         Width           =   3015
      End
      Begin VB.CommandButton cmdLabels 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Print labels"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   75
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Click to find all customers matching the retrictions selected."
         Top             =   315
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      Picture         =   "frmMailing.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6705
      Width           =   1000
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00B4FFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   6375
      TabIndex        =   15
      Top             =   6375
      Width           =   315
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Find addresses meeting these criteria"
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
      Height          =   2460
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Select any one criteria.  If using dates, a selection between dates is catered for"
      Top             =   105
      Width           =   7950
      Begin VB.TextBox txtAddressContains 
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
         Left            =   90
         TabIndex        =   21
         ToolTipText     =   "Customer name starts like this"
         Top             =   1905
         Width           =   1695
      End
      Begin VB.ComboBox cboIG3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Select a customer type or allow any customer to be found"
         Top             =   1500
         Width           =   1725
      End
      Begin VB.ComboBox cboIG2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Select a customer type or allow any customer to be found"
         Top             =   1065
         Width           =   1725
      End
      Begin VB.ComboBox cboIG1 
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
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Select a customer type or allow any customer to be found"
         Top             =   630
         Width           =   1725
      End
      Begin VB.ComboBox cboCT 
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
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Select a customer type or allow any customer to be found"
         Top             =   555
         Width           =   1725
      End
      Begin CoolButtonControl.CoolButton cbCatalogue 
         Height          =   465
         Left            =   4320
         TabIndex        =   4
         ToolTipText     =   "Click to change option"
         Top             =   345
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   820
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
         Caption         =   "Catalogue: no restrictions"
         Style           =   1
         BackStyle       =   0
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6840
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmMailing.frx":04D4
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Click to find all customers matching the retrictions selected."
         Top             =   1230
         UseMaskColor    =   -1  'True
         Width           =   1000
      End
      Begin VB.TextBox txtArg 
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
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Customer name starts like this"
         Top             =   1230
         Width           =   1695
      End
      Begin CoolButtonControl.CoolButton cbMailType 
         Height          =   465
         Left            =   4320
         TabIndex        =   5
         ToolTipText     =   "Click to change option"
         Top             =   1365
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   820
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
         Caption         =   "Surface or Airmail"
         Style           =   1
         BackStyle       =   0
      End
      Begin CoolButtonControl.CoolButton cbOverseas 
         Height          =   465
         Left            =   4320
         TabIndex        =   6
         ToolTipText     =   "Click to change option"
         Top             =   855
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   820
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
         Caption         =   "Local and overseas"
         Style           =   1
         BackStyle       =   0
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address contains"
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
         Height          =   270
         Left            =   120
         TabIndex        =   22
         Top             =   1635
         Width           =   1680
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Interest groups"
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
         Height          =   270
         Left            =   2295
         TabIndex        =   11
         Top             =   375
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer type"
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
         Height          =   270
         Left            =   150
         TabIndex        =   8
         Top             =   300
         Width           =   1605
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name like . . ."
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
         Height          =   270
         Left            =   150
         TabIndex        =   3
         Top             =   960
         Width           =   1680
      End
   End
   Begin TrueOleDBGrid60.TDBGrid CustGrid 
      Height          =   3465
      Left            =   60
      OleObjectBlob   =   "frmMailing.frx":085E
      TabIndex        =   9
      Top             =   2850
      Width           =   7935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Gets catalogue"
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
      Height          =   270
      Left            =   5025
      TabIndex        =   16
      Top             =   6405
      Width           =   1260
   End
   Begin VB.Label lblReccFound 
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
      Height          =   270
      Left            =   45
      TabIndex        =   14
      Top             =   6300
      Width           =   2790
   End
End
Attribute VB_Name = "frmMailing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCust As c_Customer
Dim oCust As a_Customer
Dim dispCust As d_Customer
Dim flgLoading As Boolean
Dim CustomerTypes_tl As z_TextList
Dim InterestGroups_tl As z_TextList
Dim lngTPID As Long
Dim strACCNum As String
Dim blnNoRecordsReturned As Boolean
Dim enArg_Cat As enCatalogue
Dim enArg_Overseas As enOverseas
Dim enArg_MailType As enMailType
Dim ofrm As frmCustomerPreview
Dim XA As New XArrayDB


Private Sub cbCatalogue_Click()
    On Error GoTo errHandler
    enArg_Cat = OptionLoop(enArg_Cat, 3)
    Select Case enArg_Cat
    Case enGetsCatalogueYes
        cbCatalogue.Caption = "Gets catalogue - YES"
    Case enGetsCatalogueNo
        cbCatalogue.Caption = "Gets catalogue - NO"
    Case enGetsCatalogueEither
        cbCatalogue.Caption = "Catalogue: no restriction"
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.cbCatalogue_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboIG1_click()
    On Error GoTo errHandler
    If cboIG1.ListIndex > 0 Then
        cboIG2.Enabled = True
    Else
        If cboIG2.Enabled Then
            cboIG2.ListIndex = 0
            cboIG2.Enabled = False
        End If
        If cboIG3.Enabled Then
            cboIG3.ListIndex = 0
            cboIG3.Enabled = False
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.cboIG1_click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cboIG2_click()
    On Error GoTo errHandler
    If cboIG2.ListIndex > 0 Then
        cboIG3.Enabled = True
    Else
        If cboIG3.Enabled Then
            cboIG3.ListIndex = 0
            cboIG3.Enabled = False
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.cboIG2_click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cbOverseas_Click()
    On Error GoTo errHandler
    enArg_Overseas = OptionLoop(enArg_Overseas, 3)
    Select Case enArg_Overseas
    Case enOverseasYes
        cbOverseas.Caption = "Overseas"
    Case enOverseasNo
        cbOverseas.Caption = "Local"
    Case enOverseasEither
        cbOverseas.Caption = "Local and overseas"
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.cbOverseas_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cbMailType_Click()
    On Error GoTo errHandler
    enArg_MailType = OptionLoop(enArg_MailType, 4)
    Select Case enArg_MailType
    Case enAirmail
        cbMailType.Caption = "Airmail"
    Case enSurfaceMail
        cbMailType.Caption = "Surface mail"
    Case enEitherMail
        cbMailType.Caption = "Surface or Airmail"
    Case enAll
        cbMailType.Caption = "Mailing and non-mailing"
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.cbMailType_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFind_Click()
    On Error GoTo errHandler
Dim bRecsFound As Boolean
Dim CTKey As Long
Dim IG1Key As Long
Dim IG2Key As Long
Dim IG3key As Long


    blnNoRecordsReturned = False
    
    Set cCust = Nothing
    Set cCust = New c_Customer
    MousePointer = vbHourglass
    
    CTKey = CustomerTypes_tl.Key(cboCT)
    IG1Key = InterestGroups_tl.Key(cboIG1)
    IG2Key = InterestGroups_tl.Key(cboIG2)
    IG3key = InterestGroups_tl.Key(cboIG3)
    If CTKey = 0 And IG1Key = 0 And IG2Key = 0 And IG3key = 0 And FNS(txtAddressContains) = "" And FNS(txtArg) = "" Then
        MsgBox "Enter criteria before searching.", vbInformation, "Can't do this"
        Exit Sub
    End If
    cCust.LoadForMailing bRecsFound, enArg_Cat, enArg_Overseas, enArg_MailType, Me.txtArg, _
            CTKey, IG1Key, IG2Key, IG3key, FNS(txtAddressContains)
    
    If blnNoRecordsReturned Then
        MsgBox "No records found", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        GoTo EXIT_Handler
    End If
    
    LoadArray
    CustGrid.ReBind
    Me.lblReccFound.Caption = CStr(cCust.Count) & " records"
'    LoadListView
'
EXIT_Handler:
    MousePointer = vbDefault
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.cmdFind_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFind_LostFocus()
    On Error GoTo errHandler
 '   LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.cmdFind_LostFocus", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdLabels_Click()
    On Error GoTo errHandler
'Dim frm As frmCatNo
'Dim strNo As String
'
'    If Me.chkCataloguePrinting Then
'        Set frm = New frmCatNo
'        frm.Show vbModal
'        If frm.Cancelled = False Then
'            strNo = frm.CatNo
'            cCust.Markup strNo
'        End If
'        Unload frm
'    End If
'    cCust.PrintLabels
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.cmdLabels_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdList_Click()
    On Error GoTo errHandler
Dim strNo As String

    cCust.PrintMaillist
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.cmdList_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub CustGrid_DblClick()
    On Error GoTo errHandler
Dim lngID As Long
Dim blnEdit As Boolean
    Set ofrm = New frmCustomerPreview
    lngID = val(XA(CustGrid.Bookmark, 5))
    Set oCust = Nothing
    Set oCust = New a_Customer
    oCust.Load lngID
    ofrm.component oCust    ', False
    ofrm.Show
    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmMailing: CustGrid_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmMailing: CustGrid_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.CustGrid_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set oCust = Nothing
    Set cCust = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    Set cCust = New c_Customer
    Set CustomerTypes_tl = New z_TextList
    Set InterestGroups_tl = New z_TextList
    InterestGroups_tl.Load ltInterestGroupAll, , "<ANY>"
    CustomerTypes_tl.Load ltCustomerTypeAll, , "<ANY>"
    LoadCombo cboIG1, InterestGroups_tl
    LoadCombo cboIG2, InterestGroups_tl
    LoadCombo cboIG3, InterestGroups_tl
    If Me.WindowState <> 2 Then
        Me.TOP = 0
        Me.Left = 130
        Me.Width = 8200
        Me.Height = 7950
    End If
    cmdList.Enabled = True
    LoadControls
    enArg_MailType = enAll
    cbMailType.Caption = "Mailing and non-mailing"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set cCust = Nothing
    Set ofrm = Nothing
    Set CustomerTypes_tl = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    LoadCombo cboCT, CustomerTypes_tl
    If CustomerTypes_tl.Count > 0 Then cboCT = "<ANY>" 'CustomerTypes_tl.Item(oPC.Configuration.DefaultCT)
    txtArg = ""
    lngTPID = 0
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.LoadControls"
End Sub
Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_Customer
Dim itmList As ListItem
Dim lngIndex As Long
    XA.ReDim 1, cCust.Count, 1, 6
    For lngIndex = 1 To cCust.Count
        With objItem
            Set objItem = cCust.Item(lngIndex)
            XA.Value(lngIndex, 1) = lngIndex
            XA.Value(lngIndex, 2) = objItem.Fullname2
            XA.Value(lngIndex, 3) = objItem.ListAddress
            XA.Value(lngIndex, 5) = objItem.ID
            XA.Value(lngIndex, 6) = objItem.GetsCatalogue
        End With
    Next
    CustGrid.Array = XA
    cmdLabels.Enabled = cCust.Count > 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.LoadArray"
End Sub

Private Sub CustGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If XA(Bookmark, 6) = True Then
        RowStyle.BackColor = RGB(282, 274, 180)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailing.CustGrid_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub



