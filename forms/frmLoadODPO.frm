VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLoadODPO 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Track purchase orders"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   6645
   Begin VB.Frame Frame4 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Select by ETA or status change"
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
      Height          =   1635
      Left            =   3390
      TabIndex        =   13
      Top             =   3810
      Width           =   3015
      Begin VB.CommandButton cmdStatusChange 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Fetch P.O.s affected by product status change or ETA change"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   225
         Picture         =   "frmLoadODPO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpChangedSince 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   345
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         _Version        =   393216
         Format          =   221839361
         CurrentDate     =   40022
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Changes since"
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
         Height          =   315
         Left            =   225
         TabIndex        =   16
         ToolTipText     =   "End of period during which the update can commence"
         Top             =   375
         Width           =   1065
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Filter by customer"
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
      Height          =   1755
      Left            =   210
      TabIndex        =   7
      Top             =   2085
      Width           =   3060
      Begin VB.CommandButton cmdSelectCust 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Select customer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   330
         Width           =   1470
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Any"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   330
         Width           =   840
      End
      Begin VB.Label lblCustomer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<Any customers>"
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
         Height          =   855
         Left            =   135
         TabIndex        =   10
         Top             =   855
         Width           =   2820
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Filter by supplier"
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
      Height          =   1755
      Left            =   210
      TabIndex        =   3
      Top             =   90
      Width           =   3060
      Begin VB.CommandButton cmdAny 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Any"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   300
         Width           =   720
      End
      Begin VB.CommandButton cmdSelectTP 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Select supplier"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   300
         Width           =   1410
      End
      Begin VB.Label lblSupplier 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<Any suppliers>"
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
         Height          =   855
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   2820
      End
   End
   Begin VB.ComboBox cboStaff 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   225
      TabIndex        =   1
      Top             =   4140
      Width           =   3075
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Select overdue orders"
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
      Height          =   3240
      Left            =   3375
      TabIndex        =   0
      Top             =   90
      Width           =   3030
      Begin VB.PictureBox Picture 
         BackColor       =   &H00D3D3CB&
         Height          =   2280
         Left            =   150
         ScaleHeight     =   2220
         ScaleWidth      =   2655
         TabIndex        =   17
         Top             =   270
         Width           =   2715
         Begin VB.OptionButton optDate 
            BackColor       =   &H00D3D3CB&
            Height          =   225
            Left            =   0
            TabIndex        =   25
            Top             =   1800
            Width           =   315
         End
         Begin VB.OptionButton optOD 
            BackColor       =   &H00D3D3CB&
            Caption         =   "overdue"
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
            Height          =   390
            Left            =   15
            TabIndex        =   23
            Top             =   285
            Width           =   2220
         End
         Begin VB.OptionButton optOD1M 
            BackColor       =   &H00D3D3CB&
            Caption         =   "> 1 month overdue"
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
            Height          =   390
            Left            =   0
            TabIndex        =   22
            Top             =   1455
            Width           =   2220
         End
         Begin VB.OptionButton optOD3W 
            BackColor       =   &H00D3D3CB&
            Caption         =   "> 3 weeks overdue"
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
            Height          =   390
            Left            =   30
            TabIndex        =   21
            Top             =   1170
            Width           =   2220
         End
         Begin VB.OptionButton optOD2W 
            BackColor       =   &H00D3D3CB&
            Caption         =   "> 2 weeks overdue"
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
            Height          =   390
            Left            =   15
            TabIndex        =   20
            Top             =   870
            Width           =   2220
         End
         Begin VB.OptionButton optOD1W 
            BackColor       =   &H00D3D3CB&
            Caption         =   "> 1 week overdue"
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
            Height          =   390
            Left            =   15
            TabIndex        =   19
            Top             =   585
            Width           =   2220
         End
         Begin VB.OptionButton optAll 
            BackColor       =   &H00D3D3CB&
            Caption         =   "All"
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
            Height          =   390
            Left            =   15
            TabIndex        =   18
            Top             =   30
            Value           =   -1  'True
            Width           =   2220
         End
         Begin MSComCtl2.DTPicker dtpSince 
            Height          =   285
            Left            =   1050
            TabIndex        =   26
            Top             =   1800
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            _Version        =   393216
            Format          =   221839361
            CurrentDate     =   40022
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Due by"
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
            Height          =   225
            Left            =   270
            TabIndex        =   24
            Top             =   1815
            Width           =   645
         End
      End
      Begin VB.CommandButton cmdGo 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Fetch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1635
         Picture         =   "frmLoadODPO.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2610
         Width           =   1000
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Or"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4530
      TabIndex        =   12
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter by staff member"
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
      Height          =   315
      Left            =   225
      TabIndex        =   2
      ToolTipText     =   "End of period during which the update can commence"
      Top             =   3930
      Width           =   1560
   End
End
Attribute VB_Name = "frmLoadODPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dteSince As Date
'Dim WithEvents cODCO As c_COLOD
Dim cODPO As c_POLSOS2
Attribute cODPO.VB_VarHelpID = -1
Dim POLSOS As ADODB.Recordset
Dim POLActions As ADODB.Recordset

Dim tlOperators As z_TextList
Dim lngOperatorID As Long
Dim lngSuppID As Long
Dim lngCUSTID As Long
Dim strType As String
Dim strCustomerName As String


Private Sub cmdAny_Click()
    On Error GoTo errHandler
    lngSuppID = 0
    If strType = "PO" Then
        lblCustomer.Caption = "<All suppliers>"
    Else
        lblCustomer.Caption = "<All customers>"
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.cmdAny_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSelectTP_Click()
    On Error GoTo errHandler
Dim frm As frmBrowseCustomers2
Dim frmS As frmBrowseSUppliers2

    Set frmS = New frmBrowseSUppliers2
    If frmS.WindowState <> vbNormal Then
        frmS.WindowState = vbNormal
        MsgBox "Setting state to normal"
    End If
    frmS.Show vbModal
    lngSuppID = frmS.SupplierID
    Me.lblSupplier.Caption = frmS.SupplierName
    Unload frmS

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.cmdSelectTP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cODCO_lngMax(p As Long)
    On Error GoTo errHandler
'    ProgressBar1.Max = p
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.cODCO_lngMax(p)", p, EA_NORERAISE
    HandleError
End Sub

Private Sub cODCO_lngProgress(p As Long)
    On Error GoTo errHandler
'    ProgressBar1.Value = p
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.cODCO_lngProgress(p)", p, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdGo_Click()
    On Error GoTo errHandler
Dim frmC As frmODCO
Dim frmP As frmODPO

    Screen.MousePointer = vbHourglass
        If optDate Then
            dteSince = dtpSince
        End If
        Set frmP = New frmODPO
        Set cODPO = New c_POLSOS2
        cODPO.LoadRecordsets POLSOS, POLActions, dteSince, lngSuppID, lngCUSTID, "", lngOperatorID
        frmP.component POLSOS, POLActions, dteSince, cboStaff, strCustomerName, lblSupplier.Caption
        frmP.Show
    
    Screen.MousePointer = vbDefault
   Unload Me

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.cmdGo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdStatusChange_Click()
    On Error GoTo errHandler
Dim frmC As frmODCO
Dim frmP As frmODPO

    Screen.MousePointer = vbHourglass
        If optDate Then
            dteSince = dtpSince
        End If
        Set frmP = New frmODPO
        Set cODPO = New c_POLSOS2
        cODPO.LoadRecordsets POLSOS, POLActions, dteSince, lngSuppID, lngCUSTID, "", lngOperatorID, dtpChangedSince
        frmP.component POLSOS, POLActions, CDate(0), cboStaff, strCustomerName, lblSupplier.Caption, dtpChangedSince
        frmP.Show
    
    Screen.MousePointer = vbDefault
   Unload Me

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.cmdStatusChange_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboStaff_Click()
    On Error GoTo errHandler
    lngOperatorID = tlOperators.Key(cboStaff.text)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.cboStaff_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSelectCust_Click()
    On Error GoTo errHandler
Dim frmC As New frmBrowseCustomers2
    frmC.Show vbModal
    lngCUSTID = frmC.CustomerID
    strCustomerName = frmC.CustomerName & " A/C no. " & frmC.Accnum
    lblCustomer.Caption = frmC.CustomerName & vbCrLf & "A/C no. " & frmC.Accnum
    Unload frmC

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.cmdSelectCust_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.DTPicker1_CallbackKeyDown(KeyCode,Shift,CallbackField,CallbackDate)", _
         Array(KeyCode, Shift, CallbackField, CallbackDate), EA_NORERAISE
    HandleError
End Sub

Private Sub DTPicker1_GotFocus()
    On Error GoTo errHandler
    optDate.Value = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.DTPicker1_GotFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub dtpSince_Change()
    On Error GoTo errHandler
    optDate.Value = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.dtpSince_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler

    Set tlOperators = New z_TextList
    Me.Caption = "Select overdue criteria"
    If Me.WindowState <> 2 Then
        Left = 100
        TOP = 100
        Width = 6930
        Height = 6200
    End If
    Me.dtpSince = DateAdd("d", -14, Date)
    Me.dtpChangedSince = DateAdd("d", -1, Date)
    tlOperators.Load ltStaff, , "<All>"
    LoadCombo cboStaff, tlOperators
    lngOperatorID = tlOperators.Key(cboStaff.text)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set tlOperators = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub optAll_Click()
    On Error GoTo errHandler
    dteSince = CDate(0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.optAll_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optOD_Click()
    On Error GoTo errHandler
    dteSince = Date
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.optOD_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optOD1M_Click()
    On Error GoTo errHandler
    dteSince = DateAdd("m", -1, Date)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.optOD1M_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optOD1W_Click()
    On Error GoTo errHandler
    dteSince = DateAdd("ww", -1, Date)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.optOD1W_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optOD2W_Click()
    On Error GoTo errHandler
    dteSince = DateAdd("ww", -2, Date)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.optOD2W_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optOD3W_Click()
    On Error GoTo errHandler
    dteSince = DateAdd("ww", -2, Date)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODPO.optOD3W_Click", , EA_NORERAISE
    HandleError
End Sub

