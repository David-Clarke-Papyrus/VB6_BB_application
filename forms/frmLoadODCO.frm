VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLoadODCO 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Fetch overdue customer orders"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
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
   ScaleHeight     =   5910
   ScaleWidth      =   6795
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
      Left            =   3495
      TabIndex        =   6
      Top             =   3960
      Width           =   3015
      Begin VB.CommandButton cmdStatusChange 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Fetch C.O.s affected by product status change or ETA change"
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
         Picture         =   "frmLoadODCO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   690
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpChangedSince 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
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
         TabIndex        =   9
         ToolTipText     =   "End of period during which the update can commence"
         Top             =   375
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   2340
      Left            =   195
      TabIndex        =   2
      Top             =   120
      Width           =   3060
      Begin VB.CommandButton cmdAny 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Any"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2070
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton cmdSelectTP 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Select customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1905
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
         Height          =   690
         Left            =   135
         TabIndex        =   4
         Top             =   750
         Width           =   2820
      End
   End
   Begin VB.ComboBox cboStaff 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   210
      TabIndex        =   1
      Top             =   3090
      Width           =   3075
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Select age"
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
      Height          =   3390
      Left            =   3480
      TabIndex        =   0
      Top             =   90
      Width           =   3030
      Begin VB.PictureBox Picture 
         BackColor       =   &H00D3D3CB&
         Height          =   2385
         Left            =   135
         ScaleHeight     =   2325
         ScaleWidth      =   2760
         TabIndex        =   13
         Top             =   255
         Width           =   2820
         Begin VB.OptionButton optOD 
            BackColor       =   &H00D3D3CB&
            Caption         =   "overdue"
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
            Height          =   390
            Left            =   240
            TabIndex        =   20
            Top             =   375
            Width           =   2220
         End
         Begin VB.OptionButton optOD1M 
            BackColor       =   &H00D3D3CB&
            Caption         =   "> 1 month overdue"
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
            Height          =   390
            Left            =   240
            TabIndex        =   19
            Top             =   1605
            Width           =   2220
         End
         Begin VB.OptionButton optOD3W 
            BackColor       =   &H00D3D3CB&
            Caption         =   "> 3 weeks overdue"
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
            Height          =   390
            Left            =   240
            TabIndex        =   18
            Top             =   1305
            Width           =   2220
         End
         Begin VB.OptionButton optOD2W 
            BackColor       =   &H00D3D3CB&
            Caption         =   "> 2 weeks overdue"
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
            Height          =   390
            Left            =   240
            TabIndex        =   17
            Top             =   990
            Width           =   2220
         End
         Begin VB.OptionButton optOD1W 
            BackColor       =   &H00D3D3CB&
            Caption         =   "> 1 week overdue"
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
            Height          =   390
            Left            =   240
            TabIndex        =   16
            Top             =   690
            Width           =   2220
         End
         Begin VB.OptionButton optAll 
            BackColor       =   &H00D3D3CB&
            Caption         =   "All"
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
            Height          =   390
            Left            =   240
            TabIndex        =   15
            Top             =   75
            Value           =   -1  'True
            Width           =   2220
         End
         Begin VB.OptionButton optDate 
            BackColor       =   &H00D3D3CB&
            Height          =   225
            Left            =   240
            TabIndex        =   14
            Top             =   2025
            Width           =   315
         End
         Begin MSComCtl2.DTPicker dtpSince 
            Height          =   285
            Left            =   1155
            TabIndex        =   21
            Top             =   2010
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            _Version        =   393216
            Format          =   221839361
            CurrentDate     =   40022
         End
         Begin VB.Label Label4 
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
            Left            =   420
            TabIndex        =   22
            Top             =   2040
            Width           =   645
         End
      End
      Begin VB.CommandButton cmdGo 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Fetch"
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
         Left            =   1635
         Picture         =   "frmLoadODCO.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2670
         Width           =   1000
      End
   End
   Begin VB.Label Label1 
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
      Left            =   255
      TabIndex        =   11
      ToolTipText     =   "End of period during which the update can commence"
      Top             =   2850
      Width           =   1560
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Or"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4635
      TabIndex        =   10
      Top             =   3630
      Width           =   720
   End
End
Attribute VB_Name = "frmLoadODCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dteSince As Date
Dim cODCO As c_COLOD2
Attribute cODCO.VB_VarHelpID = -1
'Dim WithEvents cODPO As c_POLsOS
Dim tlOperators As z_TextList
Dim lngOperatorID As Long
Dim lngTPID As Long
Dim strType As String
Dim COLS As ADODB.Recordset
Dim POLS As ADODB.Recordset
Dim COLActs As ADODB.Recordset

Public Sub component(pType As String)
    On Error GoTo errHandler
    strType = "CO"
    Caption = "Select overdue criteria"
    cmdSelectTP.Caption = "Select customer"
    lblCustomer.Caption = "<Any>"
    cmdAny.Caption = "Any"
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.component(pType)", pType
End Sub
Private Sub cmdAny_Click()
    On Error GoTo errHandler
    lngTPID = 0
    If strType = "PO" Then
        lblCustomer.Caption = "<Any>"
    Else
        lblCustomer.Caption = "<Any>"
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.cmdAny_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSelectTP_Click()
    On Error GoTo errHandler
Dim frm As frmBrowseCustomers2
Dim frmS As frmBrowseSUppliers2

    If strType = "PO" Then
        Set frmS = New frmBrowseSUppliers2
        If frmS.WindowState <> vbNormal Then
            frmS.WindowState = vbNormal
            MsgBox "Setting state to normal"
        End If
        frmS.Show vbModal
        lngTPID = frmS.SupplierID
        lblCustomer.Caption = frmS.SupplierName
        Unload frmS
    Else
        Set frm = New frmBrowseCustomers2
        frm.Show vbModal
        lngTPID = frm.CustomerID
        lblCustomer.Caption = frm.CustomerName & " A/C no. " & frm.Accnum
        Unload frm
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.cmdSelectTP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cODCO_lngMax(p As Long)
    On Error GoTo errHandler
'    ProgressBar1.Max = p
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.cODCO_lngMax(p)", p, EA_NORERAISE
    HandleError
End Sub

Private Sub cODCO_lngProgress(p As Long)
    On Error GoTo errHandler
'    ProgressBar1.Value = p
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.cODCO_lngProgress(p)", p, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdGo_Click()
    On Error GoTo errHandler
Dim frmC As frmODCO

    Screen.MousePointer = vbHourglass
   
    Set frmC = New frmODCO
    Set cODCO = New c_COLOD2
    cODCO.LoadRecordsets COLS, POLS, COLActs, dteSince, lngTPID, "", lngOperatorID
    frmC.Component3 COLS, POLS, COLActs, dteSince, cboStaff, lblCustomer.Caption, CDate(0)
    frmC.Show
    
    Screen.MousePointer = vbDefault
    
    Unload Me

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.cmdGo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdStatusChange_Click()
    On Error GoTo errHandler
Dim frmC As frmODCO

    Screen.MousePointer = vbHourglass
        If optDate Then
            dteSince = dtpSince
        End If
    Set frmC = New frmODCO
    Set cODCO = New c_COLOD2
    cODCO.LoadRecordsets COLS, POLS, COLActs, dteSince, lngTPID, "", lngOperatorID, dtpChangedSince
    frmC.Component3 COLS, POLS, COLActs, dteSince, cboStaff, lblCustomer.Caption, dtpChangedSince
    frmC.Show
'        cODPO.LoadRecordsets POLSOS, POLActions, dteSince, lngSuppID, lngCUSTID, "", lngOperatorID, dtpChangedSince
'        frmP.Component POLSOS, POLActions, CDate(0), cboStaff, strCustomerName, lblSupplier.Caption
'        frmP.Show
    
    Screen.MousePointer = vbDefault
   Unload Me


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.cmdStatusChange_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboStaff_Click()
    On Error GoTo errHandler
    lngOperatorID = tlOperators.Key(cboStaff.text)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.cboStaff_Click", , EA_NORERAISE
    HandleError
End Sub





Private Sub dtpSince_Change()
    On Error GoTo errHandler
    optDate = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.dtpSince_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler

    Set tlOperators = New z_TextList
    Me.Caption = "Select overdue criteria"
    If Me.WindowState <> 2 Then
        Left = 100
        TOP = 100
        Width = 6900
        Height = 6300
    End If
    Me.dtpSince = DateAdd("d", -14, Date)
    Me.dtpChangedSince = DateAdd("d", -1, Date)
    tlOperators.Load ltStaff, , "<All>"
    LoadCombo cboStaff, tlOperators
    lngOperatorID = tlOperators.Key(cboStaff.text)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set tlOperators = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub optAll_Click()
    On Error GoTo errHandler
    dteSince = CDate(0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.optAll_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optOD_Click()
    On Error GoTo errHandler
    dteSince = Date
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.optOD_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optOD1M_Click()
    On Error GoTo errHandler
    dteSince = DateAdd("m", -1, Date)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.optOD1M_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optOD1W_Click()
    On Error GoTo errHandler
    dteSince = DateAdd("ww", -1, Date)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.optOD1W_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optOD2W_Click()
    On Error GoTo errHandler
    dteSince = DateAdd("ww", -2, Date)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.optOD2W_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optOD3W_Click()
    On Error GoTo errHandler
    dteSince = DateAdd("ww", -2, Date)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoadODCO.optOD3W_Click", , EA_NORERAISE
    HandleError
End Sub

