VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmCashUP 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Cash-up"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   8355
   Begin VB.Frame Frame3 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Operator-sessions"
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
      Height          =   2085
      Left            =   165
      TabIndex        =   6
      Top             =   4170
      Width           =   6300
      Begin VB.CommandButton cmdXPaySum 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Print payment &summary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1635
         Width           =   2265
      End
      Begin VB.CommandButton cmdPrintXSession 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Print operator-session summary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1635
         Width           =   2670
      End
      Begin TrueOleDBGrid60.TDBGrid GX 
         Height          =   1275
         Left            =   270
         OleObjectBlob   =   "frmCashUp.frx":0000
         TabIndex        =   7
         Top             =   330
         Width           =   5775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Day-Sessions"
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
      Height          =   3990
      Left            =   165
      TabIndex        =   1
      Top             =   90
      Width           =   8070
      Begin VB.Frame Frame4 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Audit exchanges"
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
         Height          =   1650
         Left            =   4995
         TabIndex        =   11
         Top             =   2205
         Width           =   2940
         Begin VB.CommandButton cmdAudit 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Audit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1770
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1110
            Width           =   1035
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Use this to look for any possible missing exchanges in a day-session. Select the day-session and click below."
            ForeColor       =   &H8000000D&
            Height          =   810
            Left            =   225
            TabIndex        =   12
            Top             =   300
            Width           =   2535
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Reports (for selected day-session)"
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
         Height          =   1635
         Left            =   225
         TabIndex        =   2
         Top             =   2220
         Width           =   4710
         Begin VB.PictureBox Picture 
            BackColor       =   &H00D3D3CB&
            Height          =   585
            Left            =   90
            ScaleHeight     =   525
            ScaleWidth      =   3105
            TabIndex        =   15
            Top             =   1005
            Width           =   3165
            Begin VB.OptionButton optPT 
               BackColor       =   &H00D3D3CB&
               Caption         =   "by product type"
               ForeColor       =   &H8000000D&
               Height          =   270
               Left            =   180
               TabIndex        =   19
               Top             =   15
               Width           =   1485
            End
            Begin VB.OptionButton optSection 
               BackColor       =   &H00D3D3CB&
               Caption         =   "by Section"
               ForeColor       =   &H8000000D&
               Height          =   270
               Left            =   180
               TabIndex        =   18
               Top             =   255
               Width           =   1485
            End
            Begin VB.OptionButton optDetails 
               BackColor       =   &H00D3D3CB&
               Caption         =   "details"
               ForeColor       =   &H8000000D&
               Height          =   270
               Left            =   1740
               TabIndex        =   17
               Top             =   15
               Width           =   1140
            End
            Begin VB.OptionButton optPayments 
               BackColor       =   &H00D3D3CB&
               Caption         =   "payments"
               ForeColor       =   &H8000000D&
               Height          =   270
               Left            =   1740
               TabIndex        =   16
               Top             =   255
               Width           =   1140
            End
         End
         Begin VB.CommandButton cmdPrintDaysessions 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Print"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   3360
            Picture         =   "frmCashUp.frx":44D7
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   900
            Width           =   765
         End
         Begin VB.CommandButton cmdnonPOS 
            BackColor       =   &H00C4BCA4&
            Caption         =   "Print &non-POS invoices"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2370
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   270
            Width           =   2250
         End
         Begin VB.CommandButton cmdPrintZSession 
            BackColor       =   &H00C4BCA4&
            Caption         =   "Print  summary"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   105
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   270
            Width           =   2250
         End
         Begin VB.CheckBox chkWholeDay 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00D3D3CB&
            Caption         =   "Print all Day-sessions for selected day"
            ForeColor       =   &H8000000D&
            Height          =   330
            Left            =   90
            TabIndex        =   3
            Top             =   735
            Visible         =   0   'False
            Width           =   2970
         End
         Begin VB.Line Line1 
            DrawMode        =   8  'Xor Pen
            X1              =   15
            X2              =   4630
            Y1              =   720
            Y2              =   720
         End
      End
      Begin TrueOleDBGrid60.TDBGrid GZ 
         Height          =   1845
         Left            =   210
         OleObjectBlob   =   "frmCashUp.frx":4861
         TabIndex        =   5
         Top             =   300
         Width           =   7725
      End
   End
   Begin VB.CommandButton cmdclose 
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
      Left            =   7185
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCashUp.frx":9488
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print the invoice"
      Top             =   5565
      Width           =   1000
   End
End
Attribute VB_Name = "frmCashUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XE As XArrayDB
Dim XX As XArrayDB
Dim XZ As XArrayDB
Dim XCSL As XArrayDB
Dim XPAY As XArrayDB
Dim rs As ADODB.Recordset
Dim rsZ As ADODB.Recordset
Dim OPSID As Variant
Dim ocZ As c_ZSession
Dim ocCS As c_CSs
Dim ocEX As c_Exchanges


Private Sub cmdAudit_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
Dim str As String
    
    Screen.MousePointer = vbHourglass
    str = oSM.AuditDaysession(CStr(XZ(GZ.Bookmark, 6)))
    Screen.MousePointer = vbDefault
    If str > "" Then
        MsgBox str, vbInformation, "List of missing exchanges"
    Else
        MsgBox "There are no missing exchanges", vbInformation, "List of missing exchanges"
    End If
        
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.cmdAudit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub G1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuReserveList   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.G1_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo errHandler
    LoadZSessions
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.cmdRefresh_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub cmdnonPOS_Click()
    On Error GoTo errHandler
Dim rpt As New arNonPOSTRansactions
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
Dim dteFrom As Date
Dim dteTo As Date

    dteFrom = StartOfDay(CDate(XZ(GZ.Bookmark, 1)))
    dteTo = EndOfDay(CDate(XZ(GZ.Bookmark, 1)))

    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    rs.open "SELECT * FROM vBackofficeInvoices WHERE DTE BETWEEN '" & ReverseDate(dteFrom) & "' AND '" & ReverseDate(dteTo) & "'  ORDER BY CODE ", oPC.COShort, adOpenForwardOnly
    rpt.component rs, "Invoices and credit notes issued from backoffice between " & Format(dteFrom, "DD/MM/YYYY") & " and " & Format(dteTo, "DD/MM/YYYY")
    rpt.Show
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.cmdnonPOS_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrintPTs_Click()
    On Error GoTo errHandler
Dim oCU As z_Cashup
    Set oCU = New z_Cashup
    If chkWholeDay = 1 Then
        oCU.PrintPTSales CDate(XZ(GZ.Bookmark, 1)), ""
    Else
        oCU.PrintPTSales CDate(0), CStr(XZ(GZ.Bookmark, 6))
    End If
    Set oCU = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.cmdPrintPTs_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrintSections_Click()
    On Error GoTo errHandler
Dim oCU As z_Cashup
    Set oCU = New z_Cashup
    If chkWholeDay = 1 Then
        oCU.PrintSectionSales CDate(XZ(GZ.Bookmark, 1)), ""
    Else
        oCU.PrintSectionSales CDate(0), CStr(XZ(GZ.Bookmark, 6))
    End If
    Set oCU = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.cmdPrintSections_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrintDaysessions_Click()
    On Error GoTo errHandler
    If optPT = True Then
        cmdPrintPTs_Click
    ElseIf optSection = True Then
        cmdPrintSections_Click
    ElseIf optDetails = True Then
        cmdSalesDetails_Click
    ElseIf optPayments = True Then
        PrintPayments
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.cmdPrintDaysessions_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub PrintPayments()
    On Error GoTo errHandler
Dim oCU As z_Cashup
    Set oCU = New z_Cashup
    If chkWholeDay = 1 Then
        oCU.PrintPayments CDate(XZ(GZ.Bookmark, 1)), CDate(XZ(GZ.Bookmark, 1)), ""
    Else
        oCU.PrintPayments CStr(XZ(GZ.Bookmark, 1)), CDate(0), CStr(XZ(GZ.Bookmark, 6))
    End If
    Set oCU = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.PrintPayments"
End Sub

Private Sub cmdPrintXSession_Click()
    On Error GoTo errHandler
 Dim oCU As z_Cashup

    Set oCU = New z_Cashup
    If XX.Value(GX.Bookmark, 1) = "" Or XX.Value(GX.Bookmark, 2) = "" Then
        MsgBox "Printing provisional sheet (X reading) only", vbInformation, "Warning"
        
        oCU.SelectSession "X", CStr(XX.Value(GX.Bookmark, 6))
        oCU.Calculate
        oCU.PrintCashup "X", True
        Exit Sub
    Else
        If GX.Bookmark < XX.UpperBound(1) Then
            If IsDate(XX.Value(GX.Bookmark + 1, 1)) And IsDate(XX.Value(GX.Bookmark + 1, 2)) Then
                oCU.component CDate(XX.Value(GX.Bookmark, 1)), CDate(XX.Value(GX.Bookmark, 2)), CStr(XX.Value(GX.Bookmark, 3)), CStr(XZ.Value(GZ.Bookmark, 3)), _
                CDate(XX.Value(GX.Bookmark + 1, 1)), CDate(XX.Value(GX.Bookmark + 1, 2))
            Else
                oCU.component CDate(XX.Value(GX.Bookmark, 1)), CDate(XX.Value(GX.Bookmark, 2)), CStr(XX.Value(GX.Bookmark, 3)), CStr(XZ.Value(GZ.Bookmark, 3)), CDate(0), CDate(0)
            End If
        Else
            oCU.component CDate(XX.Value(GX.Bookmark, 1)), CDate(XX.Value(GX.Bookmark, 2)), CStr(XX.Value(GX.Bookmark, 3)), CStr(XZ.Value(GZ.Bookmark, 3)), CDate(0), CDate(0)
        End If
    End If
    oCU.SelectSession "X", CStr(XX.Value(GX.Bookmark, 6))
    oCU.Calculate
    oCU.PrintCashup "X", False

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.cmdPrintXSession_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrintZSession_Click()
    On Error GoTo errHandler
Dim oCU As z_Cashup
Dim oSM As New z_StockManager
Dim str As String
    
    Screen.MousePointer = vbHourglass
    str = oSM.AuditDaysession(CStr(XZ(GZ.Bookmark, 6)))
    Screen.MousePointer = vbDefault
    If str > "" Then
        MsgBox "Manually fetch the following exchanges for this session before continuing." & vbCrLf & str, vbInformation, "List of missing exchanges"
        GoTo EXIT_Handler
    End If
        

    Set oCU = New z_Cashup
    If Not GZ.Bookmark = XZ.UpperBound(1) Then
        If XZ.Value(GZ.Bookmark, 1) = "" Or XZ.Value(GZ.Bookmark, 2) = "" Then
            MsgBox "Cannot display report for open session or where previous session is open.", vbInformation, "Can't do this"
            Exit Sub
        End If
    Else
        If XZ.UpperBound(1) > GZ.Bookmark Then
            If XZ.Value(GZ.Bookmark, 1) = "" Or XZ.Value(GZ.Bookmark, 2) = "" Or XZ.Value(GZ.Bookmark + 1, 1) = "" Or XZ.Value(GZ.Bookmark + 1, 2) = "" Then
                MsgBox "Cannot display report for open session or where previous session is open.", vbInformation, "Can't do this"
                Exit Sub
            End If
        End If
    End If
    If GZ.Bookmark < XZ.UpperBound(1) Then
        If IsDate(XZ.Value(GZ.Bookmark + 1, 7)) And IsDate(XZ.Value(GZ.Bookmark + 1, 8)) Then
            oCU.component CDate(XZ.Value(GZ.Bookmark, 7)), CDate(XZ.Value(GZ.Bookmark, 8)), CStr(XZ.Value(GZ.Bookmark, 4)), CStr(XZ.Value(GZ.Bookmark, 3)), _
            CDate(XZ.Value(GZ.Bookmark + 1, 7)), CDate(XZ.Value(GZ.Bookmark + 1, 8))
        Else
            oCU.component CDate(XZ.Value(GZ.Bookmark, 7)), CDate(XZ.Value(GZ.Bookmark, 8)), CStr(XZ.Value(GZ.Bookmark, 4)), CStr(XZ.Value(GZ.Bookmark, 3)), CDate(0), CDate(0)
        End If
    Else
        oCU.component CDate(XZ.Value(GZ.Bookmark, 7)), CDate(XZ.Value(GZ.Bookmark, 8)), CStr(XZ.Value(GZ.Bookmark, 3)), CStr(XZ.Value(GZ.Bookmark, 2)), CDate(0), CDate(0)
    End If
    oCU.SelectSession "Z", CStr(XZ.Value(GZ.Bookmark, 6))
    oCU.Calculate
    oCU.PrintCashup "Z"

EXIT_Handler:
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.cmdPrintZSession_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdSalesDetails_Click()
    On Error GoTo errHandler
Dim oCU As z_Cashup
    Set oCU = New z_Cashup
    If chkWholeDay = 1 Then
        oCU.PrintDailySales CDate(XZ(GZ.Bookmark, 1)), ""
    Else
        oCU.PrintDailySales CDate(0), CStr(XZ(GZ.Bookmark, 6))
    End If
    Set oCU = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.cmdSalesDetails_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdXPaySum_Click()
    On Error GoTo errHandler
Dim frm As New frmPaymentSummary
    frm.component "", CStr(XX.Value(GX.Bookmark, 6))
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.cmdXPaySum_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Command1_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.Command1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        TOP = 50
        Left = 120
        Width = 8500
        Height = 6800
    End If
    LoadZSessions
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadZGrid()
    On Error GoTo errHandler
Dim objItem As d_ZSession
Dim itmList As ListItem
Dim lngIndex As Long
Dim lngArrayRows As Long
Dim i As Integer
    Set XZ = New XArrayDB
    XZ.Clear
    XZ.ReDim 1, ocZ.Count, 1, 8
    For i = 1 To ocZ.Count
        XZ.Value(i, 1) = ocZ.Item(i).StartDateF
        XZ.Value(i, 2) = ocZ.Item(i).EndDateF
        XZ.Value(i, 3) = ocZ.Item(i).TillPoint
        XZ.Value(i, 4) = ocZ.Item(i).SupervisorName
        XZ.Value(i, 5) = ocZ.Item(i).NominalDate
        XZ.Value(i, 6) = ocZ.Item(i).ID
        XZ.Value(i, 7) = ocZ.Item(i).StartDateSort
        XZ.Value(i, 8) = ocZ.Item(i).EndDate
    Next
    XZ.QuickSort 1, XZ.UpperBound(1), 7, XORDER_DESCEND, XTYPE_STRING
    GZ.Array = XZ
    GZ.ReBind
'    GZ.Bookmark = 0
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.LoadZGrid"
End Sub

Private Sub LoadOpsGrid()
    On Error GoTo errHandler
Dim objItem As d_CS
Dim itmList As ListItem
Dim lngIndex As Long
Dim lngArrayRows As Long
Dim i As Integer
    Set XX = New XArrayDB
    XX.Clear
    XX.ReDim 1, ocCS.Count, 1, 8
    For i = 1 To ocCS.Count
        XX.Value(i, 1) = ocCS.Item(i).StartDateF
        XX.Value(i, 2) = ocCS.Item(i).EndDateF
        XX.Value(i, 3) = ocCS.Item(i).StaffName
        XX.Value(i, 4) = ocCS.Item(i).TRID
        XX.Value(i, 5) = ocCS.Item(i).StartDateSort
        XX.Value(i, 6) = ocCS.Item(i).CSGUID
        XX.Value(i, 7) = ocCS.Item(i).StartDate
        XX.Value(i, 8) = ocCS.Item(i).EndDate
    Next
    XX.QuickSort 1, XX.UpperBound(1), 1, XORDER_DESCEND, XTYPE_DATE
    GX.Array = XX
    GX.ReBind
    GX.Bookmark = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.LoadOpsGrid"
End Sub

Private Sub ClearOpsGrid()
    On Error GoTo errHandler
    If Not XX Is Nothing Then
        XX.Clear
        XX.ReDim 0, 0, 1, 6
    End If
    GX.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.ClearOpsGrid"
End Sub

Private Sub LoadZSessions()
    On Error GoTo errHandler
    Set ocZ = New c_ZSession
    ocZ.Load DateAdd("yyyy", -1, Date), CDate(0), 0
    Screen.MousePointer = vbHourglass
    LoadZGrid
    RefreshOps
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.LoadZSessions"
End Sub
Private Sub RefreshOps()
    On Error GoTo errHandler
    If IsNull(GZ.Bookmark) Then Exit Sub
    If Not ocCS Is Nothing Then ClearOpsGrid
    Set ocCS = New c_CSs
    If Not XZ(GZ.Bookmark, 6) = Empty Then
        ocCS.LoadByZID XZ(GZ.Bookmark, 6)
        LoadOpsGrid
    Else
        ClearOpsGrid
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.RefreshOps"
End Sub




Private Sub GZ_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
    RefreshOps
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.GZ_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub GZ_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
Dim oSM As z_StockManager
Dim frm As frmCloseSession
Dim dte As Date

    If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
        If MsgBox("Do you want to close this session?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
            Exit Sub
        Else
            Set frm = New frmCloseSession
            frm.Show vbModal
            If Not frm.Cancelled Then
                dte = frm.CloseTime
                Unload frm
                Set oSM = New z_StockManager
                oSM.CloseDaysession CStr(XZ(GZ.Bookmark, 6)), ReverseDateTime(dte)
                Set oSM = Nothing
                LoadZSessions
            End If
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.GZ_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), EA_NORERAISE
    HandleError
End Sub

Private Sub GX_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
Dim oSM As z_StockManager
Dim frm As frmCloseSession
Dim dte As Date

    If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
        If MsgBox("Do you want to close this session?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
            Exit Sub
        Else
            Set frm = New frmCloseSession
            frm.Show vbModal
            If Not frm.Cancelled Then
                dte = frm.CloseTime
                Unload frm
                Set oSM = New z_StockManager
                oSM.CloseOperatorsession CStr(XX(GX.Bookmark, 6)), ReverseDateTime(dte)
                Set oSM = Nothing
                RefreshOps
            End If
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUP.GX_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), EA_NORERAISE
    HandleError
End Sub

