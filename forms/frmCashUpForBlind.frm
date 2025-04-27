VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmCashUPForBlind 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Cash-up"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   8340
   Begin VB.Frame Frame6 
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
      Height          =   1770
      Left            =   165
      TabIndex        =   8
      Top             =   2565
      Width           =   4710
      Begin VB.PictureBox Picture 
         BackColor       =   &H00D3D3CB&
         Height          =   645
         Left            =   90
         ScaleHeight     =   585
         ScaleWidth      =   3120
         TabIndex        =   16
         Top             =   1005
         Width           =   3180
         Begin VB.OptionButton optPT 
            BackColor       =   &H00D3D3CB&
            Caption         =   "by product type"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   1485
         End
         Begin VB.OptionButton optSection 
            BackColor       =   &H00D3D3CB&
            Caption         =   "by Section"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   0
            TabIndex        =   19
            Top             =   240
            Width           =   1485
         End
         Begin VB.OptionButton optDetails 
            BackColor       =   &H00D3D3CB&
            Caption         =   "details"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   1560
            TabIndex        =   18
            Top             =   0
            Width           =   1140
         End
         Begin VB.OptionButton optPayments 
            BackColor       =   &H00D3D3CB&
            Caption         =   "payments"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   1560
            TabIndex        =   17
            Top             =   240
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
         Left            =   3660
         Picture         =   "frmCashUpForBlind.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   990
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   2970
      End
      Begin VB.Line Line11 
         DrawMode        =   8  'Xor Pen
         X1              =   15
         X2              =   4630
         Y1              =   705
         Y2              =   705
      End
   End
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
      Height          =   2265
      Left            =   180
      TabIndex        =   6
      Top             =   4455
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
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1815
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1815
         Width           =   2670
      End
      Begin TrueOleDBGrid60.TDBGrid GX 
         Height          =   1140
         Left            =   285
         OleObjectBlob   =   "frmCashUpForBlind.frx":038A
         TabIndex        =   7
         Top             =   330
         Width           =   5775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "double-click row to open cash-up form for till"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   210
         Left            =   270
         TabIndex        =   15
         Top             =   1485
         Width           =   5535
      End
   End
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
      Height          =   1770
      Left            =   4965
      TabIndex        =   3
      Top             =   2565
      Width           =   3270
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
         Height          =   375
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1065
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Use this to look for any possible missing exchanges in a day-session. Select the day-session and click below."
         ForeColor       =   &H8000000D&
         Height          =   705
         Left            =   225
         TabIndex        =   5
         Top             =   300
         Width           =   2895
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
      Height          =   2445
      Left            =   165
      TabIndex        =   1
      Top             =   75
      Width           =   8070
      Begin TrueOleDBGrid60.TDBGrid GZ 
         Height          =   1995
         Left            =   210
         OleObjectBlob   =   "frmCashUpForBlind.frx":4861
         TabIndex        =   2
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
      Left            =   7215
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCashUpForBlind.frx":9488
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print the invoice"
      Top             =   6120
      Width           =   1000
   End
End
Attribute VB_Name = "frmCashUPForBlind"
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
    ErrorIn "frmCashUPForBlind.cmdAudit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.cmdClose_Click", , EA_NORERAISE
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
    ErrorIn "frmCashUPForBlind.G1_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo errHandler
    LoadZSessions
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.cmdRefresh_Click", , EA_NORERAISE
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
    
    rs.Open "SELECT * FROM vBackofficeInvoices WHERE DTE BETWEEN '" & ReverseDate(dteFrom) & "' AND '" & ReverseDate(dteTo) & "'  ORDER BY CODE ", oPC.COShort, adOpenForwardOnly
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
    ErrorIn "frmCashUPForBlind.cmdnonPOS_Click", , EA_NORERAISE
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
    ErrorIn "frmCashUPForBlind.cmdPrintPTs_Click", , EA_NORERAISE
    HandleError
End Sub
'
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
    ErrorIn "frmCashUPForBlind.cmdPrintSections_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrintDaysessions_Click()
    On Error GoTo errHandler
    If XZ.Value(GZ.Bookmark, 9) = 3 Then
        If optPT = True Then
            cmdPrintPTs_Click
        ElseIf optSection = True Then
            cmdPrintSections_Click
        ElseIf optDetails = True Then
            cmdSalesDetails_Click
        ElseIf optPayments = True Then
            PrintPayments
        End If
    Else
        MsgBox "This session has un-issued operator sessions, you cannot view it yet.", vbOKOnly + vbInformation, "Can't do this"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.cmdPrintDaysessions_Click", , EA_NORERAISE
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
    ErrorIn "frmCashUPForBlind.PrintPayments"
End Sub
'
Private Sub cmdPrintXSession_Click()
    On Error GoTo errHandler
 Dim oCU As z_Cashup
 
    If XX.Value(GX.Bookmark, 9) = 3 Then
        Set oCU = New z_Cashup
        If XX.Value(GX.Bookmark, 1) = "" Or XX.Value(GX.Bookmark, 2) = "" Then
    '        MsgBox "Printing provisional sheet (X reading) only", vbInformation, "Warning"
    '
    '        oCU.SelectSession "X", CStr(XX.Value(GX.Bookmark, 6))
    '        oCU.Calculate
    '        oCU.PrintCashup "X", True
    '        Exit Sub
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
    Else
        MsgBox "This session is un-issued, you cannot view it yet.", vbOKOnly + vbInformation, "Can't do this"
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.cmdPrintXSession_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub OpenSession()
    On Error GoTo errHandler
Dim oCU As z_Cashup
Dim oSM As New z_StockManager
Dim str As String
Dim f  As frmBlindCashup

    Screen.MousePointer = vbHourglass
    str = oSM.AuditDaysession(CStr(XZ(GZ.Bookmark, 6)))
    Screen.MousePointer = vbDefault
    If str > "" Then
        If MsgBox("Manually fetch the following exchanges for this session before continuing." & vbCrLf & str, vbInformation + vbYesNo, "List of missing exchanges") = vbYes Then
            GoTo EXIT_Handler
        End If
    End If
    
    If oPC.IsFrontDeskWorkstation Then
        If XX.Value(GX.Bookmark, 2) > "" Then   'If closed we can cash up
            Set f = New frmBlindCashup
            f.component CStr(XX.Value(GX.Bookmark, 6))
            f.Show vbModal
        Else
            MsgBox "This session has not been closed and can not be cashed up.", vbOKOnly, "Can't do this"
            Exit Sub
        End If
    Else
        If XX.Value(GX.Bookmark, 2) > "" Then   'If closed we can cash up
            Set f = New frmBlindCashup
            f.component CStr(XX.Value(GX.Bookmark, 6))
            f.Show vbModal
        Else
            MsgBox "This session has not been closed and can not be cashed up.", vbOKOnly, "Can't do this"
            Exit Sub
        End If
    
    End If

EXIT_Handler:
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.OpenSession"
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
    If XZ.Value(GZ.Bookmark, 9) = 3 Then
        Set oCU = New z_Cashup
        If Not GZ.Bookmark = XZ.UpperBound(1) Then
            If XZ.Value(GZ.Bookmark, 1) = "" Or XZ.Value(GZ.Bookmark, 2) = "" Then
                MsgBox "Cannot display report for open session or where previous session is open.", vbInformation, "Can't do this"
                Exit Sub
            End If
        Else
            If GZ.Bookmark < XZ.UpperBound(1) Then
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
    Else
        MsgBox "This session has un-issued operator sessions, you cannot view it yet.", vbOKOnly + vbInformation, "Can't do this"
    End If
    
EXIT_Handler:
    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.cmdPrintZSession_Click", , EA_NORERAISE
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
    ErrorIn "frmCashUPForBlind.cmdSalesDetails_Click", , EA_NORERAISE
    HandleError
End Sub
'
Private Sub cmdXPaySum_Click()
    On Error GoTo errHandler
Dim frm As New frmPaymentSummary
    
    If XX.Value(GX.Bookmark, 9) = 3 Then
        If XX.Value(GX.Bookmark, 8) > 0 Then
            frm.component "", CStr(XX.Value(GX.Bookmark, 6))
            If frm.RowsToDisplayCount > 0 Then
                frm.Show
            End If
        Else
            MsgBox "This session is un-issued, you cannot view it yet.", vbOKOnly + vbInformation, "Can't do this"
        End If
    Else
        MsgBox "This session is un-issued, you cannot see the summary yet.", vbOKOnly + vbInformation, "Can't do this"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.cmdXPaySum_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        TOP = 50
        Left = 120
        Width = 8500
        Height = 7500
    End If
    LoadZSessions
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.Form_Load", , EA_NORERAISE
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
    XZ.ReDim 1, ocZ.Count, 1, 9
    For i = 1 To ocZ.Count
        XZ.Value(i, 1) = ocZ.Item(i).StartDateF
        XZ.Value(i, 2) = ocZ.Item(i).EndDateF
        XZ.Value(i, 3) = ocZ.Item(i).TillPoint
        XZ.Value(i, 4) = ocZ.Item(i).SupervisorName
        XZ.Value(i, 5) = ocZ.Item(i).NominalDate
        XZ.Value(i, 6) = ocZ.Item(i).ID
        XZ.Value(i, 7) = ocZ.Item(i).StartDateSort
        XZ.Value(i, 8) = ocZ.Item(i).EndDate
        XZ.Value(i, 9) = ocZ.Item(i).Reportable
    Next
    XZ.QuickSort 1, XZ.UpperBound(1), 7, XORDER_DESCEND, XTYPE_STRING
    GZ.Array = XZ
    GZ.ReBind
'    GZ.Bookmark = 0
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.LoadZGrid"
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
    XX.ReDim 1, ocCS.Count, 1, 9
    For i = 1 To ocCS.Count
        XX.Value(i, 1) = ocCS.Item(i).StartDateF
        XX.Value(i, 2) = ocCS.Item(i).EndDateF
        XX.Value(i, 3) = ocCS.Item(i).StaffName
        XX.Value(i, 4) = ocCS.Item(i).TRID
        XX.Value(i, 5) = ocCS.Item(i).StartDateSort
        XX.Value(i, 6) = ocCS.Item(i).CSGUID
        XX.Value(i, 7) = ocCS.Item(i).StartDate
        XX.Value(i, 8) = ocCS.Item(i).EndDate
        XX.Value(i, 9) = ocCS.Item(i).Reportable
    Next
    XX.QuickSort 1, XX.UpperBound(1), 1, XORDER_DESCEND, XTYPE_DATE
    GX.Array = XX
    GX.ReBind
    GX.Bookmark = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.LoadOpsGrid"
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
    ErrorIn "frmCashUPForBlind.ClearOpsGrid"
End Sub

Private Sub LoadZSessions()
    On Error GoTo errHandler
    Set ocZ = New c_ZSession
    ocZ.Load DateAdd("yyyy", -1, Date), CDate(0), 0
    Screen.MousePointer = vbHourglass
    LoadZGrid
   ' RefreshOps
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.LoadZSessions"
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
    ErrorIn "frmCashUPForBlind.RefreshOps"
End Sub


Private Sub GX_DblClick()
    On Error GoTo errHandler
    OpenSession
    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmCashUPForBlind: GX_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmCashUPForBlind: GX_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.GX_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub GX_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If XX.Value(Bookmark, 9) = 0 Then
        RowStyle.BackColor = COLOR_InProcessGreen
    ElseIf XX.Value(Bookmark, 9) < 3 Then
        RowStyle.BackColor = COLOR_CANCELLED
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.GX_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub GZ_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If XZ.Value(Bookmark, 9) = 0 Then
        RowStyle.BackColor = COLOR_InProcessGreen
    ElseIf XZ.Value(Bookmark, 9) < 3 Then
        RowStyle.BackColor = COLOR_CANCELLED
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.GZ_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub

'
'
'Private Sub GZ_DblClick()
'    OpenSession
'End Sub

Private Sub GZ_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
    RefreshOps
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.GZ_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub GZ_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
Dim oSM As z_StockManager
Dim frm As frmCloseSession
Dim dte As Date
Dim bCancelled As Boolean
Dim bIsSUpervisor As Boolean

    If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
        If MsgBox("Do you want to close this session?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
            Exit Sub
        Else
            If SecurityControl(enSECURITY_CONFIG_SIGN, bCancelled, "Enter your signature", "You do not have permission to close session (or your signature is invalid)", bIsSUpervisor) = True Then
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
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.GZ_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub GX_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
Dim oSM As z_StockManager
Dim frm As frmCloseSession
Dim dte As Date
Dim bCancelled As Boolean
Dim bIsSUpervisor As Boolean

    If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
        If MsgBox("Do you want to close this session?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
            Exit Sub
        Else
            If SecurityControl(enSECURITY_CONFIG_SIGN, bCancelled, "Enter your signature", "You do not have permission to close session (or your signature is invalid)", bIsSUpervisor) = True Then
            
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
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashUPForBlind.GX_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub


