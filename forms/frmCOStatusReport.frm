VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmCustomerOrderStatusReport 
   Caption         =   "Customer order status report"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   12735
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
      Height          =   600
      Left            =   10830
      Picture         =   "frmCOStatusReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4695
      Width           =   1410
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Print"
      Height          =   600
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1410
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Send"
      Height          =   600
      Left            =   5655
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1410
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   4425
      Left            =   195
      OleObjectBlob   =   "frmCOStatusReport.frx":038A
      TabIndex        =   0
      Top             =   225
      Width           =   12195
   End
End
Attribute VB_Name = "frmCustomerOrderStatusReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngTRID As Long
Dim lngCOSTRID As Long
Dim oSM As New z_StockManager
Dim rs As ADODB.Recordset
Dim XA As New XArrayDB

Public Sub component(TRID As Long, pHeading As String)
    On Error GoTo errHandler
    Me.Caption = pHeading
    lngTRID = TRID
    oSM.GetCOSRLines lngTRID, rs
    LoadGrid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerOrderStatusReport.component(TRID,pHeading)", Array(TRID, pHeading)
End Sub

Private Sub cmdNew_Click()
    On Error GoTo errHandler
    oSM.CreateNewCOStatusReport lngTRID, lngCOSTRID, rs
    LoadGrid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerOrderStatusReport.cmdNew_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerOrderStatusReport.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    Me.G1.PrintInfo.PageSetup
    Me.G1.PrintInfo.PageHeader = Me.Caption
    Me.G1.PrintInfo.SettingsPaperSize = 9
    Me.G1.PrintInfo.SettingsOrientation = 2
    Me.G1.PrintInfo.SettingsMarginLeft = 1000
    Me.G1.PrintInfo.SettingsMarginRight = 1000
    Me.G1.PrintInfo.PrintPreview
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerOrderStatusReport.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSend_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager

    oSM.GeneratOrderStatusTransmission rs
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerOrderStatusReport.cmdSend_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    SetGridLayout Me.G1, Me.Name
    SetFormSize Me
   ' LoadGrid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerOrderStatusReport.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadGrid()
    On Error GoTo errHandler
Dim i As Integer

    i = 0
    Do While Not rs.eof
        i = i + 1
        XA.ReDim 1, i, 1, 10
        XA(i, 1) = FNS(rs.fields("CODEF"))
        XA(i, 2) = FNS(rs.fields("Title"))
        XA(i, 3) = FNS(rs.fields("COSRL_CustomerOrderRef"))
        XA(i, 4) = FNN(rs.fields("QtyOrdered"))
        XA(i, 5) = FNN(rs.fields("QtyDispatched"))
        XA(i, 6) = FNN(rs.fields("QtyOrdered")) - FNN(rs.fields("QtyDispatched"))
        XA(i, 7) = FNS(rs.fields("Availability"))
        XA(i, 8) = FNS(rs.fields("Action"))
        rs.MoveNext
    Loop
    G1.Array = XA
    G1.ReBind
    G1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerOrderStatusReport.LoadGrid"
End Sub


Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    G1.Width = NonNegative_Lng(Me.Width - (G1.Left + 400))
    lngDiff = G1.Height
    G1.Height = NonNegative_Lng(Me.Height - (G1.TOP + 1220))
    lngDiff = (G1.Height - lngDiff)
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    cmdSend.TOP = cmdSend.TOP + lngDiff
    cmdclose.TOP = cmdclose.TOP + lngDiff
    cmdclose.Left = NonNegative_Lng(G1.Width - 1440)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerOrderStatusReport.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    SaveLayout Me.G1, Me.Name, Me.Height, Me.Width
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerOrderStatusReport.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub G1_DblClick()
    On Error GoTo errHandler
Dim frm As New frmPublishersStatusUpdate

    If IsNull(G1.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    frm.component XA(G1.Bookmark, 2), XA(G1.Bookmark, 3)
    frm.Show vbModal
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerOrderStatusReport.G1_DblClick", , EA_NORERAISE
    HandleError
End Sub
