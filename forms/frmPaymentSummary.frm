VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{7A5C485E-4ACE-4C72-B64D-46119DEDD852}#4.0#0"; "CCubeX40.ocx"
Begin VB.Form frmPaymentSummary 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Payment summary"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   12255
   Begin TabDlg.SSTab SSTab1 
      Height          =   4590
      Left            =   165
      TabIndex        =   1
      Top             =   225
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   8096
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   617
      TabMaxWidth     =   3175
      ShowFocusRect   =   0   'False
      BackColor       =   14537420
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "List view"
      TabPicture(0)   =   "frmPaymentSummary.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdSaveLayout"
      Tab(0).Control(1)=   "G1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Pivot view"
      TabPicture(1)   =   "frmPaymentSummary.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "CC"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin CCubeX4.ContourCubeX CC 
         Height          =   3960
         Left            =   180
         TabIndex        =   2
         Top             =   495
         Width           =   11640
         Active          =   0   'False
         Transposed      =   0   'False
         NULLValueString =   ""
         Descending      =   0   'False
         NoTotals        =   0   'False
         NoGrandTotals   =   0   'False
         Caption         =   ""
         BackColor       =   13882315
         Enabled         =   -1  'True
         Alive           =   0   'False
         BorderStyle     =   1
         AllowDimOutside =   -1  'True
         AllowExpand     =   -1  'True
         AllowPivot      =   -1  'True
         TotalsString    =   "Totals"
         InactiveDimAreaBkColor=   13160660
         AutoSize        =   0   'False
         UnusedDataAreaColor=   16777215
         MousePointer    =   0
         Object.Visible         =   -1  'True
         InfoURL         =   "http://www.contourcomponents.com/contourcube_user_guide.htm"
         UseThemes       =   0   'False
         WordWrap        =   -1  'True
         FlatStyle       =   0
         FactsVAlignment =   0
         UnusedTreeAreaColor=   16645369
         DimLevelGradient=   14007466
         TreeLineColor   =   14007466
         DimLevelGradientStep=   20
         AllowDimVertical=   -1  'True
         AllowDimHorizontal=   -1  'True
         DrawOptions     =   7
         ConnectionString=   ""
         DataSourceType  =   0
         VERSION_NO      =   2
         CCubeXMetadata  =   $"frmPaymentSummary.frx":0038
      End
      Begin VB.CommandButton cmdSaveLayout 
         BackColor       =   &H00D7D1BF&
         Caption         =   "Save layout"
         Height          =   345
         Left            =   -74850
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4080
         Width           =   975
      End
      Begin TrueOleDBGrid60.TDBGrid G1 
         Height          =   3390
         Left            =   -74835
         OleObjectBlob   =   "frmPaymentSummary.frx":2064
         TabIndex        =   4
         Top             =   555
         Width           =   11730
      End
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
      Left            =   11190
      Picture         =   "frmPaymentSummary.frx":857F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4815
      Width           =   1000
   End
End
Attribute VB_Name = "frmPaymentSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim XA As XArrayDB
Dim OpenResult As Integer

Public Sub component(pZID As String, pXID As String)
    On Error GoTo errHandler
Dim lngIndex As Long
 
    If pZID > "" Then
        strSQL = "SELECT * FROM cuzPaymentsSummary WHERE ZID = '" & pZID & "' ORDER BY EXCHNUM"
    ElseIf pXID > "" Then
        strSQL = "SELECT * FROM cuzPaymentsSummary WHERE XID = '" & pXID & "' ORDER BY EXCHNUM"
    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    Set XA = New XArrayDB
    XA.ReDim 1, rs.RecordCount, 1, 10
    lngIndex = 1
    Do While Not rs.eof
        XA.Value(lngIndex, 1) = FNN(rs.fields("EXCHNUM"))
        XA.Value(lngIndex, 2) = FNS(rs.fields("SP"))
        XA.Value(lngIndex, 3) = FND(rs.fields("DTE"))
        XA.Value(lngIndex, 4) = FNS(rs.fields("NAME"))
        XA.Value(lngIndex, 5) = FNDBL(rs.fields("LOY"))
        XA.Value(lngIndex, 6) = FNDBL(rs.fields("CHANGE"))
        XA.Value(lngIndex, 7) = FNS(rs.fields("MODE"))
        XA.Value(lngIndex, 8) = FNDBL(rs.fields("AMT"))
        XA.Value(lngIndex, 9) = FNS(rs.fields("REF"))
        XA.Value(lngIndex, 10) = FNS(rs.fields("EXTRA"))
        lngIndex = lngIndex + 1
        rs.MoveNext
    Loop
    If Not rs.RecordCount = 0 Then
        rs.MoveFirst
    Else
        MsgBox "No records", vbInformation, "Status"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentSummary.component(pZID,pXID)", Array(pZID, pXID)
End Sub
Private Sub cmdClose_Click()
    On Error GoTo errHandler

    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentSummary.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSaveLayout_Click()
    On Error GoTo errHandler
    SaveLayout Me.G1, Me.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdSaveLayout_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentSummary.cmdSaveLayout_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim i As Integer
    If Me.WindowState <> 2 Then
        Me.TOP = 400
        Left = 400
        Width = 12700
        Height = 6000
    End If
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", "frmPaymentSummary", CStr(i), G1.Columns(i - 1).Width)
    Next
    G1.Array = XA
    G1.Refresh
    
    CC.Cube.Dims.Add "Mode", "Mode"
    CC.Cube.Dims.Add "EXCHNUM", "ExchNum"   ', xda_vertical, 2
    
    CC.Cube.BaseFacts.Add "Amt"
    CC.Cube.BaseFacts.Add "Change"

    CC.Cube.Facts.Add "Amt", "Amt", xfaa_SUM '    '"Amt", "Amt", xfaa_SUM, "Value"
    CC.Cube.Facts.Add "Change", "Change", xfaa_SUM ', "Change"
    
    CC.Facts("Amt").Visible = True
    CC.Facts("Change").Visible = True
    
    CC.Active = False
    
    DoEvents
    Screen.MousePointer = vbHourglass
    CC.Cube.DataSourceType = xcdt_Recordset
    If rs.eof Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    rs.MoveFirst
    If Not rs.eof Then
        CC.Cube.open rs
        CC.Cube.Active = True
    Else
        MsgBox "No records", , "Status"
    End If
    Screen.MousePointer = vbDefault
   
    Me.Refresh
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentSummary.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
        On Error Resume Next
Dim lngDiff As Long
   ' G1.Width = Me.Width - (G1.left + 400)
    If SSTab1.Tab = 0 Then
        CC.Visible = False
        G1.Visible = True
        cmdSaveLayout.Visible = True
    Else
        CC.Visible = True
        G1.Visible = False
        cmdSaveLayout.Visible = False
    End If
    CC.Left = Me.Left + 10
    G1.Left = Me.Left + 10
    CC.Width = NonNegative_Lng(Me.Width - 1400)
    G1.Width = NonNegative_Lng(Me.Width - 1400)
    Me.SSTab1.Width = NonNegative_Lng(Me.Width - 500)
    Me.SSTab1.Height = NonNegative_Lng(Me.Height - 1400)
    lngDiff = G1.Height
    G1.Height = NonNegative_Lng(Me.Height - 2560)
    CC.Height = NonNegative_Lng(Me.Height - 2160)
    lngDiff = (G1.Height - lngDiff)
    cmdclose.TOP = NonNegative_Lng(Me.TOP + Me.Height - 1460)
    cmdclose.Left = NonNegative_Lng(Me.SSTab1.Width - Me.SSTab1.Left - 700)
    Me.cmdSaveLayout.TOP = NonNegative_Lng(SSTab1.TOP + SSTab1.Height - 200)
    cmdSaveLayout.Left = Me.SSTab1.Left

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentSummary.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If rs.State = 1 Then
        rs.Close
    End If
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentSummary.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub G1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    If ColIndex = 2 Then
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1)
    Else
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 3, XORDER_ASCEND, XTYPE_DATE
    End If
    
    G1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentSummary.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 5, 6, 8
            GetRowType = XTYPE_NUMBER
        Case 3
            GetRowType = XTYPE_DATE
        Case Else
            GetRowType = XTYPE_STRING
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentSummary.GetRowType(ColIndex)", ColIndex
End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
    On Error GoTo errHandler
    If SSTab1.Tab = 0 Then
        CC.Visible = False
        G1.Visible = True
        cmdSaveLayout.Visible = True
    Else
        CC.Visible = True
        G1.Visible = False
        cmdSaveLayout.Visible = False
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentSummary.SSTab1_Click(PreviousTab)", PreviousTab, EA_NORERAISE
    HandleError
End Sub

Private Sub SSTab1_DblClick()
    On Error GoTo errHandler
'    If SSTab1.Tab = 0 Then
'        CC.Visible = False
'        G1.Visible = True
'    Else
'        CC.Visible = True
'        G1.Visible = False
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentSummary.SSTab1_DblClick", , EA_NORERAISE
    HandleError
End Sub
Public Property Get RowsToDisplayCount() As Long
    RowsToDisplayCount = CLng(XA.UpperBound(1))
End Property
