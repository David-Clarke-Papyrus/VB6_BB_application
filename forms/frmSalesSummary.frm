VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{7A5C485E-4ACE-4C72-B64D-46119DEDD852}#4.0#0"; "CCubeX40.ocx"
Begin VB.Form frmSalesSummary 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Sales summary"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   12255
   Begin VB.CommandButton cmdFetch 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Fetch"
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
      Left            =   5340
      Picture         =   "frmSalesSummary.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   330
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   345
      Left            =   780
      TabIndex        =   5
      Top             =   360
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483624
      CalendarForeColor=   -2147483635
      CalendarTitleBackColor=   -2147483624
      CalendarTitleForeColor=   -2147483635
      CalendarTrailingForeColor=   -2147483624
      CustomFormat    =   "dd/MM/yyyy HH:MM"
      Format          =   221839363
      CurrentDate     =   38519
      MaxDate         =   73415
      MinDate         =   37987
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4590
      Left            =   180
      TabIndex        =   1
      Top             =   1215
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   8096
      _Version        =   393216
      Tabs            =   2
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
      TabPicture(0)   =   "frmSalesSummary.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "G1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSaveLayout"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Pivot view"
      TabPicture(1)   =   "frmSalesSummary.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CC"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin CCubeX4.ContourCubeX CC 
         Height          =   3600
         Left            =   -74655
         TabIndex        =   2
         Top             =   705
         Width           =   9060
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
         CCubeXMetadata  =   $"frmSalesSummary.frx":03C2
      End
      Begin VB.CommandButton cmdSaveLayout 
         BackColor       =   &H00D7D1BF&
         Caption         =   "Save layout"
         Height          =   345
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3990
         Width           =   975
      End
      Begin TrueOleDBGrid60.TDBGrid G1 
         Height          =   3390
         Left            =   165
         OleObjectBlob   =   "frmSalesSummary.frx":23EE
         TabIndex        =   4
         Top             =   555
         Width           =   10680
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10350
      Picture         =   "frmSalesSummary.frx":8441
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5850
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   345
      Left            =   3165
      TabIndex        =   6
      Top             =   360
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483624
      CalendarForeColor=   -2147483635
      CalendarTitleBackColor=   -2147483624
      CalendarTitleForeColor=   -2147483635
      CalendarTrailingForeColor=   -2147483624
      CustomFormat    =   "dd/MM/yyyy HH:MM"
      Format          =   221839363
      CurrentDate     =   38519
      MaxDate         =   73415
      MinDate         =   37987
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2655
      TabIndex        =   8
      Top             =   375
      Width           =   390
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   165
      TabIndex        =   7
      Top             =   390
      Width           =   525
   End
End
Attribute VB_Name = "frmSalesSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim XA As XArrayDB


Private Sub cmdFetch_Click()
    On Error GoTo errHandler
    Dim lngIndex As Long
    Dim OpenResult As Integer
    Dim oSQL As z_SQL
    Dim dteLimitToView As Date

  If oPC.BlindCashup = True Then
      Set oSQL = New z_SQL
      dteLimitToView = oSQL.GetDateOfEarliestUnSignedSession
      If dtTo >= StartOfDay(dteLimitToView) Then
          MsgBox "There are unsigned cash ups starting prior to your selected end date (" & Format(dteLimitToView, "dd/mm/yyyy") & "). You cannot include thse in the report. Select an earlier end date.", vbInformation, "Can't do this"
          Exit Sub
      End If
  End If
    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT * FROM cuSaleAnalysis WHERE dte Between  '" & ReverseDate(dtFrom) & "' AND '" & ReverseDate(dtTo) & "' ORDER BY EXCHNUM"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    Set XA = New XArrayDB
    XA.ReDim 1, rs.RecordCount, 1, 10
    lngIndex = 1
    If Not rs.eof Then
        Do While Not rs.eof
            XA.Value(lngIndex, 1) = FNN(rs.fields("EXCHNUM"))
            XA.Value(lngIndex, 2) = FNS(rs.fields("SP"))
            XA.Value(lngIndex, 3) = Format(FND(rs.fields("DTE2")), "dd mmm hh:nn") 'FND(rs.Fields("DTE2"))
            XA.Value(lngIndex, 4) = FNS(rs.fields("NAME"))
            XA.Value(lngIndex, 5) = FNDBL(rs.fields("Ext"))
            XA.Value(lngIndex, 6) = FNDBL(rs.fields("PriceAlt"))
            XA.Value(lngIndex, 7) = FNS(rs.fields("MaxDisc"))
            XA.Value(lngIndex, 8) = IIf(FNN(rs.fields("VOIDED")) = 1, "VOIDED", "")
            XA.Value(lngIndex, 9) = FNN(rs.fields("VOIDING"))
            lngIndex = lngIndex + 1
            rs.MoveNext
        Loop
        G1.Array = XA
        G1.ReBind
        G1.Refresh
        DoEvents
        
        Me.Refresh
        
        CC.Cube.Active = False
        CC.Cube.DataSourceType = xcdt_Recordset
        rs.MoveFirst
        CC.Cube.open rs
        CC.Cube.Active = True
    Else
        G1.Array = XA
        G1.ReBind
        G1.Refresh
        CC.Active = False
        MsgBox "No records", , "Status"
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSalesSummary.cmdFetch_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
      If Err = -2147417848 Then
          MsgBox "The connection to the database has been interrupted. Please reload the form.", vbInformation + vbOKOnly, "Can't do this"
          Err.Clear
          Exit Sub
      End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSummary.cmdFetch_Click", , EA_NORERAISE, , "line number", Array(Erl)
    HandleError
End Sub
Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSalesSummary.cmdClose_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSummary.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdSaveLayout_Click()
    On Error GoTo errHandler
    SaveLayout Me.G1, Me.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdSaveLayout_Click", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSalesSummary.cmdSaveLayout_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSummary.cmdSaveLayout_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub Form_Load()
    On Error GoTo errHandler
Dim i As Integer
    If Me.WindowState <> 2 Then
        Me.TOP = 200
        Left = 200
        Width = 12500
        Height = 6900
    End If
    Me.dtTo = DateAdd("d", 1, Date)
    Me.dtFrom = DateAdd("ww", -1, dtTo)
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", "frmSaleAnalysis", CStr(i), G1.Columns(i - 1).Width)
    Next
    
   
    CC.Cube.Dims.Add "SP", "SP", , xda_vertical
    CC.Cube.Dims.Add "EXCHNUM", "EXCHNUM", , xda_vertical
    CC.Cube.Dims.Add "Dte", "Dte", , xda_vertical
    CC.Cube.Dims.Add "VOIDED", "VOIDED", , xda_vertical
    
    CC.Cube.BaseFacts.Add "EXT"
    CC.Cube.BaseFacts.Add "PRICEALT"
    CC.Cube.BaseFacts.Add "MaxDisc"
    
    CC.Cube.Facts.Add "EXT", "EXT", xfaa_SUM
    CC.Cube.Facts.Add "PRICEALT", "PRICEALT", xfaa_SUM
    CC.Cube.Facts.Add "MaxDisc", "MaxDisc", xfaa_SUM
    
    CC.Facts("EXT").Visible = True
    CC.Facts("PRICEALT").Visible = True
    CC.Facts("MaxDisc").Visible = True

    
    CC.Active = False
    
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSalesSummary.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSummary.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
        On Error Resume Next
    If Not rs Is Nothing Then
     rs.Close
    End If
    Set rs = Nothing

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSalesSummary.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSummary.Form_Unload(Cancel)", Cancel, EA_NORERAISE
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
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 3, XORDER_ASCEND, XTYPE_DATE  'XTYPE_INTEGER
    
    G1.Refresh
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSalesSummary.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSummary.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSalesSummary.GetRowType(ColIndex)", ColIndex
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSummary.GetRowType(ColIndex)", ColIndex
End Function

