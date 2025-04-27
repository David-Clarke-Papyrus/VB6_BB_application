VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmBetweendates 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Statistics - snapshots hoistory"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   14205
   Begin VB.CommandButton cmdGrid1toSpreadsheet 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Send to Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   75
      Width           =   1395
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4845
      Picture         =   "frmBetweendates.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   75
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5865
      Picture         =   "frmBetweendates.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   75
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1065
      TabIndex        =   0
      Top             =   165
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   221839361
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   3090
      TabIndex        =   1
      Top             =   180
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   221839361
      CurrentDate     =   37421
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6285
      Left            =   135
      TabIndex        =   6
      Top             =   795
      Width           =   13740
      _ExtentX        =   24236
      _ExtentY        =   11086
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Daily stock on hand and O/S orders"
      TabPicture(0)   =   "frmBetweendates.frx":0714
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "G"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Daily values of movements"
      TabPicture(1)   =   "frmBetweendates.frx":0730
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "GG"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmBetweendates.frx":074C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin TrueOleDBGrid60.TDBGrid G 
         Height          =   5550
         Left            =   -74820
         OleObjectBlob   =   "frmBetweendates.frx":0768
         TabIndex        =   7
         Top             =   480
         Width           =   13335
      End
      Begin TrueOleDBGrid60.TDBGrid GG 
         Height          =   5565
         Left            =   210
         OleObjectBlob   =   "frmBetweendates.frx":78BE
         TabIndex        =   8
         Top             =   495
         Width           =   13335
      End
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
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
      Height          =   270
      Left            =   2505
      TabIndex        =   3
      Top             =   195
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "between"
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
      Height          =   270
      Left            =   90
      TabIndex        =   2
      Top             =   210
      Width           =   840
   End
End
Attribute VB_Name = "frmBetweendates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim bCancelled As Boolean
Dim mdteFrom As Date
Dim mdteTo As Date
Dim x As New XArrayDB
Dim XX As New XArrayDB

Sub component(pdteFrom As Date, pdteTo As Date)
    On Error GoTo errHandler
    Me.dtpFrom = pdteFrom
    mdteFrom = pdteFrom
    Me.dtpTo = pdteTo
    mdteTo = pdteTo
    bCancelled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBetweendates.component(pdteFrom,pdteTo)", Array(pdteFrom, pdteTo)
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
bCancelled = True
Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBetweendates.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdGrid1toSpreadsheet_Click()
Dim fs As New FileSystemObject
Dim strExecutable As String
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder oPC.LocalFolder & "\TEMP"
    End If
    strExecutable = GetPDFExecutable(oPC.SharedFolderRoot & "\TEMPLATES\DUMMY.XLS")
    If strExecutable = "" Then
        MsgBox "Contact support, missing 'DUMMY.XLS' file in \Templates folder, or no application available to open .xls file" & vbCrLf & "Report will not open now but is saved in " & oPC.SharedFolderRoot & "\HTML\SupplierCharts.html", vbInformation, "Status"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If SSTab1.Tab = 0 Then
        If fs.FileExists(oPC.LocalFolder & "\TEMP\Statistics.HTML") Then
            fs.DeleteFile oPC.LocalFolder & "\TEMP\Statistics.HTML"
        End If
        If fs.FileExists(oPC.LocalFolder & "\TEMP\Statistics.HTML") Then
            MsgBox "Cannot delete old version of " & oPC.LocalFolder & "\TEMP\Statistics.HTML. Contact support", vbInformation, "Can't do this"
            Exit Sub
        End If
        G.ExportToFile oPC.LocalFolder & "\TEMP\Statistics.HTML", False
        Shell """" & strExecutable & """" & " " & oPC.LocalFolder & "\TEMP\Statistics.HTML", vbNormalFocus
    ElseIf SSTab1.Tab = 1 Then
        If fs.FileExists(oPC.LocalFolder & "\TEMP\Statistics2.HTML") Then
            fs.DeleteFile oPC.LocalFolder & "\TEMP\Statistics2.HTML"
        End If
        If fs.FileExists(oPC.LocalFolder & "\TEMP\Statistics2.HTML") Then
            MsgBox "Cannot delete old version of " & oPC.LocalFolder & "\TEMP\Statistics2.HTML. Contact support", vbInformation, "Can't do this"
            Exit Sub
        End If
        GG.ExportToFile oPC.LocalFolder & "\TEMP\Statistics2.HTML", False
        Shell """" & strExecutable & """" & " " & oPC.LocalFolder & "\TEMP\Statistics2.HTML", vbNormalFocus
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    bCancelled = False
    LoadGrids
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBetweendates.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub dtpFrom_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    mdteFrom = dtpFrom
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBetweendates.dtpFrom_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub dtpTo_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    mdteTo = dtpTo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBetweendates.dtpTo_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property
Public Property Get DateFrom() As Date
    DateFrom = dtpFrom
End Property
Public Property Get DateTo() As Date
    DateTo = dtpTo
End Property

Private Sub LoadGrids()
    On Error GoTo errHandler
Dim i As Integer
Dim iRow As Integer
Dim iCol As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open "Select * FROM  tSTAT2 WHERE STAT_DATE BETWEEN '" & ReverseDate(mdteFrom) & " ' AND '" & ReverseDate(mdteTo) & "' ORDER BY STAT_DATE DESC", oPC.COShort, adOpenStatic, adLockReadOnly
    x.Clear
    x.ReDim 1, rs.RecordCount, 1, 14
    iRow = 1
    If Not rs.eof Then
        Do While Not rs.eof
                x(iRow, 1) = Trim(rs.fields("STAT_DATE"))
                x(iRow, 2) = Trim(rs.fields("STAT_ValueOfStock_RetailEx"))
                x(iRow, 3) = Trim(rs.fields("STAT_ValueOfStock_Cost"))
                x(iRow, 4) = Trim(rs.fields("STAT_OnHand_Qtyproducts"))
                x(iRow, 5) = Trim(rs.fields("STAT_OnHand_QtyItems"))
                x(iRow, 6) = Trim(rs.fields("STAT_PO_OS_Value_Cost"))
                x(iRow, 7) = Trim(rs.fields("STAT_PO_OS_QtyItems"))
                x(iRow, 8) = Trim(rs.fields("STAT_CO_OS_Value_Cost"))
                x(iRow, 9) = Trim(rs.fields("STAT_CO_OS_QtyItems"))
                x(iRow, 10) = Trim(rs.fields("STAT_Appros_OS_QtyItems"))
                x(iRow, 11) = Trim(rs.fields("STAT_Appros_OS_Value_Cost"))
                
            iRow = iRow + 1
            rs.MoveNext
        Loop
        rs.MoveFirst
        G.Array = x
    End If
    XX.Clear
    XX.ReDim 1, rs.RecordCount, 1, 22
    iRow = 1
    If Not rs.eof Then
        Do While Not rs.eof
                XX(iRow, 1) = Trim(rs.fields("STAT_DATE"))
                XX(iRow, 2) = Trim(rs.fields("STAT_DEL_QtyItems_mm"))
'                XX(iRow, 3) = Trim(rs.Fields("STAT_DEL_Value_Retail"))
                XX(iRow, 3) = Trim(rs.fields("STAT_DEL_Value_Cost_mm"))
                XX(iRow, 4) = Trim(rs.fields("STAT_INV_QtyItems_mm"))
'                XX(iRow, 6) = Trim(rs.Fields("STAT_INV_Value_Retail"))
                XX(iRow, 5) = Trim(rs.fields("STAT_INV_Value_Cost_mm"))
                XX(iRow, 6) = Trim(rs.fields("STAT_CS_QtyItems_mm"))
                XX(iRow, 7) = Trim(rs.fields("STAT_CS_Value_Cost_mm"))
                XX(iRow, 8) = Trim(rs.fields("STAT_PO_QtyItems_mm"))
'                XX(iRow, 12) = Trim(rs.Fields("STAT_PO_Value_Retail_mm"))
                XX(iRow, 9) = Trim(rs.fields("STAT_PO_Value_Cost_mm"))
                XX(iRow, 10) = Trim(rs.fields("STAT_CO_QtyItems_mm"))
'                XX(iRow, 15) = Trim(rs.Fields("STAT_CO_Value_Retail_mm"))
                XX(iRow, 11) = Trim(rs.fields("STAT_CO_Value_Cost_mm"))
                XX(iRow, 12) = Trim(rs.fields("STAT_TFRIN_QtyItems_mm"))
'                XX(iRow, 18) = Trim(rs.Fields("STAT_TFRIN_Value_Retail_mm"))
                XX(iRow, 13) = Trim(rs.fields("STAT_TFRIN_Value_Cost_mm"))
                XX(iRow, 14) = Trim(rs.fields("STAT_TFROUT_QtyItems_mm"))
'                XX(iRow, 21) = Trim(rs.Fields("STAT_TFROUT_Value_Retail_mm"))
                XX(iRow, 15) = Trim(rs.fields("STAT_TFROUT_Value_Cost_mm"))
                
            iRow = iRow + 1
            rs.MoveNext
        Loop
        GG.Array = XX
    End If
'    If Me.WindowState <> 2 Then
'        Me.Height = 6000
'        Me.Width = 10000
'        Me.left = 500
'        Me.top = 1000
'    End If
        G.ReBind
        GG.ReBind
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSTAT.Form_Load"
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBetweendates.LoadGrids"
End Sub
Private Sub SetDefaultWidths()
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To G.Columns.Count
        G.Columns(i - 1).Width = 500
    Next
    For i = 1 To GG.Columns.Count
        GG.Columns(i - 1).Width = 500
    Next

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBetweendates.SetDefaultWidths"
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    SetGridLayout Me.G, Me.Name
    SetGridLayout Me.GG, "GG"
    SetFormSize Me
    SSTab1.Tab = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBetweendates.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    SaveLayout Me.G, Me.Name, Me.Height, Me.Width
    SaveLayout Me.GG, "GG", SSTab1.Height, SSTab1.Width
    x.Clear
    XX.Clear
    Set x = Nothing
    Set XX = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBetweendates.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    SSTab1.Width = NonNegative_Lng(Me.Width - 480)
    SSTab1.Height = NonNegative_Lng(Me.Height - 1540)
    
    G.Width = NonNegative_Lng(SSTab1.Width - 580)
    G.Height = NonNegative_Lng(SSTab1.Height - 800)
    GG.Width = NonNegative_Lng(SSTab1.Width - 580)
    GG.Height = NonNegative_Lng(SSTab1.Height - 800)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBetweendates.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

