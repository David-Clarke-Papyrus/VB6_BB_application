VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowseScanFiles 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Review scanned items"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShowALl 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Show all files"
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
      Left            =   7770
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   240
      Width           =   1000
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Find"
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
      Left            =   5745
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   255
      Width           =   1000
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   3855
      TabIndex        =   9
      Top             =   255
      Width           =   1830
   End
   Begin VB.CommandButton cmdPrintMissing 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
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
      Left            =   7890
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6585
      Width           =   1000
   End
   Begin VB.CommandButton cmdPrintContents 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
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
      Left            =   7845
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   1000
   End
   Begin VB.ComboBox cboFilename 
      Appearance      =   0  'Flat
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
      Height          =   360
      ItemData        =   "frmBrowseScanFiles.frx":0000
      Left            =   1185
      List            =   "frmBrowseScanFiles.frx":0002
      TabIndex        =   2
      Top             =   1200
      Width           =   6600
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Go"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1155
      Width           =   1000
   End
   Begin MSComctlLib.ListView lvwReview 
      Height          =   2520
      Left            =   135
      TabIndex        =   0
      Top             =   1830
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   4445
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Delivered price"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMissing 
      Height          =   2490
      Left            =   150
      TabIndex        =   4
      Top             =   4710
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   4392
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Preceding"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Trailing"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Qty"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Find scan files containing . . ."
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
      Height          =   315
      Left            =   1035
      TabIndex        =   10
      Top             =   285
      Width           =   2685
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Scanned items not on database (All files)"
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
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   4425
      Width           =   6210
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contents of selected file"
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
      Height          =   315
      Left            =   165
      TabIndex        =   5
      Top             =   1560
      Width           =   2430
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "File name"
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
      Height          =   225
      Left            =   150
      TabIndex        =   3
      Top             =   1230
      Width           =   885
   End
End
Attribute VB_Name = "frmBrowseScanFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents oRep As z_reports
Attribute oRep.VB_VarHelpID = -1
Dim strSQL As String
Dim strEAN As String
Dim strPID As String
Dim oBatch As z_SQL
Dim strFilename As String
Dim strMsg As String




Private Sub cboFilename_Click()
    cboFilename.ToolTipText = cboFilename.Text
End Sub

Private Sub cmdFind_Click()
Dim rs As ADODB.Recordset

    If Trim(txtCode) = "" Then
        LoadListOfImports
    Else
        Set oBatch = New z_SQL
        Set rs = New ADODB.Recordset
    
        oBatch.RunGetRecordset "sp_FindScannedItem", enStoredProcedure, Array(Trim(txtCode)), "", rs
        cboFilename.Clear
        If rs.State > 0 Then   'the recordset is open i.e. rows have been found
            Do While Not rs.eof And Not rs.BOF
                cboFilename.AddItem rs.Fields(0)
                strPID = FNS(rs.Fields(1))
                rs.MoveNext
            Loop
        End If
        If rs.State <> 0 Then rs.Close
        Set rs = Nothing
        If cboFilename.ListCount > 0 Then
            cboFilename.ListIndex = 0
        End If
    End If
    
End Sub

Private Sub cmdGo_Click()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim itmList As ListItem
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT * from vScanSTImportFile WHERE FN='" & cboFilename & "'"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, oPC.COSHORT
    lvwReview.ListItems.Clear
    Do While Not rs.eof
        Set itmList = lvwReview.ListItems.Add
        itmList.Key = FNS(rs.Fields("ID")) & "k"
        itmList.Text = FNS(rs.Fields("CODE"))
        If itmList.Text = Trim(txtCode) Or FNS(rs.Fields("PID")) = strPID Then
            itmList.ForeColor = vbRed
        End If
        itmList.SubItems(1) = FNS(rs.Fields("P_Title"))
        itmList.SubItems(2) = FNS(rs.Fields("QTY"))
        itmList.SubItems(3) = Format(FNN(rs.Fields("P_SP")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
        itmList.SubItems(4) = Format(FNN(rs.Fields("P_LastPriceDelivered")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdGo_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdPrintContents_Click()
Dim ar As New arScannedfile
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT * from vScanSTImportFile WHERE FN='" & cboFilename & "'"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, oPC.COSHORT
    
    ar.Component rs, "Contents of file: " & cboFilename, Trim(strPID)
    ar.left = 400
    ar.top = 1000
    ar.Width = 12000
    ar.Height = 6000
    ar.Show vbModal
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

Private Sub cmdPrintMissing_Click()
Dim ar As New arMissing_1
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT * from STOCKTAKE_WORKM ORDER BY MISSINGID"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, oPC.COSHORT

    ar.Component rs, "Items missing from files scanned"
    ar.Printer.Orientation = ddOLandscape
  '  ar.StartUpPosition = vbStartUpScreen
    ar.left = 400
    ar.top = 1000
    ar.Width = 12000
    ar.Height = 6000
    ar.Show vbModal
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

Private Sub cmdShowALl_Click()
txtCode = ""
cmdFind_Click
End Sub

Private Sub Form_Load()
    Set oRep = New z_reports
    LoadListOfImports
    LoadMissing
End Sub

Private Sub LoadListOfImports()
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cboFilename.Clear
    Set rs = oRep.Filenames
    If rs.State > 0 Then   'the recordset is open i.e. rows have been found
        Do While Not rs.eof And Not rs.BOF
            cboFilename.AddItem rs.Fields(1)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    cboFilename.ToolTipText = cboFilename.Text
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

Private Sub LoadMissing()
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim itmList As ListItem
Dim fs As FileSystemObject
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set fs = New FileSystemObject
    
    strSQL = "SELECT * from STOCKTAKE_WORKM ORDER BY MISSINGID"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, oPC.COSHORT
    lvwMissing.ListItems.Clear
    Do While Not rs.eof
        Set itmList = lvwMissing.ListItems.Add
        itmList.Key = FNS(rs.Fields("MISSINGID")) & "k"
        itmList.Text = FNS(rs.Fields("CODE"))
        itmList.SubItems(1) = fs.GetFileName((rs.Fields("FILENAME")))
        itmList.SubItems(2) = FNS(rs.Fields("PRECEDING"))
        itmList.SubItems(3) = FNS(rs.Fields("TRAILING"))
        itmList.SubItems(4) = FNS(rs.Fields("QTY"))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If Not UnloadMode = 1 Then
'        If MsgBox("Do you want to close the stock-take application?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
'            Cancel = True
'        End If
'    End If
'End Sub
'


Private Sub lvwMissing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If lvwMissing.SortOrder = lvwAscending Then
        lvwMissing.SortOrder = lvwDescending
    Else
        lvwMissing.SortOrder = lvwAscending
    End If
    lvwMissing.SortKey = ColumnHeader.Index - 1
    lvwMissing.Sorted = True
    
End Sub

Private Sub lvwReview_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If lvwReview.SortOrder = lvwAscending Then
        lvwReview.SortOrder = lvwDescending
    Else
        lvwReview.SortOrder = lvwAscending
    End If
    lvwReview.SortKey = ColumnHeader.Index - 1
    lvwReview.Sorted = True
    
End Sub
