VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Step_3 
   BackColor       =   &H00E8E8DD&
   Caption         =   "Step 3 - Review scanned items"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00D8D9C4&
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
      Height          =   360
      Left            =   5655
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   675
      Width           =   765
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
      Left            =   3765
      TabIndex        =   12
      Top             =   675
      Width           =   1830
   End
   Begin VB.CommandButton cmdPrintMissing 
      BackColor       =   &H00D8D9C4&
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
      Height          =   390
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6795
      Width           =   840
   End
   Begin VB.CommandButton cmdPrintContents 
      BackColor       =   &H00D8D9C4&
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
      Height          =   390
      Left            =   7845
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3945
      Width           =   840
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
      ItemData        =   "frm_Step_3.frx":0000
      Left            =   1185
      List            =   "frm_Step_3.frx":0002
      TabIndex        =   5
      Top             =   1200
      Width           =   6600
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00D8D9C4&
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
      Height          =   420
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1155
      Width           =   840
   End
   Begin VB.CommandButton cmdPrev_to_2 
      BackColor       =   &H00D8D9C4&
      Caption         =   "&Prev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7275
      Width           =   840
   End
   Begin VB.CommandButton cmdNext_To_4 
      BackColor       =   &H00D8D9C4&
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6945
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7275
      Width           =   840
   End
   Begin MSComctlLib.ListView lvwReview 
      Height          =   2520
      Left            =   135
      TabIndex        =   3
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
      TabIndex        =   7
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
      Left            =   945
      TabIndex        =   13
      Top             =   705
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   6
      Top             =   1230
      Width           =   885
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "You are starting a new stock-take.   Click the button to continue"
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
      Height          =   645
      Left            =   570
      TabIndex        =   0
      Top             =   75
      Width           =   6615
   End
End
Attribute VB_Name = "frm_Step_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents oSA As a_Stktke
Attribute oSA.VB_VarHelpID = -1
Dim strSql As String
Dim strEAN As String
Dim strPID As String
Dim oBatch As z_SQL
Dim strFilename As String
Dim strMsg As String

Public Sub Component(pSA As a_Stktke)
    Set oSA = pSA
End Sub


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
            Do While Not rs.EOF And Not rs.BOF
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
    strSql = "SELECT * from vScanSTImportFile WHERE FN='" & cboFilename & "'"
    Set rs = New ADODB.Recordset
    rs.Open strSql, oPC.COshort
    lvwReview.ListItems.Clear
    Do While Not rs.EOF
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

Private Sub cmdNext_To_4_Click()
    Set frm4 = New frm_Step_4
    frm4.Component oSA
    frm4.Show
    Unload Me
End Sub

Private Sub cmdPrev_to_2_Click()
    Set frm2 = New frm_Step_2
    frm2.Component oSA
    frm2.Show
    Unload Me
End Sub

Private Sub cmdPrintContents_Click()
Dim ar As New arScannedfile
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSql = "SELECT * from vScanSTImportFile WHERE FN='" & cboFilename & "'"
    Set rs = New ADODB.Recordset
    rs.Open strSql, oPC.COshort
    
    ar.Component rs, "Contents of file: " & cboFilename, Trim(strPID)
    ar.Left = 400
    ar.Top = 1000
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
    strSql = "SELECT * from STOCKTAKE_WORKM ORDER BY MISSINGID"
    Set rs = New ADODB.Recordset
    rs.Open strSql, oPC.COshort

    ar.Component rs, "Items missing from files scanned"
    ar.Printer.Orientation = ddOLandscape
  '  ar.StartUpPosition = vbStartUpScreen
    ar.Left = 400
    ar.Top = 1000
    ar.Width = 12000
    ar.Height = 6000
    ar.Show vbModal
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

Private Sub Form_Load()
    strMsg = "The files that have been imported are listed below." & vbCrLf _
    & "You can review their contents by double-clicking the file name."
    lblMsg.Caption = strMsg
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
    Set rs = oSA.Filenames
    If rs.State > 0 Then   'the recordset is open i.e. rows have been found
        Do While Not rs.EOF And Not rs.BOF
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
Dim strSql As String
Dim rs As ADODB.Recordset
Dim itmList As ListItem
Dim fs As FileSystemObject
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set fs = New FileSystemObject
    
    strSql = "SELECT * from STOCKTAKE_WORKM ORDER BY MISSINGID"
    Set rs = New ADODB.Recordset
    rs.Open strSql, oPC.COshort
    lvwMissing.ListItems.Clear
    Do While Not rs.EOF
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not UnloadMode = 1 Then
        If MsgBox("Do you want to close the stock-take application?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
            Cancel = True
        End If
    End If
End Sub



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
