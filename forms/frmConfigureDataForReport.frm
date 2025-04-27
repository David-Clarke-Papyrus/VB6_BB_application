VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfigureDataForReport 
   BackColor       =   &H00E6E6D9&
   Caption         =   "Configure data for report"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5535
   ScaleWidth      =   10095
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Close"
      Height          =   465
      Left            =   8805
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4470
      Left            =   210
      TabIndex        =   2
      Top             =   720
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   7885
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   15132377
      TabCaption(0)   =   "Sorting"
      TabPicture(0)   =   "frmConfigureDataForReport.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lbColumnsToSort"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdRemoveSortColumn"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdAddSortColumnDesc"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAddSortColumnAsc"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbColumnsAvailableToSort"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Filters"
      TabPicture(1)   =   "frmConfigureDataForReport.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lbColumnsAvailableToFilter"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtMarkers"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtTokens"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtSQL"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdTestFilter"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lvw1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cboTokens"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdToken"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.CommandButton cmdToken 
         Height          =   330
         Left            =   9105
         Picture         =   "frmConfigureDataForReport.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2895
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.ComboBox cboTokens 
         Height          =   315
         ItemData        =   "frmConfigureDataForReport.frx":03C2
         Left            =   6615
         List            =   "frmConfigureDataForReport.frx":03CF
         TabIndex        =   19
         Text            =   "last_quarter"
         Top             =   2910
         Visible         =   0   'False
         Width           =   2460
      End
      Begin MSComctlLib.ListView lvw1 
         Height          =   1995
         Left            =   6555
         TabIndex        =   18
         Top             =   690
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   3519
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "marker"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "token"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdTestFilter 
         BackColor       =   &H00D5D5C1&
         Caption         =   "Test filter (and save)"
         Height          =   285
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4080
         Width           =   1710
      End
      Begin VB.TextBox txtSQL 
         Height          =   3300
         Left            =   3330
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   630
         Width           =   5790
      End
      Begin VB.TextBox txtTokens 
         Height          =   795
         Left            =   3270
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1875
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.TextBox txtMarkers 
         Height          =   795
         Left            =   3345
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.ListBox lbColumnsAvailableToFilter 
         Height          =   3570
         Left            =   180
         TabIndex        =   10
         Top             =   495
         Width           =   2850
      End
      Begin VB.ListBox lbColumnsToSort 
         Height          =   1035
         Left            =   -70995
         TabIndex        =   7
         Top             =   495
         Width           =   2850
      End
      Begin VB.CommandButton cmdRemoveSortColumn 
         Height          =   495
         Left            =   -71880
         Picture         =   "frmConfigureDataForReport.frx":03F8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2670
         Width           =   705
      End
      Begin VB.CommandButton cmdAddSortColumnDesc 
         Caption         =   "DESC"
         Height          =   495
         Left            =   -71865
         Picture         =   "frmConfigureDataForReport.frx":0782
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1800
         Width           =   705
      End
      Begin VB.CommandButton cmdAddSortColumnAsc 
         Caption         =   "ASC"
         Height          =   495
         Left            =   -71865
         Picture         =   "frmConfigureDataForReport.frx":0B0C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1290
         Width           =   705
      End
      Begin VB.ListBox lbColumnsAvailableToSort 
         Height          =   3570
         Left            =   -74820
         TabIndex        =   3
         Top             =   495
         Width           =   2850
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter text in SQL"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   3360
         TabIndex        =   16
         Top             =   390
         Width           =   1740
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter text with tokens"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   3375
         TabIndex        =   14
         Top             =   1665
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter text with markers"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   4980
         TabIndex        =   12
         Top             =   375
         Visible         =   0   'False
         Width           =   1740
      End
   End
   Begin VB.ComboBox cboReport_View 
      Height          =   315
      Left            =   195
      TabIndex        =   1
      Text            =   "cboDatabase_View"
      Top             =   315
      Width           =   2535
   End
   Begin VB.CommandButton cmdLoadData 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Load data"
      Height          =   345
      Left            =   2790
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   270
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Database view"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   195
      TabIndex        =   8
      Top             =   90
      Width           =   1635
   End
End
Attribute VB_Name = "frmConfigureDataForReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oS As z_SearchUDR
Dim tl As New z_TextListSimple
Dim tlCol As New z_TextListSimple
Dim strXML As String
Dim rs As ADODB.Recordset
Dim rpt_Metadata As String
Dim rpt_Name As String
Dim rpt_View As String
Dim strSQL As String
Dim sortString As String
Dim filterstring As String
Dim WithEvents oMD As z_ReportMetadata
Attribute oMD.VB_VarHelpID = -1
Dim flgLoading As Boolean
Dim o As Collection

Event ReloadReportList()

Public Sub Component(pMD As z_ReportMetadata)
    Set oMD = pMD
End Sub

Private Sub cboReport_view_Click()
    If flgLoading Then Exit Sub
    oMD.Report_view = Me.cboReport_View
    oMD.clearSortingFilteringSelection
    LoadViewColumns

End Sub

Private Sub cmdAddSortColumnAsc_Click()
Dim i As Integer
    If lbColumnsAvailableToSort.SelCount = 0 Then Exit Sub
    For i = 0 To lbColumnsAvailableToSort.ListCount - 1
        If Me.lbColumnsAvailableToSort.Selected(i) = True Then
            oMD.AddSortedVolumn lbColumnsAvailableToSort.Text, "ASC"
            Me.lbColumnsToSort.AddItem lbColumnsAvailableToSort.Text & "(ASC)"
        End If
    Next
End Sub

Private Sub cmdAddSortColumnDesc_Click()
Dim i As Integer
    If lbColumnsAvailableToSort.SelCount = 0 Then Exit Sub
    For i = 0 To lbColumnsAvailableToSort.ListCount - 1
        If Me.lbColumnsAvailableToSort.Selected(i) = True Then
            oMD.AddSortedVolumn lbColumnsAvailableToSort.Text, "DESC"
            Me.lbColumnsToSort.AddItem lbColumnsAvailableToSort.Text & "(DESC)"
        End If
    Next

End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdLoadData_Click()
    LoadViewColumns
End Sub

Private Sub cmdRemoveSortColumn_Click()
Dim i As Integer
    If Me.lbColumnsToSort.SelCount = 0 Then Exit Sub
    oMD.RemoveSortedVolumn Left(lbColumnsToSort.Text, InStr(1, lbColumnsToSort.Text, "(") - 1)
    Me.lbColumnsToSort.RemoveItem lbColumnsToSort.ListIndex

End Sub

Private Sub cmdTestFilter_Click()
Dim oSQL As New z_SQL
Dim SQLMsg As String

    If oSQL.TestSQL(oMD.GetSQLShort & " " & Me.txtSQL & " " & oMD.GetSQLOrderBy, SQLMsg) = False Then
        MsgBox "False SQL" & SQLMsg
    Else
        MsgBox "SQL OK"
        oMD.AppendFilter Me.txtSQL
    End If
    
End Sub

Private Sub cmdToken_Click()
    lvw1.SelectedItem.SubItems(1) = cboTokens
End Sub

Private Sub Form_Initialize()
Set oS = New z_SearchUDR
End Sub

Private Sub Form_Load()
    flgLoading = True
    tl.Load sltAdhocQueries
    LoadComboFromTextListSimple Me.cboReport_View, tl
    LoadControls
    flgLoading = False
    SSTab1.Tab = 0
End Sub
Private Sub LoadControls()
Dim s As String
Dim i As Integer

    'Load the view and the available column list
'    Me.txtReportName = rpt_Name
    cboReport_View = oMD.Report_view
    
    LoadViewColumns
    
    'Load the sorted by list
    oMD.ParseMetadata
    sortString = oMD.sortString
    filterstring = oMD.filterstring
    Me.txtSQL = filterstring
End Sub

Private Sub LoadViewColumns()
    tl.Load sltFieldList, Me.cboReport_View
    LoadListboxSimple Me.lbColumnsAvailableToSort, tl
    LoadListboxSimple Me.lbColumnsAvailableToFilter, tl
End Sub

Public Property Get Metadata_XML() As String
    Metadata_XML = oMD.Metadata_XML
End Property

Private Sub oMD_LoadSortedItem(Item As String)
    lbColumnsToSort.AddItem Item
End Sub

Private Sub txtMarkers_Change()
    LoadExtractedMarkers ExtractMarkers(txtMarkers)
End Sub
Private Function ExtractMarkers(sIn As String) As String()
Dim EOM As Boolean
Dim s() As String
Dim markerCount As Integer
Dim startPos As Integer
Dim markerStartPos As Integer
Dim markerEndPos As Integer

    ReDim s(0)
    markerCount = 0
    startPos = 1
    EOM = False
    Do While Not EOM
        markerStartPos = InStr(startPos, sIn, "{")
        If markerStartPos > 0 Then
            markerCount = markerCount + 1
            markerEndPos = InStr(startPos + 1, sIn, "}")
            If markerEndPos <= markerStartPos Then
                Exit Do
            End If
            ReDim Preserve s(UBound(s) + 1)
            s(markerCount) = Mid(sIn, markerStartPos, markerEndPos - markerStartPos + 1)
            startPos = markerEndPos
        Else
            EOM = True
        End If
    Loop
    ExtractMarkers = s
End Function
Private Sub LoadExtractedMarkers(s() As String)
Dim i As Integer

    If UBound(s) = 0 Then
        Exit Sub
    End If
    lvw1.ListItems.Clear
    For i = 1 To UBound(s)
        lvw1.ListItems.Add , , s(i)
    Next
End Sub
