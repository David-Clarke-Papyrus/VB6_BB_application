VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   Caption         =   "Papyrus support tasks"
   ClientHeight    =   6735
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   11033
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Properties"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPropertyCount"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dd2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Grid1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdReload"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdOpen"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DD1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Adodc1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Marketing"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdClearMarketing"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Database"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).ControlCount=   2
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   7185
         Top             =   5640
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin TrueOleDBGrid60.TDBDropDown DD1 
         Bindings        =   "frmMain.frx":0054
         Height          =   1305
         Left            =   3630
         OleObjectBlob   =   "frmMain.frx":0069
         TabIndex        =   19
         Top             =   4425
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00D7BDBD&
         Caption         =   "Initialize properties"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5115
         Width           =   2625
      End
      Begin VB.CommandButton cmdOpen 
         BackColor       =   &H00D7BDBD&
         Caption         =   "Connect to database and show tasks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6165
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   390
         Width           =   3960
      End
      Begin VB.Frame Frame3 
         Caption         =   "Purge database prior to reloading"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2430
         Left            =   -68205
         TabIndex        =   10
         Top             =   780
         Width           =   4560
         Begin MSComCtl2.DTPicker dtePriorTo 
            Height          =   450
            Left            =   1380
            TabIndex        =   12
            Top             =   765
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   794
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   20643841
            CurrentDate     =   38827
         End
         Begin VB.CommandButton cmdPurge 
            BackColor       =   &H00D7BDBD&
            Caption         =   "Purge local (secondary) database"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1350
            Width           =   3435
         End
         Begin VB.Label Label1 
            Caption         =   "Clear history prior to"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1320
            TabIndex        =   13
            Top             =   495
            Width           =   1980
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Save to, and load from file"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   885
         Left            =   240
         TabIndex        =   7
         Top             =   5130
         Visible         =   0   'False
         Width           =   2805
         Begin VB.CommandButton cmdLoadFromFile 
            BackColor       =   &H00D7BDBD&
            Caption         =   "Load"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1410
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   270
            Width           =   1080
         End
         Begin VB.CommandButton cmdSaveToFile 
            BackColor       =   &H00D7BDBD&
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   270
            Width           =   1140
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Backup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3900
         Left            =   -74640
         TabIndex        =   4
         Top             =   780
         Width           =   4545
         Begin VB.TextBox txBUName 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1095
            TabIndex        =   14
            Text            =   "PBKSFD"
            Top             =   945
            Width           =   2385
         End
         Begin VB.CommandButton cmdRestore 
            BackColor       =   &H00D7BDBD&
            Caption         =   "Restore local (secondary) database from backup"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   660
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   2700
            Width           =   3180
         End
         Begin VB.CommandButton cmdBackup 
            BackColor       =   &H00D7BDBD&
            Caption         =   "Backup local (secondary) database"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   660
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1845
            Width           =   3180
         End
         Begin VB.Label Label2 
            Caption         =   "Name of backup"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1260
            TabIndex        =   15
            Top             =   615
            Width           =   1980
         End
      End
      Begin VB.CommandButton cmdClearMarketing 
         BackColor       =   &H00D7BDBD&
         Caption         =   "Clear marketing rules"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74700
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   780
         Width           =   2280
      End
      Begin VB.CommandButton cmdReload 
         BackColor       =   &H00C8B9B3&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   9060
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5100
         Width           =   2385
      End
      Begin TrueOleDBGrid60.TDBGrid Grid1 
         Height          =   4305
         Left            =   210
         OleObjectBlob   =   "frmMain.frx":2049
         TabIndex        =   1
         Top             =   735
         Width           =   11205
      End
      Begin TrueOleDBGrid60.TDBDropDown dd2 
         Bindings        =   "frmMain.frx":7E1B
         Height          =   1305
         Left            =   6375
         OleObjectBlob   =   "frmMain.frx":7E30
         TabIndex        =   20
         Top             =   4410
         Width           =   2535
      End
      Begin VB.Label lblPropertyCount 
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   4425
         TabIndex        =   18
         Top             =   480
         Width           =   3810
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPropertyTypes 
         Caption         =   "Set property types"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim rsSave As ADODB.Recordset
Dim XA As XArrayDB
Public strLocalRootFolder As String
Public strLocalServername As String
Dim arCOMMANDLINE() As String
Public oConn  As z_POSConnection
Dim flgLoading As Boolean
Dim cnPapyShort As ADODB.Connection
Public strPassword As String
Dim XP As XArrayDB

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long


Private Sub cmdLoadFromFile_Click()
    On Error GoTo errHandler
    LoadRecordset
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdLoadFromFile_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOpen_Click()
    Connecttodatabase
        LoadPersons

    loadGrid
End Sub


Private Sub cmdReload_Click()
    On Error GoTo errHandler
    XA.Clear
    Set XA = Nothing
    loadGrid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdReload_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdSaveToFile_Click()
    On Error GoTo errHandler
    SaveRecordset
    MsgBox "File saved to " & oConn.SharedFolderRoot & "\BU\Props.xml", vbOKOnly
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdSaveToFile_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Command1_Click()
    Connecttodatabase
    oConn.DBConn.Execute "InitializeProperties"
    loadGrid
End Sub

Private Sub DD1_UnboundFindData(StartLocation As Variant, ByVal ReadPriorRows As Boolean, ByVal IncludeCurrent As Boolean, ByVal Col As Integer, Value As Variant, ByVal SeekFlags As Integer, NewLocation As Variant)
    StartLocation = 0
    ReadPriorRows = True
    IncludeCurrent = True
    Col = 0
    Value = FNN(Grid1.Columns(5))
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim strDataSource As String
    flgLoading = True
    arCOMMANDLINE = Split(Command(), " ")
    Set oConn = New z_POSConnection

    
    If UBound(arCOMMANDLINE) > 0 Then
        oConn.DatabaseName = arCOMMANDLINE(0)
    Else
        oConn.DatabaseName = ""
    End If
    
    InitializeSettings
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Connecttodatabase()
    oConn.dbConnect strLocalServername, strPassword
End Sub

Private Sub InitializeSettings()
    On Error GoTo errHandler
Dim strPCName As String
    
    strPCName = Trim(Me.NameOfPC)
        
    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & strPCName & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
    End If
    

    strLocalServername = GetIniKeyValue(strLocalRootFolder & "\PBKS_Support.INI", "NETWORK", "MAINSQLServer", "")
    strPassword = GetIniKeyValue(strLocalRootFolder & "\PBKS_Support.INI", "NETWORK", "PASSWORD", "")
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.InitializeSettings"
End Sub

Public Property Get NameOfPC() As String
    On Error GoTo errHandler
Dim NameSize As Long
Dim MachineName As String * 16
Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NameOfPC"
End Property
Private Sub loadGrid()
    On Error GoTo errHandler
Dim lngIndex As Long
On Error Resume Next
    rs.Close
    On Error GoTo errHandler
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    'rs.Open "SELECT * FROM dbo.tProperty Order By PropertyKey", oConn.DBConn, adOpenKeyset, adLockOptimistic
    rs.Open "SELECT * FROM dbo.tTask Order By T_SpecifiedDate", oConn.DBConn, adOpenKeyset, adLockOptimistic

    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, rs.RecordCount, 1, 10
    lngIndex = 1
    Do While Not rs.EOF
            XA.Value(lngIndex, 1) = Format(FND(rs.Fields("T_SpecifiedDate")), "dd-mm-YYYY")
            XA.Value(lngIndex, 2) = FNS(rs.Fields("T_Description"))
            XA.Value(lngIndex, 3) = Format(FND(rs.Fields("T_StartDate")), "dd-mm-YYYY")
            XA.Value(lngIndex, 4) = Format(FND(rs.Fields("T_EndDate")), "dd-mm-YYYY")
            XA.Value(lngIndex, 5) = XP(XP.Find(1, 1, FNN(rs.Fields("T_OwnerID"))), 2)
            XA.Value(lngIndex, 6) = XP(XP.Find(1, 1, FNN(rs.Fields("T_SignedOffByID"))), 2)
            XA.Value(lngIndex, 7) = rs.Bookmark
            lngIndex = lngIndex + 1
            rs.MoveNext
    Loop
    Grid1.Array = XA
    Grid1.ReBind
    lblPropertyCount.Caption = CStr(lngIndex - 1) & " properties"


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.loadGrid"
End Sub
Private Sub SaveRecordset()
    On Error GoTo errHandler
Dim fs As New FileSystemObject

    Set rsSave = New ADODB.Recordset
    rsSave.CursorLocation = adUseClient
    rsSave.Open "SELECT * FROM dbo.tProperty Order By PropertyKey", oConn.DBConn, adOpenDynamic, adLockOptimistic, adCmdText
    
    If fs.FileExists(oConn.SharedFolderRoot & "\BU\Props.xml") Then
        fs.DeleteFile oConn.SharedFolderRoot & "\BU\Props.xml", True
    End If
    rsSave.Save oConn.SharedFolderRoot & "\BU\Props.xml", adPersistXML
    rsSave.Close
    Set rsSave = Nothing
    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.SaveRecordset"
End Sub
Private Sub LoadRecordset()
    On Error GoTo errHandler
    rs.Close
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tTask Order BY T_SpecifiedDate", oConn.DBConn, adOpenKeyset, adLockOptimistic

'    Set rsSave = Nothing
'    Set rsSave = New ADODB.Recordset
'    rsSave.Open oConn.SharedFolderRoot & "\BU\Props.xml", "Provider=MSPersist;", adOpenForwardOnly, adLockBatchOptimistic, adCmdFile
'
'    rsSave.MoveFirst
'    Do While Not rsSave.EOF
'        rs.Bookmark = rsSave.Bookmark
'       ' rs.Fields(0) = rsSave.Fields(0)
'        rs.Fields(1) = rsSave.Fields(1)
'        rs.Fields(2) = rsSave.Fields(2)
'     '   rs.Fields(3) = rsSave.Fields(3)
'        rs.Update
'        rsSave.MoveNext
'    Loop
'    rsSave.Close
'    Set rsSave = Nothing
    rs.Close
    Set rs = Nothing
    loadGrid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.LoadRecordset"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    On Error Resume Next
    rs.Close
    Set rs = Nothing

    oConn.dbCloseConnect
    Set oConn = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo errHandler
    If IsNull(Grid1.Bookmark) Then Exit Sub
    rs.Bookmark = XA.Value(Grid1.Bookmark, 7)
    If ColIndex = 4 Then
        rs.Fields("T_OWNERID") = FNN(DD1.Columns(0))
        rs.Update
        Grid1.Text = XP(XP.Find(1, 1, FNN(rs.Fields("T_OwnerID"))), 2)
    ElseIf ColIndex = 5 Then
        rs.Fields("T_SignedOffByID") = FNN(dd2.Columns(0))
        rs.Update
        Grid1.Text = XP(XP.Find(1, 1, FNN(rs.Fields("T_SignedOffByID"))), 2)
    ElseIf ColIndex = 1 Then
        rs.Fields("T_Description") = FNS(Grid1.Text)
        rs.Update
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Grid1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_BeforeInsert(Cancel As Integer)
    XA.ReDim 1, XA.UpperBound(1) + 1, 1, 10
    rs.AddNew
    rs.Fields("T_OWNERID") = 1
End Sub

Private Sub mnuPropertyTypes_Click()
Dim frm As frmPropertyTypes

    Set frm = New frmPropertyTypes
    frm.Show
    
'    Set tlType = Nothing
'    Set tlType = New z_TextList
'    Select Case lngType
'    tlType.Load ltInterestGroupAll
'    lvw.ListItems.Clear
'    For i = 1 To tlType.Count
'        Set lstItem = lvw.ListItems.Add
'        With lstItem
'            .Text = tlType.ItemByOrdinalIndex(i)
'            .SubItems(1) = tlType.f3byOrdinalIndex(i)
'        End With
'    Next

End Sub
Private Sub Grid1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant
    If flgLoading Then Exit Sub

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, XTYPE_STRING
    
    Grid1.Refresh
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.Grid_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Sub LoadPersons()
On Error GoTo errHandler
Dim lngIndex As Long
Dim ArrayIdx As Long
Dim vntItem As Variant
Dim rs As New ADODB.Recordset
Dim i As Integer

    Set rs = Nothing
    rs.Open "SELECT * FROM tPerson", oConn.DBConn, adOpenStatic
    Set XP = New XArrayDB
    rs.MoveFirst
    i = 1
    Do While Not rs.EOF
            XP.ReDim 1, i, 1, 3
            XP.Value(i, 1) = FNN(rs.Fields(0))
            XP.Value(i, 2) = FNS(rs.Fields(1))
            i = i + 1
            rs.MoveNext
    Loop
    
    DD1.Array = XP
    DD1.ReBind
    DD1.Refresh
    dd2.Array = XP
    dd2.ReBind
    dd2.Refresh
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.LoadPersons"
End Sub

