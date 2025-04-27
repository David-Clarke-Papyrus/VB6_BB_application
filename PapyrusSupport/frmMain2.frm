VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Papyrus support tasks"
   ClientHeight    =   6735
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15540
   Icon            =   "frmMain2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   15540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddTask 
      Caption         =   "Command2"
      Height          =   465
      Left            =   14850
      TabIndex        =   19
      Top             =   5790
      Width           =   510
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   11033
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tasks"
      TabPicture(0)   =   "frmMain2.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPropertyCount"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblElapsedTime"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Adodc4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Adodc3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "G2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dd2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdReload"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdOpen"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DD1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Adodc1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboCategory"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Timer1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Tab2"
      TabPicture(1)   =   "frmMain2.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "txtTicketRequest"
      Tab(1).Control(2)=   "cmdLoadTicketRequest"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "tab3"
      TabPicture(2)   =   "frmMain2.frx":0E7A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdLoadTicketResponse"
      Tab(2).Control(1)=   "txtTicketResponse"
      Tab(2).ControlCount=   2
      Begin VB.Timer Timer1 
         Interval        =   60000
         Left            =   7740
         Top             =   510
      End
      Begin VB.TextBox txtTicketResponse 
         Height          =   5640
         Left            =   -74790
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "frmMain2.frx":0E96
         Top             =   435
         Width           =   8625
      End
      Begin VB.CommandButton cmdLoadTicketResponse 
         Caption         =   "Command2"
         Height          =   660
         Left            =   -66090
         TabIndex        =   15
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CommandButton cmdLoadTicketRequest 
         Caption         =   "Command2"
         Height          =   660
         Left            =   -66075
         TabIndex        =   14
         Top             =   5400
         Width           =   1455
      End
      Begin VB.TextBox txtTicketRequest 
         Height          =   5430
         Left            =   -74775
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "frmMain2.frx":0E9C
         Top             =   645
         Width           =   8625
      End
      Begin VB.ComboBox cboCategory 
         Height          =   315
         ItemData        =   "frmMain2.frx":0EA2
         Left            =   225
         List            =   "frmMain2.frx":0EB8
         TabIndex        =   11
         Text            =   "All"
         Top             =   600
         Width           =   2370
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   7005
         Top             =   5775
         Visible         =   0   'False
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
         Bindings        =   "frmMain2.frx":0EEE
         Height          =   1305
         Left            =   3645
         OleObjectBlob   =   "frmMain2.frx":0F03
         TabIndex        =   8
         Top             =   4425
         Visible         =   0   'False
         Width           =   1935
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
         TabIndex        =   6
         Top             =   5115
         Visible         =   0   'False
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
         Left            =   2655
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   570
         Width           =   3960
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
         TabIndex        =   2
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
            TabIndex        =   4
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
            TabIndex        =   3
            Top             =   270
            Width           =   1140
         End
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
         TabIndex        =   1
         Top             =   5100
         Visible         =   0   'False
         Width           =   2385
      End
      Begin TrueOleDBGrid60.TDBDropDown dd2 
         Bindings        =   "frmMain2.frx":2A2F
         Height          =   1305
         Left            =   6375
         OleObjectBlob   =   "frmMain2.frx":2A44
         TabIndex        =   9
         Top             =   4410
         Visible         =   0   'False
         Width           =   2535
      End
      Begin TrueOleDBGrid60.TDBGrid G2 
         Bindings        =   "frmMain2.frx":4570
         Height          =   5055
         Left            =   195
         OleObjectBlob   =   "frmMain2.frx":4585
         TabIndex        =   10
         Top             =   975
         Width           =   14025
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   4455
         Top             =   5805
         Visible         =   0   'False
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
         Caption         =   "Adodc3"
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
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   330
         Left            =   5745
         Top             =   5790
         Visible         =   0   'False
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
         Caption         =   "Adodc4"
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
      Begin VB.Label lblElapsedTime 
         Caption         =   "Label2"
         Height          =   315
         Left            =   8760
         TabIndex        =   18
         Top             =   570
         Width           =   2265
      End
      Begin VB.Label Label1 
         Caption         =   "Paste the email from the ticketing system here"
         Height          =   240
         Left            =   -74775
         TabIndex        =   17
         Top             =   420
         Width           =   3690
      End
      Begin VB.Label Label3 
         Caption         =   "Category"
         Height          =   240
         Left            =   675
         TabIndex        =   12
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label lblPropertyCount 
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   4425
         TabIndex        =   7
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
Dim rs2 As ADODB.Recordset
Dim rsLoggedByPerson As ADODB.Recordset
Dim rsPersonsSignedOff As ADODB.Recordset
Dim XA As XArrayDB
Dim arCOMMANDLINE() As String
Public oConn  As z_POSConnection
Dim flgLoading As Boolean
Dim cnPapyShort As ADODB.Connection
Dim XP As XArrayDB
Dim frmSTask As frmSTask
Dim lngCurrentTaskID As Long
Dim lngCurrentActivityID As Long
Dim lngAccumulatedTime As Long
Dim lngCurrentACtiveTaskGridRow As Long
Dim i As Integer


Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Private Sub cmdAddTask_Click()
    XA.ReDim 1, XA.UpperBound(1) + 1, 1, 20
    loadGrid


















End Sub























Private Sub G2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    i = ColIndex + 1
    On Error Resume Next
            XA(G2.Bookmark, ColIndex + 1) = G2.Text
    UpdateCurrentRow
End Sub
Private Sub UpdateCurrentRow()
    oConn.DBConn.Execute "UPDATE tTask SET " _
        & " T_SPECIFIEDDATE = '" & ReverseDate(XA(G2.Bookmark, 1)) & "'," _
        & " T_OWNERID = '" & XA(G2.Bookmark, 2) & "'," _
        & " T_DESCRIPTION = '" & XA(G2.Bookmark, 3) & "'," _
        & " T_Note = '" & XA(G2.Bookmark, 11) & "'," _
        & " T_ACTIVE = '" & XA(G2.Bookmark, 4) & "'," _
        & " T_CATEGORY = '" & XA(G2.Bookmark, 5) & "'," _
        & " T_DONEBYID = '" & XA(G2.Bookmark, 6) & "'," _
        & " T_STARTDATE = '" & ReverseDate(XA(G2.Bookmark, 7)) & "'," _
        & " T_DUEBYDATE = '" & ReverseDate(XA(G2.Bookmark, 8)) & "'," _
        & " T_SIGNEDOFFDATE = '" & ReverseDate(XA(G2.Bookmark, 9)) & "'," _
        & " T_SIGNEDOFFBYID = '" & XA(G2.Bookmark, 10) & "'" _
        & " WHERE T_ID = " & CStr(XA(G2.Bookmark, 20))
End Sub

Private Sub G2_ButtonClick(ByVal ColIndex As Integer)
Dim tmp
    If ColIndex = 3 Then
        Me.lblElapsedTime = ""
        G2.Bookmark = G2.Bookmark
        If G2.Text <> "ACTIVE" Then
            EndActivityRecord True
            G2.Text = "ACTIVE"
            lngCurrentACtiveTaskGridRow = G2.Bookmark
            InsertNewActivityRecord
            ShowElapsedTime
            Me.Timer1.Enabled = True
        Else
           G2.Text = ""
            EndActivityRecord False
            lngCurrentTaskID = 0
        End If
    End If
End Sub
Private Sub G2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  '  If IsNull(LastRow) Or LastCol = -1 Then Exit Sub
  '  G2.Col(LastRow, LastCol) = ""
End Sub


Private Sub cmdLoadTicketRequest_Click()
Dim strDate As String
Dim strTicketMessage As String
Dim strMember As String
Dim strSUbject As String
Dim sTicketID As String
Dim lngTicketId As Long

Dim sLine As String
Dim s() As String
Dim i As Integer

    s = Split(txtTicketRequest, Chr(13))
    
    For i = 0 To UBound(s) - 1
        sLine = s(i)
        If InStr(2, sLine, "Ticket ID:") > 0 Then
            sTicketID = Right(sLine, Len(sLine) - 12)
            If IsNumeric(sTicketID) Then
                lngTicketId = CLng(sTicketID)
            End If
        End If
    Next i
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



Private Sub Command1_Click()
    Connecttodatabase
    oConn.DBConn.Execute "InitializeProperties"
    loadGrid
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
    
MsgBox "local root name = " & strLocalRootFolder & "\PBKS_Support.INI"
    strLocalServername = GetIniKeyValue(strLocalRootFolder & "\PBKS_Support.INI", "NETWORK", "MAINSQLServer", "")
    MsgBox "local server name = " & strLocalServername
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
Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
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
Dim strSQL As String
Dim i As Integer

    rs.Close
    On Error GoTo errHandler
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    If cboCategory = "All" Then
        strSQL = "SELECT * FROM dbo.tTask WHERE T_ParentTaskID IS NULL Order By T_SpecifiedDate DESC "
        rs.Open strSQL, oConn.DBConn, adOpenKeyset, adLockOptimistic
    Else
        strSQL = "SELECT * FROM dbo.tTask WHERE T_CATEGORY = '" & Me.cboCategory & "' AND  T_ParentTaskID IS NULL Order By T_SpecifiedDate DESC "
        rs.Open strSQL, oConn.DBConn, adOpenKeyset, adLockOptimistic
    End If
    Set XA = Nothing
    Set XA = New XArrayDB
    For i = 1 To rs.RecordCount
        XA.ReDim 1, i, 1, 20
        XA(i, 1) = Format(FND(rs.Fields("T_SPECIFIEDDATE")), "dd/mm/yyyy")
        XA(i, 2) = FNS(rs.Fields("T_OWNERID"))
        XA(i, 3) = FNS(rs.Fields("T_DESCRIPTION"))
        XA(i, 4) = FNS(rs.Fields("T_ACTIVE"))
        XA(i, 5) = FNS(rs.Fields("T_CATEGORY"))
        XA(i, 6) = FNS(rs.Fields("T_DONEBYID"))
        XA(i, 7) = Format(FND(rs.Fields("T_STARTDATE")), "dd/mm/yyyy")
        XA(i, 8) = Format(FND(rs.Fields("T_DUEBYDATE")), "dd/mm/yyyy")
        XA(i, 9) = Format(FND(rs.Fields("T_SIGNEDOFFDATE")), "dd/mm/yyyy")
        XA(i, 10) = FNS(rs.Fields("T_SIGNEDOFFBYID"))
        XA(i, 11) = FNS(rs.Fields("T_NOTE"))
        XA(i, 20) = CStr(FNN(rs.Fields("T_ID")))
        rs.MoveNext
    Next
        Set G2.Array = XA
        G2.ReBind
        G2.Refresh
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.loadGrid"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    
    If lngCurrentActivityID > 0 Then
        If MsgBox("Closing application will end the activity on the currently active task. Do you want to continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
            Cancel = True
            Exit Sub
        Else
            EndActivityRecord True
        End If
    End If
    On Error Resume Next
    
    rs.Close
    Set rs = Nothing
    oConn.DBConn.Execute "UPDATE tTASK SET T_ACTIVE = ''"
    EndActivityRecord False
    oConn.dbCloseConnect
    Set oConn = Nothing
'    DeleteObject BackBrush
'    DeleteObject EllBrush
'    DeleteObject NewFont
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub




Private Sub G2_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
Dim lngSupplierID As Long
Dim lngDEALID As Long

'    RowStyle.BackColor = G2.EvenRowStyle.BackColor
'    If G2.Text = "ACTIVE" Then
'        RowStyle.BackColor = RGB(232, 174, 180)
'    End If
'    DoEvents

End Sub

Private Sub G2_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant
    If flgLoading Then Exit Sub

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
        Select Case ColIndex
        Case 0
            rs.Sort = "T_SPECIFIEDDATE ASC"
        Case 1
            rs.Sort = "T_OWNERID ASC"
        Case 2
            rs.Sort = "T_DESCRIPTION ASC"
        Case 3
            rs.Sort = "T_CATEGORY ASC"
        Case 4
            rs.Sort = "T_DONEBYID ASC"
        Case 5
            rs.Sort = "T_STARTDATE ASC"
        Case 6
            rs.Sort = "T_DUEBYDATE ASC"
        Case 7
            rs.Sort = "T_SIGNEDOFFDATE ASC"
        Case 8
            rs.Sort = "T_SIGNEDOFFBYID ASC"
        End Select
    Else
        Direction = 0
        Select Case ColIndex
        Case 0
            rs.Sort = "T_SPECIFIEDDATE DESC"
        Case 1
            rs.Sort = "T_OWNERID DESC"
        Case 2
            rs.Sort = "T_DESCRIPTION DESC"
        Case 3
            rs.Sort = "T_CATEGORY DESC"
        Case 4
            rs.Sort = "T_DONEBYID DESC"
        Case 5
            rs.Sort = "T_STARTDATE DESC"
        Case 6
            rs.Sort = "T_DUEBYDATE DESC"
        Case 7
            rs.Sort = "T_SIGNEDOFFDATE DESC"
        Case 8
            rs.Sort = "T_SIGNEDOFFBYID DESC"
        End Select
    End If
    G2.Refresh
    Screen.MousePointer = vbDefault

'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseCOs.Grid_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.G2_HeadClick(ColIndex)", ColIndex
End Sub

Private Sub LoadPersons()
On Error GoTo errHandler
Dim lngIndex As Long
Dim ArrayIdx As Long
Dim vntItem As Variant
Dim i As Integer

    Set rsLoggedByPerson = Nothing
    Set rsLoggedByPerson = New ADODB.Recordset
    rsLoggedByPerson.CursorLocation = adUseClient
    rsLoggedByPerson.Open "SELECT * FROM tPerson", oConn.DBConn, adOpenStatic
    Set Me.Adodc3.Recordset = rsLoggedByPerson
    DD1.ReBind
    DD1.Refresh
    
    Set rsPersonsSignedOff = Nothing
    Set rsPersonsSignedOff = New ADODB.Recordset
    rsPersonsSignedOff.CursorLocation = adUseClient
    rsPersonsSignedOff.Open "SELECT * FROM tPerson", oConn.DBConn, adOpenStatic
    Set Me.Adodc4.Recordset = rsPersonsSignedOff
    dd2.ReBind
    dd2.Refresh
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.LoadPersons"
End Sub

Private Sub G2_DblClick()
    Set frmSTask = New frmSTask
    frmSTask.Component XA(G2.Bookmark, 11), XA(G2.Bookmark, 3), G2.Bookmark, XA
    frmSTask.Show vbModal
    XA(G2.Bookmark, 3) = frmSTask.Description
    XA(G2.Bookmark, 11) = frmSTask.Note
    UpdateCurrentRow
    Unload frmSTask
End Sub


Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub InsertNewActivityRecord()
    
    
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "InsertActivityRecord"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    Set par = cmd.CreateParameter("@OldActivityID", adInteger, adParamInput, , lngCurrentActivityID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TaskID", adInteger, adParamInput, , CInt(XA(G2.Bookmark, 20)))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@PersonID", adVarChar, adParamInput, 20, oConn.OperatorName)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@NewActivityID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oConn.DBConn
    cmd.Execute
    lngCurrentTaskID = CInt(XA(G2.Bookmark, 20))
    lngCurrentActivityID = FNN(cmd.Parameters("@NewActivityID"))
    Set cmd = Nothing
    
    
End Sub
Private Sub EndActivityRecord(pUpdateOnly As Boolean)
    
    
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "EndActivityRecord"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    Set par = cmd.CreateParameter("@OldActivityID", adInteger, adParamInput, , lngCurrentActivityID)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oConn.DBConn
    cmd.Execute
    
    If pUpdateOnly = False Then
        lngCurrentActivityID = 0
        Me.Timer1.Enabled = False
    End If
    Set cmd = Nothing
    
    
End Sub

Private Sub GetAccumulatedTime()
    
    
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetAccumulatedTime"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    Set par = cmd.CreateParameter("@TaskID", adInteger, adParamInput, , CInt(XA(G2.Bookmark, 20)))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@CurrentActivityID", adInteger, adParamInput, , lngCurrentActivityID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@PersonID", adVarChar, adParamInput, 20, oConn.OperatorName)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pAccumulatedMinutes", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oConn.DBConn
    cmd.Execute
    
    lngAccumulatedTime = FNN(cmd.Parameters("@pAccumulatedMinutes"))
    
    Set cmd = Nothing
    
    
End Sub

Private Sub Timer1_Timer()
    ShowElapsedTime
End Sub


Private Sub ShowElapsedTime()
    GetAccumulatedTime
    lblElapsedTime.Caption = CStr(lngAccumulatedTime) & " min"
End Sub
