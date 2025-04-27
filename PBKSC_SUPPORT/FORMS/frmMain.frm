VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Papyrus II support"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7995
      Top             =   105
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   195
      TabIndex        =   0
      Top             =   195
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   10610
      _Version        =   393216
      Style           =   1
      TabHeight       =   670
      BackColor       =   14737632
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
      TabCaption(0)   =   "User support      "
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblOne"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblThree"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOne"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtResults"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdThree"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Database tuning and analysis    "
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "G1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chkAutoshrink"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdTables"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdEXport"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdRebuildIndexes"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdTableStats"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdShrink"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdDumpTriggers"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdComp"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdConnect"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Extra     "
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdUpdateFromScript"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdExtract"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.CommandButton cmdUpdateFromScript 
         BackColor       =   &H00D3D3CB&
         Caption         =   "run update scripts (UPDATES.SQL and REPLACEMENTS.SQL) These files should be in the PBKS\Downloads folder."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   -74535
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1920
         Width           =   5145
      End
      Begin VB.CommandButton cmdExtract 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Extract downloaded files"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   -74535
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1110
         Width           =   5145
      End
      Begin VB.CommandButton cmdThree 
         BackColor       =   &H00D3D3CB&
         Caption         =   "B: Send "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   5955
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1515
         Width           =   1635
      End
      Begin VB.TextBox txtResults 
         BackColor       =   &H00E3F9FD&
         ForeColor       =   &H8000000D&
         Height          =   2940
         Left            =   540
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2730
         Width           =   7590
      End
      Begin VB.CommandButton cmdOne 
         BackColor       =   &H00D3D3CB&
         Caption         =   "A: Fetch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   5970
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   615
         Width           =   1635
      End
      Begin VB.CommandButton cmdConnect 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Connect to database (before any other action on this tab)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -72930
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   570
         Width           =   3525
      End
      Begin VB.CommandButton cmdComp 
         BackColor       =   &H00D3D3CB&
         Caption         =   "run compare script"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -71025
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2895
         Width           =   3495
      End
      Begin VB.CommandButton cmdDumpTriggers 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Save Trigger scripts to TRIGGERS.TXT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -74805
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1695
         Width           =   3495
      End
      Begin VB.CommandButton cmdShrink 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Shrink now"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -74805
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3480
         Width           =   3480
      End
      Begin VB.CommandButton cmdTableStats 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Table statistics"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -71025
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1680
         Width           =   3495
      End
      Begin VB.CommandButton cmdRebuildIndexes 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Rebuild indexes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -71025
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2295
         Width           =   3495
      End
      Begin VB.CommandButton cmdEXport 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Export script for tables,views etc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -74805
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2295
         Width           =   3495
      End
      Begin VB.CommandButton cmdTables 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Export script for tables only"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -74820
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2895
         Width           =   3495
      End
      Begin VB.CheckBox chkAutoshrink 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Auto-shrink"
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
         Height          =   405
         Left            =   -70095
         TabIndex        =   1
         Top             =   3525
         Width           =   1635
      End
      Begin TrueOleDBGrid60.TDBGrid G1 
         Height          =   1590
         Left            =   -72345
         OleObjectBlob   =   "frmMain.frx":0054
         TabIndex        =   9
         Top             =   4320
         Width           =   5760
      End
      Begin VB.Label lblThree 
         Caption         =   "Send database script to support"
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
         Height          =   570
         Left            =   1035
         TabIndex        =   15
         Top             =   1665
         Width           =   4695
      End
      Begin VB.Label lblOne 
         Caption         =   "Label1"
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
         Height          =   570
         Left            =   1065
         TabIndex        =   10
         Top             =   735
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iFilenum1 As Integer
Dim iFilenum2 As Integer
Dim strLocalRootFolder As String
Dim strFolderOut As String
Dim strLocalPath As String
Dim strServerMachine As String
Dim strInternetDialup As String
Dim strConnectionName As String
Dim strDownloadFolder As String
Dim strUsername As String
Dim strPWD As String
Dim oDatabase As SQLDMO.Database2
Dim oSQLServer As SQLDMO.SQLServer2
Dim rs As ADODB.Recordset

Dim FTPAddress As String
Dim FTPFolder As String

Const FTPUsername As String = "bt000SA1"
Const FTPPassword As String = "1beach"
Private strClientCode As String
Private Type TableStats
DataSpaceUsed As String
End Type
Dim strServerName As String
Dim X1 As New XArrayDB
Dim strServerMachineName As String
Dim ADOConn As New ADODB.Connection

Private Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Private Declare Function GetClassName Lib "user32" _
    Alias "GetClassNameA" (ByVal hwnd&, _
    ByVal lpClassName$, ByVal nMaxCount&) As Long



Sub Initializesettings()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim strPCName As String
Dim strPos As String

    strPos = "0"
    strPCName = Trim(Me.NameOfPC)
    strLocalRootFolder = "\\" & strPCName & "\PBKS_S"
    strServerName = GetIniKeyValue(strLocalRootFolder & "\CENTRAL.INI", "NETWORK", "MAINSQLSERVER", strPCName)
    strUsername = GetIniKeyValue(strLocalRootFolder & "\CENTRAL.INI", "NETWORK", "USERNAME", "sa")
    strPWD = GetIniKeyValue(strLocalRootFolder & "\CENTRAL.INI", "NETWORK", "PASSWORD", "")
    strServerMachine = GetIniKeyValue(strLocalRootFolder & "\CENTRAL.INI", "NETWORK", "CENTRALSERVERMACHINE", strPCName)
    strServerMachineSharedFolder = "\\" & strServerMachine & "\PBKS_S"
    strPos = "1"
    FTPAddress = GetIniKeyValue(strLocalRootFolder & "\CENTRAL.INI", "SUPPORT", "FTPADDRESS", "")
    FTPFolder = GetIniKeyValue(strLocalRootFolder & "\CENTRAL.INI", "SUPPORT", "FTPFOLDER", "")
    strPos = "2"
    strInternetDialup = GetIniKeyValue(strLocalRootFolder & "\CENTRAL.INI", "SUPPORT", "INTERNETDIALUP", "")
    strConnectionName = GetIniKeyValue(strLocalRootFolder & "\CENTRAL.INI", "SUPPORT", "CONNECTIONNAME", "")
    strFolderOut = strLocalRootFolder & "\FilesForExport"
    strDownloadFolder = strLocalRootFolder & "\DownloadFolder"
    strPos = "3"
    Set ADOConn = New ADODB.Connection
    ADOConn.Provider = "sqloledb"
    ADOConn.ConnectionTimeout = 10
    ADOConn.Open "Data Source=" & strServerName & ";Initial Catalog=PBKSC;User Id=sa;Password=;Network Library=dbmssocn;Connect Timeout=45"
    strPos = "4"
    Set rs = New ADODB.Recordset
    rs.Open "SELECT dbo.tStore.STORE_Code FROM dbo.tConfiguration INNER JOIN dbo.tStore ON dbo.tConfiguration.CF_DefaultStoreID = dbo.tStore.STORE_ID", ADOConn, adOpenStatic
    strPos = "5"
    If Not rs.EOF And Not rs.BOF Then
        strClientCode = FNS(rs.Fields(0))
    End If
    strPos = "6"
    rs.Close
    Set rs = Nothing
    strPos = "7"
    ADOConn.Close
    strPos = "8"
    If Not fs.FolderExists(strFolderOut) Then
        strPos = "9.01"
        fs.CreateFolder strFolderOut
        strPos = "9.1"
    End If
    If Not fs.FolderExists(strDownloadFolder) Then
        strPos = "9.02"
        fs.CreateFolder strDownloadFolder
        strPos = "9.2"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.Initializesettings", , , , "Position", Array(strPos, strFolderOut, strDownloadFolder)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Initializesettings"
End Sub
Private Sub Connect()
    On Error GoTo errHandler

    Set oSQLServer = New SQLDMO.SQLServer
    oSQLServer.LoginTimeout = 0 '-1 is the ODBC default (60) seconds
    With oSQLServer
        .LoginSecure = False
        .AutoReConnect = False
        .Connect strServerName, "sa", ""
    End With
    
    Set oDatabase = oSQLServer.Databases("PBKSC")
    If ADOConn.State <> adStateOpen Then
        ADOConn.Provider = "sqloledb"
        ADOConn.Open "Data Source=" & strServerName & ";Initial Catalog=PBKSC;User Id=sa;Password=; Network Library=dbmssocn;"
    End If
    LoadTriggers
    strServerMachineName = GetIniKeyValue(strLocalPath & "\CENTRAL.INI", "NETWORK", "CENTRALSERVERMACHINE", "")
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.Connect"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Connect"
End Sub
Private Function Disconnect()
    On Error GoTo errHandler
    oSQLServer.Disconnect
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Disconnect"
End Function
Public Property Get Clientcode() As String
    Clientcode = strClientCode
End Property
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

Public Sub ExportScript()
    On Error GoTo errHandler
Dim s As String
Dim Flag As SQLDMO_SCRIPT_TYPE
Dim oTable As SQLDMO.Table
Dim oStoredProc As SQLDMO.StoredProcedure2
Dim oView As SQLDMO.View2
Dim oUser As SQLDMO.User
Dim oUDF As SQLDMO.UserDefinedFunction
Dim oDBRole As SQLDMO.DatabaseRole2

    Screen.MousePointer = vbHourglass
    Set oDatabase = oSQLServer.Databases("PBKSC")
    s = ""
  For Each oStoredProc In oDatabase.StoredProcedures
   ' Debug.Print oStoredProc.Name
    s = s & oStoredProc.Script
  Next
  For Each oView In oDatabase.Views
    s = s & oView.Script
  Next
  For Each oUser In oDatabase.Users
    s = s & oUser.Script
  Next

  Flag = SQLDMOScript_Default Or SQLDMOScript_Indexes Or SQLDMOScript_DRI_AllConstraints Or SQLDMOScript_Triggers Or SQLDMOScript_DRI_ForeignKeys
  For Each oTable In oDatabase.Tables
    If Not oTable.SystemObject Then
      s = s & oTable.Script(Flag)
    End If
  Next
    iFilenum2 = FreeFile
Dim fs As New FileSystemObject
    fs.DeleteFile strFolderOut & "\DBScript.SQL"
    Open strFolderOut & "\DBScript.SQL" For Output As #iFilenum2
    Print #iFilenum2, s
    Close #iFilenum2
    Screen.MousePointer = vbDefault
    MsgBox "Done"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.ExportScript"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ExportScript"
End Sub


Private Sub chkAutoshrink_Click()
    On Error GoTo errHandler
    oDatabase.DBOption.AutoShrink = (chkAutoshrink = 1)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.chkAutoshrink_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.chkAutoshrink_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkAutoshrink_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oDatabase.DBOption.AutoShrink = (chkAutoshrink = 1)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.chkAutoshrink_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.chkAutoshrink_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub





Private Sub cmdConnect_Click()
    On Error GoTo errHandler
    Connect
    chkAutoshrink = IIf(oDatabase.DBOption.AutoShrink, 1, 0)
    cmdDumpTriggers.Enabled = True
    cmdEXport.Enabled = True
    cmdTables.Enabled = True
    cmdShrink.Enabled = True
    cmdTableStats.Enabled = True
    cmdRebuildIndexes.Enabled = True
    cmdComp.Enabled = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.cmdConnect_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdConnect_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEXport_Click()
    On Error GoTo errHandler
    ExportScript
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.cmdEXport_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdEXport_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdOne_Click()
    On Error GoTo errHandler
Dim OpenHndl As Long
Dim pWNDW As Long
Dim lThreadId  As Long
Dim lProcessId As Long

    If MsgBox("Ensure all Papyrus II applications (including the print server) are closed on the server before clicking OK." & vbCrLf _
    & "You can click Cancel to leave this procedure.", vbOKCancel + vbInformation, "Warning") = vbCancel Then
        Exit Sub
    End If
    
    'Force ending of all remnant type processes belonging to PBKSC
    
    
    OpenHndl = FindWindow(vbNullString, "CENTRAL Application")
    If OpenHndl <> 0 Then
        Call SendMessage(pWNDW, WM_CLOSE, 0&, 0&)
    End If
    
    
    FetchFiles
    MsgBox "The Fetch operation has finished", vbInformation + vbOKOnly, "Status"
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdOne_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRebuildIndexes_Click()
    On Error GoTo errHandler
Dim oTable As SQLDMO.Table
    For Each oTable In oDatabase.Tables
        If Not oTable.SystemObject Then oTable.RebuildIndexes
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdRebuildIndexes_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdShrink_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    oDatabase.Shrink 10, SQLDMOShrink_Default
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdShrink_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdTables_Click()
    On Error GoTo errHandler
Dim s As String
Dim Flag As SQLDMO_SCRIPT_TYPE
Dim oTable As SQLDMO.Table
Dim oStoredProc As SQLDMO.StoredProcedure2
Dim oView As SQLDMO.View2
Dim oUser As SQLDMO.User
Dim oUDF As SQLDMO.UserDefinedFunction
Dim oDBRole As SQLDMO.DatabaseRole2
Dim srtrs As ADODB.Recordset
Dim sTmp As String
Dim objDMO  As z_DMO

    Set objDMO = New z_DMO
    objDMO.Component oDatabase
    Screen.MousePointer = vbHourglass
    objDMO.CreateTableScript strFolderOut
    Screen.MousePointer = vbDefault
    MsgBox "Done"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.cmdTables_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdTables_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdTableStats_Click()
    On Error GoTo errHandler
Dim oTable As SQLDMO.Table
Dim oStats As TableStats
Dim rs As ADODB.Recordset
Dim frm As New frmTableSTats

    Set rs = New ADODB.Recordset
    rs.Fields.Append "Name", adVarChar, 40
    rs.Fields.Append "DataSpaceUsed", adVarChar, 30
    rs.Fields.Append "IndexSpaceUsed", adVarChar, 30
    rs.Fields.Append "Rows", adVarChar, 30
    rs.Open
    For Each oTable In oDatabase.Tables
        rs.AddNew
        rs.Fields("Name") = oTable.Name
        rs.Fields("DataSpaceUsed") = oTable.DataSpaceUsed
        rs.Fields("IndexSpaceUsed") = oTable.IndexSpaceUsed
        rs.Fields("Rows") = oTable.Rows
        rs.Update
    Next
    frm.Component rs
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdTableStats_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdThree_Click()
    On Error GoTo errHandler
Dim objDMO As New z_DMO
    Screen.MousePointer = vbHourglass
    lg "Connecting . . . "
    Connect
    objDMO.Component oDatabase
    lg "Creating table script . . . "
    objDMO.CreateTableScript strFolderOut
    lg "Creating trigger script . . . "
    objDMO.CreateTriggerScript strFolderOut
    Disconnect
    lg "Transmitting to support . . . " & FTPAddress & " : " & FTPFolder
    ManageTransmit True
    Screen.MousePointer = vbDefault
    lg "Finished "
    MsgBox "Files sent to support", , "Status"
'errHandler:
'    ErrPreserve
'    Screen.MousePointer = vbDefault
'    Disconnect
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.cmdThree_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdThree_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdUpdateFromScript_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    HandleScript
    Screen.MousePointer = vbDefault
'errHandler:
'    ErrPreserve
'    Screen.MousePointer = vbDefault
'
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.cmdUpdateFromScript_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdUpdateFromScript_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub CommandButton1_Click()
    On Error GoTo errHandler
MsgBox "Hello"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.CommandButton1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    cmdDumpTriggers.Enabled = False
    cmdEXport.Enabled = False
    cmdTables.Enabled = False
    cmdShrink.Enabled = False
    cmdTableStats.Enabled = False
    cmdRebuildIndexes.Enabled = False
    cmdComp.Enabled = False
    lblOne.Caption = "Fetch the latest update files from the support site."
    lblThree.Caption = "Send database script to support site"
    Me.SSTab1.Tab = 0
    Initializesettings
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Load", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If Not oSQLServer Is Nothing Then
        oSQLServer.Disconnect
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub LoadTriggers()
    On Error GoTo errHandler
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim oT As SQLDMO.Table
    Screen.MousePointer = vbHourglass

    k = 0
    For i = 1 To oDatabase.Tables.Count
        Set oT = oDatabase.Tables(i)
        For j = 1 To oT.Triggers.Count
            k = k + 1
            X1.ReDim 1, k, 1, 5
            X1(k, 1) = oT.Name
            X1(k, 2) = oT.Triggers(j).Name
            X1(k, 3) = oT.Triggers(j).Enabled
            X1(k, 4) = j
            X1(k, 5) = i
        Next j
    Next i
    G1.Array = X1
    G1.ReBind
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.LoadTriggers"
End Sub
Private Sub cmdDumpTriggers_Click()
    On Error GoTo errHandler

Dim str As String
Dim fs As New FileSystemObject
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim oT As SQLDMO.Table
Dim objDMO  As z_DMO

    Set objDMO = New z_DMO
    objDMO.Component oDatabase
    
    Screen.MousePointer = vbHourglass
    
    objDMO.CreateTriggerScript strFolderOut

    Screen.MousePointer = vbDefault
    MsgBox "Done"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdDumpTriggers_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub G1_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo errHandler
    oDatabase.Tables(X1(G1.Bookmark, 5)).Triggers(X1(G1.Bookmark, 4)).Enabled = G1.Text
    LoadTriggers
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.G1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Public Sub ManageTransmit(bIncludeScripts As Boolean)
    On Error GoTo errHandler
Dim lngResult As Long
Dim Res As Boolean
Dim zip
Dim FTP1 As New FTPClass
Dim fs As New FileSystemObject

    Screen.MousePointer = vbHourglass
'ZIP file to be sent============================================
    Set zip = CreateObject("FathZIP.FathZIPCtrl.1")
    If fs.FileExists(strServerMachineSharedFolder & "\BU\ERRORS.ZIP") Then
        fs.DeleteFile strServerMachineSharedFolder & "\BU\ERRORS.ZIP"
    End If
    zip.CreateZip strServerMachineSharedFolder & "\BU\ERRORS.ZIP", ""
    zip.ProcessSubfolders = True
    zip.BasePath = ""
    zip.PreservePaths = False
    zip.AddFile strServerMachineSharedFolder & "\Errors.txt", ""
    zip.AddFile strServerMachineSharedFolder & "\Printers\*.*", ""
    zip.AddFile strServerMachineSharedFolder & "\Transmit*.*", ""
    zip.AddFile strServerMachineSharedFolder & "\CENTRAL.INI", ""
    If bIncludeScripts Then
        zip.AddFile strFolderOut & "\*.*", ""
    End If
    If zip.LastError <> 0 Then
        MsgBox "Zipping errors file was not successful. Contact support person"
    End If
    zip.Close
    Set zip = Nothing
''''''''''''''''''''''''
    Set fINET = New wininet
    If strInternetDialup = "YES" Then
        lngResult = fINET.StartDUN(0, strConnectionName, True)
    End If
    
''OPEN FTP Connection===========================================
    Res = FTP1.OpenFTP(FTPAddress, FTPUsername, FTPPassword, True)    ', EXC_GENERAL, "Error opening FTP site"
    If Res Then
    ''Check Clientname folder on FTP exists
        If Not FTP1.FolderExists(FTPFolder & "/" & strClientCode & "*") Then
            Res = FTP1.CreatFolder(FTPFolder & "/" & strClientCode)   ', EXC_GENERAL, "Error creating folder on FTP"
            If Res = False Then
                lg "Cannot create folder " & FTPFolder & "/" & strClientCode
                Exit Sub
            End If
        End If
        If Not FTP1.FolderExists(FTPFolder & "/" & strClientCode & "/UP" & "*") Then
            Res = FTP1.CreatFolder(FTPFolder & "/" & strClientCode & "/UP")  ', EXC_GENERAL, "Error creating folder on FTP"
            If Res = False Then
                lg "Cannot create folder " & FTPFolder & "/" & strClientCode & "/UP"
                Exit Sub
            End If
        End If
        Res = FTP1.SetCurrentFolder(FTPFolder & "/" & strClientCode & "/UP") ', EXC_GENERAL, "Error setting FTP folder"
        If Res = False Then
            lg "Cannot set current folder " & FTPFolder & "/" & strClientCode & "/UP"
            Exit Sub
        End If
    '''''''''''''''''''
    'SEND File=======================================================
        Res = FTP1.PutFile(strServerMachineSharedFolder & "\BU\ERRORS.ZIP", strClientCode & "_ERR.ZIP", True)  ', EXC_GENERAL, "Error transmitting FTP file"
        If Res = False Then
            lg "Cannot put file " & strServerMachineSharedFolder & "\BU\ERRORS.ZIP"
            Exit Sub
        End If
    '''''''
    'Close FTP connection============================================
        FTP1.CloseFTP
    Else
        lg "Cannot create folder " & strServerMachineSharedFolder & "/LOYALTY/UP/    "
    End If
'    If fINET.IsNetConnectOnline Then
'Close Internet connection=======================================
    fINET.HangUp
''''''''''''''''''''''''''''
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ManageTransmit(bIncludeScripts)", bIncludeScripts
End Sub

Public Sub FetchFiles()
    On Error GoTo errHandler
Dim lngResult As Long
Dim Res As Boolean
Dim zip
Dim FTP1 As New FTPClass
Dim fs As New FileSystemObject
Dim F, fc, fol
Dim FTPFile As FTPFileClass
Dim strBUFolder As String
Dim cmd As ADODB.Command
Dim strPos As String

    lg "Backing up the database . . ."
    strPos = "1"
    Connect
    
    If fs.FolderExists(strLocalRootFolder & "\BU") Then
        strBUFolder = strLocalRootFolder & "\BU"
    Else
        strBUFolder = strLocalRootFolder
    End If
    strPos = "2"
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    Set cmd.ActiveConnection = ADOConn
    oDatabase.Shrink 10, SQLDMOShrink_Default
    If fs.FileExists(strBUFolder & "PBKSC.BAK") Then fs.DeleteFile strBUFolder & "PBKSC.BAK", True
    If fs.FileExists(strBUFolder & "PBKSCMASTER.BAK") Then fs.DeleteFile strBUFolder & "PBKSCMASTER.BAK", True
 
    strPos = "3"
 '   Timer1.Enabled = True
    cmd.CommandType = adCmdText
    cmd.CommandText = "BACKUP DATABASE PBKSC to disk = '" & strBUFolder & "PBKSC.BAK' WITH INIT, NAME = 'Full Backup of PBKSCDATA'"
    cmd.Execute
    cmd.CommandText = "BACKUP DATABASE MASTER to disk = '" & strBUFolder & "PBKSCMASTER.BAK' WITH INIT, NAME = 'Full Backup of PBKSCMASTER'"
    cmd.Execute
 '   Timer1.Enabled = False
    Set cmd = Nothing
    
    If Not (fs.FileExists(strLocalRootFolder & "\BU\" & "PBKSC.BAK") And fs.FileExists(strLocalRootFolder & "\BU\" & "PBKSCMASTER.BAK")) Then
        MsgBox "Backup was not successful.Contact support person"
        Exit Sub
    End If
    strPos = "4"
    lg vbCrLf & "Connecting . . ."
    Set fINET = New wininet
 '   If Not fINET.IsNetConnectOnline Then
    If strInternetDialup = "YES" Then
        lngResult = fINET.StartDUN(0, strConnectionName, True)
    End If
       ' Check lngResult = 0, ERR_DUNALREADYOPEN, "Cannot open connection,perhaps it is already open"
    lg "Opening FTP connection . . ."
''OPEN FTP Connection===========================================
    Check FTP1.OpenFTP(FTPAddress, FTPUsername, FTPPassword, True), EXC_GENERAL, "Opening FTP"
'''''''''''''''''''

    strPos = "5"
'Clear all old files in receiving folder
    Set fol = fs.GetFolder(strDownloadFolder)
    If fol.Files.Count > 0 Then
        fs.DeleteFile strDownloadFolder & "\*.*", True
    End If
'Fetch Files=======================================================
'Look for files in the 'Common' folder and in the folder named after the client (as in PBKSC.INI)
'Fetch all these into the download folder on the client's server machine
    lg "Fetching files in Common. . ." & FTPFolder & "/COMMON"
    Check FTP1.SetCurrentFolder(FTPFolder & "/COMMON"), EXC_GENERAL, "setting FTP folder"
    For Each FTPFile In FTP1.Files
        lg ". . . " & FTPFile.FileName
       Check FTP1.GetFile(FTPFile.FileName, strDownloadFolder & "\" & FTPFile.FileName, True), EXC_GENERAL, "Getting FTP file"
    Next
'    lg "Fetching files in Client. . ." & FTPFolder & "/" & strClientCode & "/DOWN"
'    Res = FTP1.SetCurrentFolder(FTPFolder & "/" & strClientCode & "/DOWN") ', EXC_GENERAL, "setting FTP folder: " & strClientCode   'FTPFolder & "/" & strClientCode &
'    If Res Then
'        For Each FTPFile In FTP1.Files
'            lg ". . . " & FTPFile.FileName
'           Check FTP1.GetFile(FTPFile.FileName, strDownloadFolder & "\" & FTPFile.FileName, True), EXC_GENERAL, "Getting FTP file"
'        Next
'    End If
    strPos = "6"
'unzip all zip files
    lg vbCrLf & "Unzipping. . ."
    Set fol = fs.GetFolder(strDownloadFolder)
    Set fc = fol.Files
    For Each F In fc
        If UCase(Right(F.Name, 4)) = ".ZIP" Then
            Set zip = CreateObject("FathZIP.FathZIPCtrl.1")
            zip.OpenZip (F.Path)
            zip.BasePath = strDownloadFolder
            zip.ExtractFile ("*.*")
            zip.Close
            Set zip = Nothing
            F.Delete True
        End If
    Next
    strPos = "7"
'Distribute them according to their type as follows
'.EXE .DLL go to folder:Patches;  .DOT go to folder:templates;  .SQL stay in download folder; .NOT stay in download folder
'Display progress in test box and when downloading is complete display contents of .NOT file if it exists
    lg "Moving and registering files . . ."
    
    'copies to folders and registers DLLs
    HandleDownload
    
    strPos = "8"
    'Runs SQL scripts if they exist
    lg "Updating database . . ."
    HandleScript
    strPos = "9"
    
    
    
Dim oTF As z_TextFile
    Set oTF = New z_TextFile
    oTF.OpenTextFile strServerMachineSharedFolder & "\UpdateLog.txt"
    oTF.WriteToTextFile Trim(txtResults)
    oTF.CloseTextFile
    
    Check FTP1.SetCurrentFolder(FTPFolder & "/" & strClientCode & "/UP"), EXC_GENERAL, "Error setting FTP folder"
'''''''''''''''''''
'SEND File=======================================================
    If fs.FileExists(strServerMachineSharedFolder & "\UpdateLog.txt") Then
        If FTP1.FileExists("UpdateLog.txt") Then
            FTP1.DeleteFile ("UpdateLog.txt")
        End If
        Res = FTP1.PutFile(strServerMachineSharedFolder & "\UpdateLog.txt", "UpdateLog.txt", True)
        Check Res, EXC_GENERAL, "Error transmitting FTP file"
        If FTP1.FileExists("OSQL_LOG.txt") Then
            FTP1.DeleteFile ("OSQL_LOG.txt")
        End If
        If fs.FileExists(strServerMachineSharedFolder & "\OSQL_LOG.txt") Then
            Res = FTP1.PutFile(strServerMachineSharedFolder & "\OSQL_LOG.txt", "OSQL_LOG.txt", True)
            Check Res, EXC_GENERAL, "Error transmitting FTP file"
        End If
        If fs.FileExists(strServerMachineSharedFolder & "\OSQL2_LOG.txt") Then
            Res = FTP1.PutFile(strServerMachineSharedFolder & "\OSQL2_LOG.txt", "OSQL2_LOG.txt", True)
            Check Res, EXC_GENERAL, "Error transmitting FTP file"
        End If
    End If
    strPos = "10"
'Close FTP connection============================================
    lg "Closing connection. . ."
    FTP1.CloseFTP
'    If fINET.IsNetConnectOnline Then
'Close Internet connection=======================================
    fINET.HangUp
''''''''''''''''''''''''''''
    Set FTP1 = Nothing
    Disconnect
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.FetchFiles", , , , "Error position = ", Array(strPos)
End Sub






Private Sub lg(pText As String)
    On Error GoTo errHandler
    txtResults = txtResults & vbCrLf & pText
    txtResults.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.lg(pText)", pText
End Sub
Private Sub ClearLog()
    On Error GoTo errHandler
    txtResults = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ClearLog"
End Sub
Private Sub cmdExtract_Click()
    On Error GoTo errHandler
    HandleDownload
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdExtract_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub HandleDownload()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim F, fc, fol
Dim lngReturn As Long
Dim strErrMsg As String
Dim strNewName As String
Dim strErrPos As String


'Get names of all DLLs on the DownloadFolder shared folder on the server
    Set fol = fs.GetFolder(strDownloadFolder).Files
    strErrPos = "1"
'Unregister all files of the same names as the downloaded files in the PBKS\Executables folder on the workstation and rename or delete then
    For Each F In fol
        If fs.FileExists(strServerMachineSharedFolder & "\Executables\" & F.Name) Then
            If UCase(Right(F.Name, 4)) = ".DLL" Then
                If Not UnregisterComEx(strServerMachineSharedFolder & "\Executables\" & F.Name, lngReturn, strErrMsg) Then
                    lg "Cannot unregister " & strServerMachineSharedFolder & "\Executables\" & F.Name & "Procedure halted without completing." & vbCrLf & "Error message is: " & strErrMsg
                    Exit Sub
                End If
            End If
            strNewName = strServerMachineSharedFolder & "\Executables\" & "o" & F.Name
            If fs.FileExists(strNewName) Then
                fs.DeleteFile strNewName, True
            End If
            Name strServerMachineSharedFolder & "\Executables\" & F.Name As strNewName
        End If
    Next
    strErrPos = "2"
'Copy all the DLLs and EXEs on the Patches_S shared folder to the PBKS\Executables folder on the workstation
    Set fol = fs.GetFolder(strDownloadFolder).Files
    strErrPos = "3"
    For Each F In fol
        If UCase(Right(F.Name, 4)) = ".DLL" Or UCase(Right(F.Name, 4)) = ".EXE" Then
            fs.CopyFile strDownloadFolder & "\" & F.Name, strServerMachineSharedFolder & "\Executables\" & F.Name, True
        End If
    Next
    strErrPos = "4"
    If fs.FileExists(strDownloadFolder & "\CENTRALDLL.DLL") Then
        fs.CopyFile strDownloadFolder & "\CENTRALDLL.DLL", strServerMachineSharedFolder & "\Patches\", True
    End If
    strErrPos = "5"
    If fs.FileExists(strDownloadFolder & "\CENTRAL.EXE") Then
        fs.CopyFile strDownloadFolder & "\CENTRAL.EXE", strServerMachineSharedFolder & "\Patches\", True
    End If
    strErrPos = "8"
    If fs.FileExists(strDownloadFolder & "\CENTRAL.INI") Then
        fs.CopyFile strDownloadFolder & "\CENTRAL.INI", strServerMachineSharedFolder & "\", True
    End If
    strErrPos = "9"
'Register all DLLs on the workstation PBKS\Executables folder
'    Set fol = fs.GetFolder(strServerMachineSharedFolder & "\Executables").Files
    For Each F In fol
        If UCase(Right(F.Name, 4)) = ".DLL" Then
            If Not RegisterComEx(strServerMachineSharedFolder & "\Executables\" & F.Name, lngReturn, strErrMsg) Then
                lg "Cannot register " & strServerMachineSharedFolder & "\Executables\" & F.Name & " Procedure halted without completing." & vbCrLf & "Error message is: " & strErrMsg
            End If
        End If
    Next
    strErrPos = "10"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.HandleDownload", , , , strErrPos, Array(strErrPos)
    Exit Sub
errHandler:
    ErrPreserve
    If strErrPos > "4" And strErrPos < "9" Then
        lg "Cannot copy a file to \Patches or CENTRAL.INI to \PBKS_S." & vbCrLf & "Position is: " & strErrPos
        Resume Next
    End If
        
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.HandleDownload", , , , "Position", Array(strErrPos)
End Sub

Private Sub HandleScript()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim strSQL As String
Dim oTF As New z_TextFile
Dim strPath As String
Dim strMessages As String
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim QR As QueryResults2
Dim strCommand As String

    strPath = strDownloadFolder & "\UPDATES.SQL"
    
    strCommand = "OSQL.EXE -Usa -P -S" & strServerName & " -dCENTRAL -i" & strPath & " -o" & strServerMachineSharedFolder
    
    If fs.FileExists(strPath) Then
        lg "Updating using UPDATES.SQL executing . . ."
        ShellAndWait strCommand & "\OSQL_LOG.TXT", vbHide, False
    End If
    
    strPath = strDownloadFolder & "\REPLACEMENTS.SQL"
    strCommand = "OSQL.EXE -Usa -P -S" & strServerName & " -dCENTRAL -i" & strPath & " -o" & strServerMachineSharedFolder
    If fs.FileExists(strPath) Then
        lg "Updating using REPLACEMENTS.SQL executing . . ."
        ShellAndWait strCommand & "\OSQL2_LOG.TXT", vbHide, False
    End If
    Exit Sub
   
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.HandleScript"
End Sub
Private Function fRunningInIde() As Boolean
    On Error GoTo errHandler
Dim sClassName As String
Dim nStrLen    As Long

    '
    ' See if we're running in the IDE.
    '
    sClassName = String$(260, vbNullChar)
    nStrLen = GetClassName(Me.hwnd, sClassName, Len(sClassName))
    If nStrLen Then sClassName = Left$(sClassName, nStrLen)
    
    fRunningInIde = (sClassName = "ThunderFormDC")
   ' MsgBox sClassName & "    " & fRunningInIde
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.fRunningInIde"
End Function

