VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   Caption         =   "Papyrus II  P.O.S. configuration settings"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
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
      TabPicture(0)   =   "frmMain1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Grid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdReload"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdOpen"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Marketing"
      TabPicture(1)   =   "frmMain1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdClearMarketing"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Database"
      TabPicture(2)   =   "frmMain1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton cmdOpen 
         BackColor       =   &H00D7BDBD&
         Caption         =   "Connect to database and show properties"
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
         Left            =   225
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
            Format          =   16121857
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
         OleObjectBlob   =   "frmMain1.frx":0054
         TabIndex        =   1
         Top             =   735
         Width           =   11205
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
Dim strLocalRootFolder As String
Dim strLocalServername As String
Dim cnPapyShort As ADODB.Connection
Dim strPassword As String
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long



Private Sub cmdClearMarketing_Click()
    On Error GoTo errHandler
Dim oDB As New z_POSConnection
    If MsgBox("Confirm you want to delete the marketing rules?", vbQuestion + vbYesNo, "Warning") = vbNo Then
        Exit Sub
    End If
    oDB.dbConnect strLocalServername, strPassword
    oDB.DBConn.Execute "DELETE FROM tMARKETING"
    MsgBox "Done", vbInformation + vbOKOnly, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdClearMarketing_Click", , EA_NORERAISE
    HandleError
End Sub



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
End Sub

Private Sub cmdPurge_Click()
    If MsgBox("Confirm you want to purge the local (secondary) database prior to reloading from the POS server?", vbQuestion + vbYesNo, "Warning") = vbNo Then
        Exit Sub
    End If
    
    
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_PurgePOSFiles"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PriorTo", adDate, adParamInput, , Me.dtePriorTo)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oConn.DBConn
    cmd.Execute
    
    Set cmd = Nothing
    
    

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

Private Sub cmdRestore_Click()
Dim oSQLServer As SQLDMO.SQLServer2
Dim oRestore As SQLDMO.Restore
Dim fs As New FileSystemObject

    MsgBox "Ensure the POS application is not running before continuing", vbInformation, "Warning"
    Set oSQLServer = New SQLDMO.SQLServer2
    oSQLServer.LoginTimeout = 0 '-1 is the ODBC default (60) seconds
    With oSQLServer
        .LoginSecure = False
        .AutoReConnect = False
        .Connect strLocalServername, "sa", ""
    End With
    oConn.dbCloseConnect
    On Error Resume Next
    oSQLServer.DetachDB txBUName
Set oRestore = New SQLDMO.Restore
 
'use the 'with' construct to minimize property lookups
With oRestore
    'this is where your backup will be restored to
    .Database = txBUName
    'same as EM or TSQL, you can restore database, file, or log, here we're going to use database
    .Action = SQLDMORestore_Database
    'this is the "force restore over existing database" option
    .ReplaceDatabase = True
    'this does a restore from a file instead of a device - note that we're still restoring a database, NOT a file group
    .Files = oConn.LocalRoot & "\BU\" & txBUName & ".BAK"

    .SQLRestore oSQLServer
End With
    Screen.MousePointer = vbDefault
    MsgBox "Restore complete.", vbInformation, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdBackup_Click", , EA_NORERAISE
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


Private Sub cmdBackup_Click()
Dim oBU As z_POSops
Dim fs As New FileSystemObject

On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    Set oBU = New z_POSops
    Check oBU.BackupToLocal(txBUName), EXC_GENERAL, "Backup was not successful.Contact support person"
    Set oBU = Nothing
    Screen.MousePointer = vbDefault
    MsgBox "Backup complete.", vbInformation, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdBackup_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim strDataSource As String

    InitializeSettings
'    oConn.dbConnect strLocalServername
'    loadGrid

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Connecttodatabase()
    oConn.dbConnect strLocalServername, strPassword
    loadGrid
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

    
    strLocalServername = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "POSSQLServer", "")
    strPassword = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PASSWORD", "")
    
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
Dim lngArrayRows As Integer
Dim lngIndex As Long
On Error Resume Next
    rs.Close
    On Error GoTo errHandler
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    lngArrayRows = 100
    rs.CursorLocation = adUseClient

    rs.Open "SELECT * FROM dbo.tProperty Order By PropertyKey", oConn.DBConn, adOpenKeyset, adLockOptimistic


    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, lngArrayRows, 1, 5
    lngIndex = 1
    Do While Not rs.EOF
   '         XA.Value(lngIndex, 1) = Trim(rs.Fields("Sequence"))
            XA.Value(lngIndex, 1) = Trim(rs.Fields("PropertyKey"))
            XA.Value(lngIndex, 2) = Trim(rs.Fields("PropertyValue"))
            XA.Value(lngIndex, 3) = Trim(rs.Fields("PropertyDescription"))
            XA.Value(lngIndex, 5) = rs.Bookmark
            lngIndex = lngIndex + 1
            rs.MoveNext
    Loop
    Grid1.Array = XA
    Grid1.ReBind
    Exit Sub


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
    rs.Open "SELECT * FROM dbo.tProperty Order By PropertyKey", oConn.DBConn, adOpenKeyset, adLockOptimistic

    Set rsSave = Nothing
    Set rsSave = New ADODB.Recordset
    rsSave.Open oConn.SharedFolderRoot & "\BU\Props.xml", "Provider=MSPersist;", adOpenForwardOnly, adLockBatchOptimistic, adCmdFile

    rsSave.MoveFirst
    Do While Not rsSave.EOF
        rs.Bookmark = rsSave.Bookmark
       ' rs.Fields(0) = rsSave.Fields(0)
        rs.Fields(1) = rsSave.Fields(1)
        rs.Fields(2) = rsSave.Fields(2)
     '   rs.Fields(3) = rsSave.Fields(3)
        rs.Update
        rsSave.MoveNext
    Loop
    rsSave.Close
    Set rsSave = Nothing
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
    rs.Bookmark = XA.Value(Grid1.Bookmark, 5)
    If ColIndex = 1 Then
        rs.Fields("PROPERTYVALUE") = Trim(Grid1.Text)
        rs.Update
    ElseIf ColIndex = 2 Then
        rs.Fields("PROPERTYDESCRIPTION") = Trim(Grid1.Text)
        rs.Update
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Grid1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

