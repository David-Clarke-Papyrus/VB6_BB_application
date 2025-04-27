VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Papyrus II  P.O.S. configuration settings"
   ClientHeight    =   6735
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12525
   Icon            =   "frmMain.frx":0000
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
      TabPicture(0)   =   "frmMain.frx":27A2
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
      TabPicture(1)   =   "frmMain.frx":27BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdClearMarketing"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Database"
      TabPicture(2)   =   "frmMain.frx":27DA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdUpdateData"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "CD1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.CommandButton cmdUpdateData 
         BackColor       =   &H00D7BDBD&
         Caption         =   "Run UPDATE_DATA"
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
         Left            =   -74475
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5100
         Width           =   3435
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   -69090
         Top             =   4200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
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
         Height          =   3870
         Left            =   -68205
         TabIndex        =   10
         Top             =   780
         Width           =   4560
         Begin VB.TextBox txtExchangeNumber 
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
            Left            =   1080
            TabIndex        =   19
            Top             =   2670
            Width           =   2385
         End
         Begin VB.CommandButton cmdResetExchangeNumber 
            BackColor       =   &H00D7BDBD&
            Caption         =   "Reset the exchange number"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   495
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   3150
            Width           =   3435
         End
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
            Format          =   156041217
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
            Top             =   1290
            Width           =   3435
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Next exchange number to issue"
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   270
            TabIndex        =   20
            Top             =   2445
            Width           =   3915
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "This does not clear the properties"
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   300
            TabIndex        =   17
            Top             =   2055
            Width           =   3915
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
            Top             =   2715
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
         OleObjectBlob   =   "frmMain.frx":27F6
         TabIndex        =   1
         Top             =   735
         Width           =   11205
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuDebug 
         Caption         =   "Debug on"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
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
Dim arCOMMANDLINE() As String
Dim cnPapyShort As ADODB.Connection
Dim strPassword As String
Dim oConn As z_POSConnection
'Dim oDatabase As SQLDMO.Database2
'Dim oSQLServer As SQLDMO.SQLServer2

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long


'Private Sub cmdBackup_Click()
'Dim fs As New FileSystemObject
'    Connecttodatabase
'    Check BackupFDToLocal(strLocalRootFolder), EXC_GENERAL, "Backup was not successful. Contact support person"
'    Check (fs.FileExists(strLocalRootFolder & "\PBKSFD.BAK")), EXC_GENERAL, "Backup was not successful.Contact support person"
'    MsgBox "The backup is complete. The file is to be found in " & strLocalRootFolder & "\PBKSFD.BAK"
'End Sub

Private Sub cmdLoadFromFile_Click()
    On Error GoTo errHandler
    If MsgBox("This will overwrite values that may have been already set in the properties of this workstation. You should only continue on the instructions of a Papyrus support person.", vbCritical + vbOKCancel, "Warning") = vbOK Then
        If MsgBox("Do not load data unless this is a new installation of POS. Click YES to go ahead NO to skip loading.", vbCritical + vbYesNo + vbDefaultButton2, "WARNING") = vbNo Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
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
    
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Screen.MousePointer = vbHourglass
    Connecttodatabase
    Set cmd = New ADODB.Command
    cmd.CommandText = "SP_PURGEPOSFILES"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oConn.DBConn
    
    Set par = cmd.CreateParameter("@PriorTo", adDate, adParamInput, , dtePriorTo)
    cmd.Parameters.Append par
    
    cmd.Execute
    Screen.MousePointer = vbDefault
    MsgBox "Database purged"
    
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


Private Sub cmdResetExchangeNumber_Click()
    If MsgBox("Confirm you want to reset the exchange number to " & Me.txtExchangeNumber & vbCrLf & "This action will delete all existing exchange transactions on the database.", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Connecttodatabase
    oConn.DBConn.Execute "UPDATE tAPPSETTINGS SET EX_NUMBER = " & CStr(CLng(txtExchangeNumber) - 1)
    oConn.DBConn.Execute "TRUNCATE TABLE tCSL"
    oConn.DBConn.Execute "TRUNCATE TABLE tPAYMENT"
    oConn.DBConn.Execute "TRUNCATE TABLE tEXCHANGE"
    oConn.DBConn.Execute "TRUNCATE TABLE tOpSession"
    oConn.DBConn.Execute "TRUNCATE TABLE tZSESSION"
    MsgBox "Done", vbOKOnly, "Done"
End Sub

'Private Sub cmdRestore_Click()
'Dim strFilename As String
'    Connecttodatabase
'    CD1.DialogTitle = "Find .BAK file to restore"
'    CD1.DefaultExt = "BAK"
'    CD1.InitDir = strLocalRootFolder
'    CD1.CancelError = True
' On Error GoTo CANCELERROR_ROUTINE
'    CD1.ShowOpen
' On Error GoTo 0
'
'    Screen.MousePointer = vbHourglass
'    strFilename = CD1.FileName
'    RestoreFDDatabase strFilename
'
'    Screen.MousePointer = vbDefault
'    MsgBox "The PBKSFD database has been created from the backup file: " & strFilename, vbOKOnly, "Status"
'
'
'CANCELERROR_ROUTINE:
'    Exit Sub
'End Sub
'
'
'Private Sub cmdSaveToFile_Click()
'    On Error GoTo errHandler
'    SaveRecordset
'    MsgBox "File saved to " & oConn.SharedFolderRoot & "\BU\Props.xml", vbOKOnly
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.cmdSaveToFile_Click", , EA_NORERAISE
'    HandleError
'End Sub


Private Sub cmdUpdateData_Click()
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Screen.MousePointer = vbHourglass
    Connecttodatabase
    Set cmd = New ADODB.Command
    cmd.CommandText = "UPDATE_DATA"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oConn.DBConn
    
    
    cmd.Execute
    Screen.MousePointer = vbDefault
    MsgBox "UPDATA_DATA run complete"

End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim strDataSource As String
    
    arCOMMANDLINE = Split(Command(), " ")
    Set oConn = New z_POSConnection

    
    If UBound(arCOMMANDLINE) > 0 Then
        oConn.DatabaseName = arCOMMANDLINE(0)
    Else
        oConn.DatabaseName = ""
    End If
    
    InitializeSettings

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub OPenConnection()

End Sub

Private Sub Connecttodatabase()
    oConn.dbConnect strLocalServername, strPassword
    loadGrid
End Sub

Private Sub InitializeSettings()
    On Error GoTo errHandler
Dim strPCName As String
Dim fs As New FileSystemObject

    strPCName = Trim(Me.NameOfPC)
        
    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & strPCName & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
    End If
    
    If Not fs.FolderExists(strLocalRootFolder) Then
        MsgBox "Local Root Folder (" & strLocalRootFolder & ") does not exist, connection will not be made to the database", vbCritical, "Can't continue"
    End If
        
    If Not fs.FileExists(strLocalRootFolder & "\PBKSWS.INI") Then
        MsgBox strLocalRootFolder & "\PBKSWS.INI" & " does not exist, connection will not be made to the database", vbCritical, "Can't continue"
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

Private Sub mnuDebug_Click()
    mnuDebug.Checked = Not mnuDebug.Checked
    bDebug = mnuDebug.Checked
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub


'Public Function BackupFDToLocal(strLocalRootFolder As String) As Boolean
'Dim strPos As String
'Dim strBUFolder As String
'Dim cmd As ADODB.Command
'
'Dim fs As New FileSystemObject
'    On Error GoTo errHandler
'
'    If fs.FolderExists(strLocalRootFolder) Then
'        strBUFolder = strLocalRootFolder & "\"
'    Else
'        strBUFolder = strLocalRootFolder & "\"
'    End If
' Screen.MousePointer = vbHourglass
'    Set cmd = New ADODB.Command
'    cmd.CommandTimeout = 0
'    Set cmd.ActiveConnection = oConn.DBConn
'    If fs.FileExists(strBUFolder & "PBKSFD.BAK") Then fs.DeleteFile strBUFolder & "PBKSFD.BAK", True
'
'    cmd.CommandType = adCmdText
'    cmd.CommandText = "BACKUP DATABASE PBKSFD to disk = '" & strBUFolder & "PBKSFD.BAK' WITH INIT, NAME = 'Full Backup of PBKSFDDATA'"
'    cmd.Execute
' Screen.MousePointer = vbDefault
'    Set cmd = Nothing
'    BackupFDToLocal = True
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_PBKSBackup.BackupFDToLocal", , , , "position,BackupFolder", Array(strPos, strBUFolder)
'End Function
'
'Public Sub RestoreFDDatabase(pFilename As String, Optional pName As String = "PBKSFD")
'Dim oRestore As New SQLDMO.Restore
'    On Error Resume Next
'
'    oConn.dbCloseConnect
'    Set oSQLServer = New SQLDMO.SQLServer
'    oSQLServer.LoginTimeout = 0 '-1 is the ODBC default (60) seconds
'    With oSQLServer
'        .LoginSecure = False
'        .AutoReConnect = False
'        .Connect strLocalServername, "sa", "car"
'    End With
'
' '   Set oDatabase = oSQLServer.Databases("PBKSFD")
'
'
'
'    oSQLServer.DetachDB pName
'    On Error GoTo errHandler
'
'    With oRestore
'    'this is where your backup will be restored to
'    .Database = pName
'    'same as EM or TSQL, you can restore database, file, or log, here we're going to
'    'use database
'    .Action = SQLDMORestore_Database
'    'this is the "force restore over existing database" option
'    .ReplaceDatabase = True
'    'this does a restore from a file instead of a device - note that we're still
'    'restoring a database, NOT a file group
'    .RelocateFiles = "[PBKSFD_Data]" + "," + "[C:\PBKS\DATA\PBKSFD_DATA.mdf]" _
'        + "," + "[PBKSFD_LOG]" + "," + "[C:\PBKS\DATA\PBKSFD_LOG.ldf]"
'
'    .Files = pFilename
'    'do it
'    .SQLRestore oSQLServer
'    End With
'
'    'standard clean up
'    Set oRestore = Nothing
'    oSQLServer.Disconnect
'    Set oSQLServer = Nothing
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_SQLDMO.RestoreFDDatabase(pFilename)", pFilename
'End Sub


Private Sub txtExchangeNumber_Change()
    cmdResetExchangeNumber.Enabled = IsNumeric(txtExchangeNumber)
End Sub

