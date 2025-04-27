VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form frmMain 
   BackColor       =   &H00C8B9B3&
   Caption         =   "Papyrus II dispatcher"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10680
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStopStart 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Pause processing"
      Height          =   435
      Left            =   3705
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5130
      Width           =   2670
   End
   Begin VB.CommandButton cmdTest 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Test"
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
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6480
      Width           =   1380
   End
   Begin VB.CommandButton cmdOpenErrorLog 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&View error log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5130
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5355
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   6285
      Width           =   255
   End
   Begin VB.CommandButton cmdMinimize 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Minimize"
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
      Left            =   8805
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5220
      Width           =   1380
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   9763
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BackColor       =   14339533
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Email dispatcher"
      TabPicture(0)   =   "frmMain.frx":014A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "eResult"
      Tab(0).Control(1)=   "cmdGo"
      Tab(0).Control(2)=   "File1"
      Tab(0).Control(3)=   "cmdRefresh"
      Tab(0).Control(4)=   "cmdClear"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Printing"
      TabPicture(1)   =   "frmMain.frx":0166
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label41"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lvwPrinters"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkPreview"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkKeepCopies"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "File2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdClearPrint"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "CommonDialog1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Command2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdRefreshPrint"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmdDeletePrinter"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "AcroPDF"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "E.D.I."
      TabPicture(2)   =   "frmMain.frx":0182
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "G_EDI"
      Tab(2).Control(2)=   "cmdRefreshEDI"
      Tab(2).ControlCount=   3
      Begin AcroPDFLibCtl.AcroPDF AcroPDF 
         Height          =   2835
         Left            =   6900
         TabIndex        =   25
         Top             =   1485
         Width           =   3210
         _cx             =   5662
         _cy             =   5001
      End
      Begin VB.CommandButton cmdDeletePrinter 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Delete selected printer"
         Height          =   405
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4170
         Width           =   3615
      End
      Begin VB.CommandButton cmdRefreshPrint 
         BackColor       =   &H00C4BCA4&
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
         Picture         =   "frmMain.frx":019E
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4185
         Width           =   615
      End
      Begin VB.CommandButton cmdRefreshEDI 
         BackColor       =   &H00C4BCA4&
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
         Left            =   -74790
         Picture         =   "frmMain.frx":0528
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4170
         Width           =   630
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Find Printers"
         Height          =   255
         Left            =   6105
         TabIndex        =   15
         Top             =   4410
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7350
         Top             =   3660
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdClearPrint 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Clear folder"
         Height          =   390
         Left            =   1665
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4185
         Width           =   1290
      End
      Begin VB.TextBox Text2 
         Height          =   345
         Left            =   3750
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   4365
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   405
         Left            =   3660
         TabIndex        =   11
         Top             =   4140
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Clear folder"
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
         Left            =   -73410
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4155
         Width           =   1365
      End
      Begin VB.FileListBox File2 
         Height          =   3600
         Left            =   225
         Pattern         =   "*.html;*.xml"
         TabIndex        =   8
         Top             =   555
         Width           =   2730
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00C4BCA4&
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
         Left            =   -74760
         Picture         =   "frmMain.frx":08B2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4155
         Width           =   615
      End
      Begin VB.FileListBox File1 
         Height          =   3600
         Left            =   -74880
         Pattern         =   "*.xml"
         TabIndex        =   6
         Top             =   555
         Width           =   2760
      End
      Begin VB.CommandButton cmdGo 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -68190
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4230
         Width           =   1275
      End
      Begin VB.TextBox eResult 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   3555
         Left            =   -71895
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   585
         Width           =   4980
      End
      Begin VB.CheckBox chkKeepCopies 
         Alignment       =   1  'Right Justify
         Caption         =   "Keep copies of document files in local PBKS\BU folder"
         ForeColor       =   &H8000000D&
         Height          =   465
         Left            =   6870
         TabIndex        =   2
         Top             =   990
         Width           =   3120
      End
      Begin VB.CheckBox chkPreview 
         Alignment       =   1  'Right Justify
         Caption         =   "Preview only"
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   8340
         TabIndex        =   1
         Top             =   585
         Visible         =   0   'False
         Width           =   1650
      End
      Begin TrueOleDBGrid60.TDBGrid G_EDI 
         Height          =   3345
         Left            =   -74775
         OleObjectBlob   =   "frmMain.frx":0C3C
         TabIndex        =   17
         Top             =   795
         Width           =   7980
      End
      Begin MSComctlLib.ListView lvwPrinters 
         Height          =   3165
         Left            =   3105
         TabIndex        =   22
         Top             =   1020
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   5583
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Printer"
            Object.Width           =   7832
         EndProperty
      End
      Begin VB.Label Label41 
         Caption         =   "Printers usually available (grey indicates not presently connected)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   3285
         TabIndex        =   24
         Top             =   555
         Width           =   3360
      End
      Begin VB.Label Label1 
         Caption         =   "EDI transmission log"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   -73560
         TabIndex        =   19
         Top             =   540
         Width           =   6675
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuDebug 
         Caption         =   "Debug on"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuErrorLog 
         Caption         =   "View error log"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   +^{F2}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents zPrint As z_Printing
Attribute zPrint.VB_VarHelpID = -1
Public WithEvents zEMail As z_EMail
Attribute zEMail.VB_VarHelpID = -1
Public WithEvents zEDI As z_EDI_Transmission
Attribute zEDI.VB_VarHelpID = -1
Dim XEDI As New XArrayDB
Dim flgLoading As Boolean
Dim bCanClose As Boolean
Private WithEvents fInet As wininet
Attribute fInet.VB_VarHelpID = -1
Dim rsEDI As ADODB.Recordset
'manage tray ---------
Dim bSysTrayLoaded As Boolean
Dim SysTrayText As String
Dim nid As NOTIFYICONDATA

Private Sub AcroPDF_OnError()
MsgBox "Error in Acropdf"
End Sub

Private Sub cmdClear_Click()
Dim fs As New FileSystemObject
Dim fol
Dim fil
Dim f

    Set fol = fs.GetFolder(strSharedServerFolder & "\Emails\")
    Set fil = fol.files
    For Each f In fil
        f.Delete True
    Next
        File1.Refresh

End Sub

Private Sub cmdClearPrint_Click()
Dim fs As New FileSystemObject
Dim fol
Dim fil
Dim f

    Set fol = fs.GetFolder(strSharedServerFolder & "\Printing\")
    Set fil = fol.files
    For Each f In fil
        f.Delete True
    Next
        File2.Refresh

End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
'        cmdFindLogFile_Click
'    Shell "NOTEPAD.EXE '" & strFileName & "'", vbNormalFocus

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRefreshEDI_Click()
    Set rsEDI = zEDI.GetRecentTransmissions
    LoadEDIArray
End Sub

Private Sub cmdRefreshPrint_Click()
    File2.Refresh
End Sub


Private Sub cmdStopStart_Click()

    If zPrint.TimerEnabled Then
        cmdStopStart.Caption = "Resume processing"
    Else
        cmdStopStart.Caption = "Pause processing"
    End If
    zPrint.EnableTimer (Not zPrint.TimerEnabled)
    zEMail.EnableTimer (Not zEMail.TimerEnabled)
    zEDI.EnableTimer (Not zEDI.TimerEnabled)

End Sub

Private Sub cmdTest_Click()
chkPreview.Visible = Not chkPreview.Visible
chkPreview.Enabled = Not chkPreview.Enabled
End Sub

Private Sub Command1_Click()
Dim i As Integer
Dim p As Printer
Dim tmp As String
    For Each p In Printers
        If Right(p.DeviceName, Len(p.DeviceName) - InStrRev(p.DeviceName, "\")) = Text2 Then
            tmp = p.DeviceName
            Exit For
        End If
    Next

    Dim wshNetwork, strPrinterPath As String
    Set wshNetwork = CreateObject("WScript.Network")
    wshNetwork.SetDefaultPrinter tmp
    MsgBox tmp & " has been set as your default printer.", , "Action complete"
End Sub



Private Sub Command2_Click()
   Dim BeginPage, EndPage, NumCopies, Orientation, i
   ' Set Cancel to True.
   CommonDialog1.CancelError = True
   On Error GoTo errHandler
   ' Display the Print dialog box.
   CommonDialog1.ShowPrinter
   ' Get user-selected values from the dialog box.
   BeginPage = CommonDialog1.FromPage
   EndPage = CommonDialog1.ToPage
   NumCopies = CommonDialog1.Copies
   Orientation = CommonDialog1.Orientation
   For i = 1 To NumCopies
   ' Put code here to send data to your printer.
   Next
   'MsgBox CommonDialog1.FileName
   Exit Sub
errHandler:
   ' User pressed Cancel button.
   Exit Sub
End Sub


Private Sub Form_Terminate()
    Set zPrint = Nothing
    Set zEMail = Nothing
End Sub

Private Sub mnuDebug_Click()
    mnuDebug.Checked = Not mnuDebug.Checked
    mDebugmodeOn = mnuDebug.Checked
End Sub

Private Sub mnuErrorLog_Click()
Dim strPath As String

    strPath = strSharedServerFolder & "\Errors.txt"
    Shell "NOTEPAD.EXE '" & strPath & "'", vbNormalFocus

End Sub

Private Sub mnuExit_Click()
    cmdClose_Click
End Sub

Private Sub zEDI_STATUS(pMsg As String)
    Set rsEDI = zEDI.GetRecentTransmissions
    LoadEDIArray
End Sub

Private Sub zPrint_Status(MSG As String)
    If MSG = "COMPLETE" Then
        File2.Refresh
    End If
End Sub

Private Sub zEmail_Status(MSG As String)
    eResult = Right(eResult, 2000)
    eResult = MSG & vbCrLf & eResult
    DoEvents
End Sub
Private Sub zEMail_Complete()
    cmdGo.Enabled = True
    DoEvents
End Sub

Private Sub chkKeepCopies_Click()
    SaveSetting "PS", "OPTIONS", "KEEPDOCUMENTS", CStr(chkKeepCopies)
    zPrint.KeepCopies = (chkKeepCopies = 1)
End Sub

Private Sub chkPreview_Click()
    zPrint.Preview = (chkPreview = 1)
    SaveSetting "PS", "PrintingSettings", "PrintPreview", CStr(chkPreview)
End Sub

Private Sub cmdGo_Click()
Dim bOnline As Boolean
''Dim frm As frmManageconnection
Dim strMsg As Variant
Dim lngResult As Long

    Set fInet = New wininet
    If InternetDialup = True Then
        lngResult = fInet.StartDUN(0, strConnectionName, True)
    End If
    
    DoEvents
    Screen.MousePointer = vbHourglass
    
    cmdGo.Enabled = False
    bCanClose = False
    
    zEMail.SendMail
    
    cmdGo.Enabled = True
    bCanClose = True
    File1.Refresh
    
    lngResult = fInet.Hangup
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdRefresh_Click()
    File1.Refresh
End Sub
Private Sub Form_Initialize()
    On Error GoTo errHandler
    Set zPrint = New z_Printing
    Set zEMail = New z_EMail
    Set zEDI = New z_EDI_Transmission
    zEMail.SharedRootFolder = strSharedServerFolder
    zEMail.GetSettings
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Initialize"
End Sub

Private Sub Form_Load()
Dim fs As New FileSystemObject
    On Error GoTo errHandler
    If fs.FileExists(strSharedServerFolder & "\TEMPLATES\" & "Logo.jpg") Then
        mApproLogoFilename = strSharedServerFolder & "\TEMPLATES\" & "Logo.jpg"
    ElseIf fs.FileExists(strSharedServerFolder & "\TEMPLATES\" & "Logo.bmp") Then
        mApproLogoFilename = strSharedServerFolder & "\TEMPLATES\" & "Logo.bmp"
    End If
    
    If Not fs.FolderExists(strSharedServerFolder & "\EDI") Then
        fs.CreateFolder strSharedServerFolder & "\EDI"
    End If
    If Not fs.FolderExists(strSharedServerFolder & "\EDI\POs") Then
        fs.CreateFolder strSharedServerFolder & "\EDI\POs"
    End If
    If Not fs.FolderExists(strSharedServerFolder & "\EDI\POs\OUT") Then
        fs.CreateFolder strSharedServerFolder & "\EDI\POs\OUT"
    End If
    chkKeepCopies = GetSetting("PS", "OPTIONS", "KEEPDOCUMENTS", 0)
    mnuDebug.Checked = False
    zPrint.KeepCopies = (chkKeepCopies = 1)
    zPrint.Preview = (1 = chkPreview)
    zPrint.InitializeManager strSharedServerFolder & "\TEMPLATES"
    File1.Path = strSharedServerFolder & "\Emails\"
    File1.Pattern = "*.HTML"     '= "*.HTML;*.XML"
    File2.Path = strSharedServerFolder & "\Printing\"
    File2.Pattern = "*.txt;*.xml"
    File1.Refresh
    cmdGo.Visible = InternetDialup
    bCanClose = True
    SSTab1.Tab = 1
    SetGridLayout Me.G_EDI, Me.Name
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not bCanClose
    If Not Cancel Then
        Cancel = MsgBox("You want to close the dispatch application?", vbQuestion + vbYesNo, "Confirm") = vbNo
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set zPrint = Nothing
    'Delete Icon from SysTray
    If bSysTrayLoaded Then UnloadSysTray
    SaveLayout Me.G_EDI, Me.Name

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


'====================================================================
' STUFF for TRAY
'====================================================================
Public Sub InitSysTray()
    On Error GoTo errHandler
    'the form must be fully visible before calling Shell_NotifyIcon
Dim Res As Boolean
Dim cnt As Integer
    SysTrayText = "Papyrus dispatcher running" & vbNullChar
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = SysTrayText
    End With
    cnt = 0
    Do While (Res = False) And (cnt < 60)
        Shell_NotifyIcon NIM_DELETE, nid
        Res = Shell_NotifyIcon(NIM_ADD, nid)
        If Res = False Then
            EventPause 1
            cnt = cnt + 1
        End If
    Loop
    bSysTrayLoaded = True
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.InitSysTray"
End Sub


Private Sub ChangeSysTray(IsRunning As Boolean)
Dim Res As Boolean
Dim cnt As Integer
    On Error GoTo errHandler
    If IsRunning Then
        SysTrayText = "Papyrus dispatcher running" & vbNullChar
    Else
        SysTrayText = "Papyrus dispatcher stopped" & vbNullChar
    End If
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = SysTrayText
    End With
    Do While (Res = False) And (cnt < 2)
        Res = Shell_NotifyIcon(NIM_MODIFY, nid)
        If Res = False Then
            EventPause 1
            cnt = cnt + 1
        End If
    Loop
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ChangeSysTray(IsRunning)", IsRunning
End Sub
Private Sub UnloadSysTray()
    On Error GoTo errHandler
Dim Res As Boolean
Dim cnt As Integer
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = ""
    End With
    
    Do While (Res = False) And (cnt < 2)
        Res = Shell_NotifyIcon(NIM_DELETE, nid)
        If Res = False Then
            EventPause 1
            cnt = cnt + 1
        End If
    Loop
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.UnloadSysTray"
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As _
                          Single, Y As Single)
    'On Error GoTo errHandler
    On Error Resume Next
    'this procedure receives the callbacks from the System Tray icon.
Dim Result As Long
Dim MSG As Long
    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        MSG = x
    Else
        MSG = x / Screen.TwipsPerPixelX
    End If
    Select Case MSG
'        Case WM_LBUTTONUP        '514 restore form window
'            Me.WindowState = vbNormal
'            Result = SetForegroundWindow(Me.hwnd)
'            Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
'        Case WM_RBUTTONUP        '517 display popup menu
'            Result = SetForegroundWindow(Me.hwnd)
'            Me.PopupMenu Me.menPopup
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_MouseMove(Button,Shift,X,Y)", Array(Button, Shift, x, Y), EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
    
    If Me.WindowState = vbMinimized Then
        Me.Hide
    Else
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdMinimize_Click()
Dim bOnline As Boolean
    Set fInet = New wininet
    bOnline = fInet.Connected
    Set fInet = Nothing
    If bOnline Then
        MsgBox "It appears that the internet connection has not disconnected. Please check and disconnect if desired.", vbInformation + vbOKOnly, "Warning"
    End If
    InitSysTray
    On Error GoTo errHandler
    Me.WindowState = vbMinimized
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdMinimize_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub LoadEDIArray()
    On Error GoTo errHandler
Dim i As Integer
    XEDI.Clear
    XEDI.ReDim 1, rsEDI.RecordCount, 1, 10
    For i = 1 To rsEDI.RecordCount
        With rsEDI
            XEDI.Value(i, 1) = FNS(rsEDI.fields(2))
            XEDI.Value(i, 2) = Format(FND(rsEDI.fields(1)), "dd/yy/yy HH:NN")
           ' XEDI.Value(i, 3) = ""
            XEDI.Value(i, 3) = FNS(rsEDI.fields(3))
            XEDI.Value(i, 4) = FNS(rsEDI.fields(4))
            XEDI.Value(i, 5) = FNS(rsEDI.fields(5))
            XEDI.Value(i, 6) = FND(rsEDI.fields(1))
            XEDI.Value(i, 10) = ""
        End With
        rsEDI.MoveNext
    Next
    XEDI.QuickSort 1, XEDI.UpperBound(1), 6, XORDER_DESCEND, XTYPE_DATE
    Set G_EDI.Array = XEDI
    G_EDI.ReBind
    Exit Sub
errHandler:
    ErrorIn "frmMain.LoadEDIArray"
End Sub

Public Sub FillPrintersList()
10        On Error GoTo errHandler
      Dim itmList As ListItem
      Dim lngIndex As Long
20        Me.lvwPrinters.ListItems.Clear
30        For lngIndex = 1 To oPC.Configuration.Printers.Count
40            Set itmList = lvwPrinters.ListItems.Add(Key:=oPC.Configuration.Printers.Key(oPC.Configuration.Printers.ItemByOrdinalIndex(lngIndex)) & "k")
50            With itmList
60                .Text = oPC.Configuration.Printers.ItemByOrdinalIndex(lngIndex)
70                If oPC.Configuration.Printers.ActiveByOrdinal(lngIndex) = False Then
80                    .ForeColor = COLOR_CANCELLED
90                End If
100           End With
110       Next
120       Exit Sub
errHandler:
130       If ErrMustStop Then Debug.Assert False: Resume
140       ErrorIn "frmMain.FillPrintersList"
End Sub

Private Sub cmdDeletePrinter_Click()
    On Error GoTo errHandler
    
    oPC.Configuration.DeletePrinter CLng(Val(lvwPrinters.SelectedItem.Key))
    FillPrintersList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdDeletePrinter_Click", , EA_NORERAISE
    HandleError
End Sub

