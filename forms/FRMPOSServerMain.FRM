VERSION 5.00
Begin VB.Form frmPosServerMain 
   BackColor       =   &H0080443E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POS Server"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3945
   ControlBox      =   0   'False
   FillColor       =   &H0080443E&
   Icon            =   "FRMPOSServerMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUtilities 
      BackColor       =   &H00BBBBA8&
      Caption         =   "&Utilities"
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
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1935
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   105
      TabIndex        =   5
      Top             =   165
      Width           =   3735
      Begin VB.Label lblUpdates 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2430
         TabIndex        =   11
         Top             =   375
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Number of Updates sent"
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
         Height          =   240
         Left            =   180
         TabIndex        =   10
         Top             =   375
         Width           =   2100
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Number of Sales Saved"
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
         Height          =   195
         Left            =   345
         TabIndex        =   9
         Top             =   120
         Width           =   2010
      End
      Begin VB.Label lblNumSales 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2430
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblOther 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2430
         TabIndex        =   7
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Other Requests processed"
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
         Height          =   210
         Left            =   90
         TabIndex        =   6
         Top             =   630
         Width           =   2265
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox ImgList 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   525
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   13
      Top             =   -60
      Width           =   1200
   End
   Begin VB.CheckBox chkRun 
      BackColor       =   &H00BBBBA8&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   480
      Left            =   675
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1245
      Width           =   1035
   End
   Begin VB.CommandButton cmdNewRecSet 
      Caption         =   "&New Recordset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   885
      TabIndex        =   2
      Top             =   1695
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CommandButton cmdEditClient 
      Caption         =   "&Edit Client List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   285
      TabIndex        =   1
      Top             =   1515
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00BBBBA8&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2730
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1740
      Width           =   1035
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   15
      Top             =   -135
   End
   Begin VB.Label lblRun 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080443E&
      Caption         =   "Server is stopped"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1725
      TabIndex        =   4
      Top             =   1335
      Width           =   2040
   End
   Begin VB.Menu menPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu menStop 
         Caption         =   "POS Server - Stop"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInboxName 
         Caption         =   "Show InBox Name"
         Shortcut        =   ^I
      End
      Begin VB.Menu menClose 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmPosServerMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim WithEvents oSFQ As clsServerFQ
Attribute oSFQ.VB_VarHelpID = -1

Dim nid As NOTIFYICONDATA

Dim bSysTrayLoaded As Boolean



Private Sub cmdUtilities_Click()
Dim frm As New frmPOSSVRUtilities
    Timer1.Enabled = False
    frm.Show vbModal
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Not Initialize Then Exit Sub
    Me.KeyPreview = True
    InitSysTray
'    Set oHR = New HandleRequest
    Me.chkRun.Value = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set oSFQ = Nothing
    'Delete Icon from SysTray
    If bSysTrayLoaded Then UnloadSysTray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Private Sub oSFQ_FileProcessed(FileType As Integer)
    On Error GoTo errHandler
    Select Case FileType
        Case 0
            Me.lblOther = val(lblOther) + 1
        Case 1
            Me.lblNumSales = val(lblNumSales) + 1
        Case 2
            Me.lblUpdates = val(lblUpdates) + 1
    End Select
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.oSFQ_FileProcessed(FileType)", FileType, EA_NORERAISE
    HandleError
End Sub

Private Sub Timer1_Timer()
On Error GoTo errHandler
    
    Timer1.Enabled = False
    
    oSFQ.PollForNewClient
    oSFQ.FetchClientSalesandUpdateServer
    oSFQ.PutMasterDataInOutbox
    oSFQ.TransmitUpdatesToClients
    
    Timer1.Enabled = True
    
    Exit Sub
    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.Timer1_Timer", , EA_NORERAISE
    HandleError
End Sub


Private Function Initialize()
    On Error GoTo errHandler
Dim msg As String
Dim i As Integer
'Dim fInit As frmInitialize

    'check if we got server path
    Set oSFQ = New clsServerFQ
    If Command() <> "" Then
        Me.BackColor = vbRed
    End If
    If Not oSFQ.DBConnectionOK Then
        msg = "Can't load Server Database!" & vbLf & _
               "Make sure the ODBC Manager contains an entry by name:" & vbLf & _
               "'" & oSFQ.DBName & "' whitch points to the POS Server Database."
        GoTo errHandler
    End If
    If Not oSFQ.FoldersOK Then
        msg = "Can't access shared InBox Folder."
        GoTo errHandler
    End If
    i = oSFQ.LoadClientList
    If i = 0 Then
        MsgBox "Client list is empty!", vbOKOnly + vbCritical, "No Clients loaded yet!"
    ElseIf i = -99 Then
        msg = "Can't load client list from Database!"
        GoTo errHandler
    End If
    Initialize = True
MEX:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.Initialize"
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If Shift = 2 And KeyCode = vbKeyI Then
        mnuInboxName_Click
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.Form_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub chkRun_Click()
    On Error GoTo errHandler
    With Me.chkRun
        If .Value = 1 Then
            .Caption = "Stop"
            Me.lblRun.Caption = "Server is running"
            Me.lblRun.ForeColor = vbGreen
            Me.Timer1.Enabled = True
          '  Set Me.Icon = Me.ImgList.ListImages(1).Picture
            ChangeSysTray True
        ElseIf .Value = 0 Then
            .Caption = "Start"
            Me.lblRun.Caption = "Server is stopped"
            Me.Timer1.Enabled = False
         '   Set Me.Icon = Me.ImgList.ListImages(2).Picture
            ChangeSysTray False

        End If
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.chkRun_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
 End
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEditClient_Click()
    On Error GoTo errHandler
Dim fCL As New frmClientList
    With fCL
        .Component oSFQ
        .Show vbModal
    End With
MEX:
    Unload fCL
    Set fCL = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.cmdEditClient_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub menClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.menClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub menStop_Click()
    On Error GoTo errHandler
    If Me.chkRun.Value = 1 Then
        Me.chkRun.Value = 0
        Me.menStop.Caption = "POS Server - Start"
    Else
        Me.chkRun.Value = 1
        Me.menStop.Caption = "POS Server - Stop"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.menStop_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub mnuInboxName_Click()
    On Error GoTo errHandler
    MsgBox oSFQ.SharedInboxName, vbOKOnly + vbInformation, "Shared InBox Name"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.mnuInboxName_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub StartListener()
'Dim QI As MSMQQueueInfo
'
'    Set QI = New MSMQQueueInfo
'    QI.PathName = "urs\private$\PapyPos_Queue"
'    Set ReqQ = QI.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
'    Set ReqEvent = New MSMQEvent
'    ReqQ.EnableNotification ReqEvent
'End Sub



'Sub ReqEvent_Arrived(ByVal Queue As Object, ByVal Cursor As Long)
'Dim qReqQ As MSMQQueue
'Dim qMsg As MSMQMessage
'Dim lID As Long
'Dim sISBN As String
'Dim xVal()
'
'
'    Set qReqQ = Queue
'    Set qMsg = qReqQ.Receive(ReceiveTimeOut:=0)
'    If Not qMsg Is Nothing Then
'        xVal = Split(qMsg.Body, "=")
'
'        If xVal(0) = "IDNum" Then
'            lID = Val(xVal(1))
'        Else
'            sISBN = xVal(1)
'        End If
''        GetRequest lID, sISBN
'    End If
'    ReqQ.EnableNotification ReqEvent
'End Sub


'Stuff for System Tray functionality~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Sub InitSysTray()
    On Error GoTo errHandler
    'the form must be fully visible before calling Shell_NotifyIcon
    Me.Show
    Me.Refresh
    SysTrayText = "POS Server: Running" & vbNullChar
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = SysTrayText
    End With
    Shell_NotifyIcon NIM_ADD, nid
    bSysTrayLoaded = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.InitSysTray"
End Sub

Private Sub ChangeSysTray(IsRunning As Boolean)
    On Error GoTo errHandler
    If IsRunning Then
        SysTrayText = "POS Server: Running" & vbNullChar
    Else
        SysTrayText = "POS Server: Stopped" & vbNullChar
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
    Shell_NotifyIcon NIM_MODIFY, nid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.ChangeSysTray(IsRunning)", IsRunning
End Sub
Private Sub UnloadSysTray()
    On Error GoTo errHandler
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = ""
    End With
    Shell_NotifyIcon NIM_DELETE, nid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.UnloadSysTray"
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As _
                          Single, Y As Single)
    On Error GoTo errHandler
    'this procedure receives the callbacks from the System Tray icon.
Dim Result As Long
Dim msg As Long
    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
            Result = SetForegroundWindow(Me.hwnd)
            Me.PopupMenu Me.menPopup
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.Form_MouseMove(Button,Shift,X,Y)", Array(Button, Shift, X, Y), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
    'this is necessary to assure that the minimized window is hidden
    If Me.WindowState = vbMinimized Then Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub mPopRestore_Click()
    On Error GoTo errHandler
Dim r As Long
    'called when the user clicks the popup menu Restore command
    Me.WindowState = vbNormal
    r = SetForegroundWindow(Me.hwnd)
    Me.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.mPopRestore_Click", , EA_NORERAISE
    HandleError
End Sub
