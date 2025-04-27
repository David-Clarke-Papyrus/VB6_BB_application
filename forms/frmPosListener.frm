VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPosListener 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POS Server Listener"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3945
   Icon            =   "frmPosListener.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   105
      TabIndex        =   5
      Top             =   -30
      Width           =   3735
      Begin VB.Label lblUpdates 
         Alignment       =   2  'Center
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
         Height          =   225
         Left            =   2430
         TabIndex        =   11
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Number of Updates send"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   10
         Top             =   405
         Width           =   2145
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   300
         TabIndex        =   9
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label lblNumSales 
         Alignment       =   2  'Center
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
         Height          =   225
         Left            =   2430
         TabIndex        =   8
         Top             =   165
         Width           =   1215
      End
      Begin VB.Label lblOther 
         Alignment       =   2  'Center
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
         Height          =   240
         Left            =   2430
         TabIndex        =   7
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   210
         Left            =   90
         TabIndex        =   6
         Top             =   675
         Width           =   2265
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   525
      Top             =   -255
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosListener.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosListener.frx":11A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkRun 
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
      ForeColor       =   &H00400040&
      Height          =   390
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1050
      Width           =   1245
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
      Left            =   255
      TabIndex        =   2
      Top             =   1980
      Visible         =   0   'False
      Width           =   1875
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
      Width           =   1875
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
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
      Left            =   2565
      TabIndex        =   0
      Top             =   1515
      Width           =   1035
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   15
      Top             =   -135
   End
   Begin VB.Label lblRun 
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
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1785
      TabIndex        =   4
      Top             =   1110
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
Attribute VB_Name = "frmPosListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim WithEvents oSFQ As clsServerFQ
Attribute oSFQ.VB_VarHelpID = -1

Dim nid As NOTIFYICONDATA

Dim bSysTrayLoaded As Boolean



Private Sub Form_Load()
    If Not Initialize Then Exit Sub
    Me.KeyPreview = True
    InitSysTray
'    Set oHR = New HandleRequest
    Me.chkRun.Value = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set oSFQ = Nothing
    'Delete Icon from SysTray
    If bSysTrayLoaded Then UnloadSysTray
End Sub



Private Sub oSFQ_FileProcessed(FileType As Integer)
    Select Case FileType
        Case 0
            Me.lblOther = Val(lblOther) + 1
        Case 1
            Me.lblNumSales = Val(lblNumSales) + 1
        Case 2
            Me.lblUpdates = Val(lblUpdates) + 1
    End Select
    
End Sub

Private Sub Timer1_Timer()
    On Error GoTo EH
    oSFQ.PollForNewClient
    oSFQ.PollSales
    oSFQ.PollDB
    oSFQ.SendUpdates
    Exit Sub
EH:
    MsgBox "Error in:" & vbLf & Err.Source & vbLf & vbLf & "Error: " & Err.Description
    Me.chkRun.Value = 0
End Sub


Private Function Initialize()
Dim msg As String
Dim i As Integer
'Dim fInit As frmInitialize

    'check if we got server path
    Set oSFQ = New clsServerFQ
    
    If Not oSFQ.DBConnectionOK Then
        msg = "Can't load Server Database!" & vbLf & _
               "Make sure the ODBC Manager contains an entry by name:" & vbLf & _
               "'" & oSFQ.DBName & "' whitch points to the POS Server Database."
        GoTo EH
    End If
    If Not oSFQ.FoldersOK Then
        msg = "Can't access shared InBox Folder."
        GoTo EH
    End If
    i = oSFQ.LoadClientList
    If i = 0 Then
        MsgBox "Client list is empty!", vbOKOnly + vbCritical, "No Clients loaded yet!"
    ElseIf i = -99 Then
        msg = "Can't load client list from Database!"
        GoTo EH
    End If
    Initialize = True
MEX:
    Exit Function
EH:
    MsgBox msg, vbOKOnly + vbCritical, "Can't start Application!"
    Unload Me
    GoTo MEX
    
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = vbKeyI Then
        mnuInboxName_Click
    End If
End Sub

Private Sub chkRun_Click()
    With Me.chkRun
        If .Value = 1 Then
            .Caption = "Stop"
            .ForeColor = vbRed
            Me.lblRun.Caption = "Server is running"
            Me.lblRun.ForeColor = vbGreen
            Me.Timer1.Enabled = True
            Set Me.Icon = Me.ImgList.ListImages(1).Picture
            ChangeSysTray True
        ElseIf .Value = 0 Then
            .Caption = "Start"
            .ForeColor = vbGreen
            Me.lblRun.Caption = "Server is stopped"
            Me.lblRun.ForeColor = vbRed
            Me.Timer1.Enabled = False
            Set Me.Icon = Me.ImgList.ListImages(2).Picture
            ChangeSysTray False

        End If
    End With
End Sub

Private Sub cmdClose_Click()
 End
End Sub

Private Sub cmdEditClient_Click()
Dim fCL As New frmClientList
    With fCL
        .Component oSFQ
        .Show vbModal
    End With
MEX:
    Unload fCL
    Set fCL = Nothing
Exit Sub
EH:
    MsgBox "Error editing client list!" & vbLf & "Error: " & Error
    GoTo MEX
End Sub

Private Sub cmdNewRecSet_Click()
'    oHR.CreateDisconnRecSet True
End Sub

Private Sub menClose_Click()
    Unload Me
End Sub

Private Sub menStop_Click()
    If Me.chkRun.Value = 1 Then
        Me.chkRun.Value = 0
        Me.menStop.Caption = "POS Server - Start"
    Else
        Me.chkRun.Value = 1
        Me.menStop.Caption = "POS Server - Stop"
    End If
End Sub



Private Sub mnuInboxName_Click()
    MsgBox oSFQ.SharedInboxName, vbOKOnly + vbInformation, "Shared InBox Name"
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
End Sub

Private Sub ChangeSysTray(IsRunning As Boolean)
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
End Sub
Private Sub UnloadSysTray()
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
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As _
                          Single, Y As Single)
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
End Sub

Private Sub Form_Resize()
    'this is necessary to assure that the minimized window is hidden
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub mPopRestore_Click()
Dim r As Long
    'called when the user clicks the popup menu Restore command
    Me.WindowState = vbNormal
    r = SetForegroundWindow(Me.hwnd)
    Me.Show
End Sub
