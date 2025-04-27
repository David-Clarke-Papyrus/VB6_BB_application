VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPosServerMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POS Server"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   4545
   ControlBox      =   0   'False
   FillColor       =   &H00C00000&
   Icon            =   "frmPOSSvrMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboTimerInterval 
      Height          =   315
      ItemData        =   "frmPOSSvrMain.frx":038A
      Left            =   1560
      List            =   "frmPOSSvrMain.frx":03B8
      TabIndex        =   15
      Text            =   "5"
      Top             =   2925
      Width           =   555
   End
   Begin VB.CommandButton cmdMinimize 
      BackColor       =   &H00C4BCA4&
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
      Height          =   465
      Left            =   3405
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3315
      Width           =   1080
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   1890
      TabIndex        =   1
      Top             =   4365
      Visible         =   0   'False
      Width           =   1965
      Begin VB.OptionButton optStop 
         BackColor       =   &H00C8B9B3&
         Caption         =   "Stop"
         Height          =   345
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   855
      End
      Begin VB.OptionButton optSTart 
         BackColor       =   &H00C8B9B3&
         Caption         =   "Start"
         Height          =   345
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label lblRun 
         Alignment       =   2  'Center
         BackColor       =   &H0080443E&
         BackStyle       =   0  'Transparent
         Caption         =   "Running"
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
         Height          =   285
         Left            =   -90
         TabIndex        =   4
         Top             =   600
         Width           =   2040
      End
   End
   Begin MSComctlLib.ListView lvwClients 
      Height          =   1320
      Left            =   90
      TabIndex        =   0
      Top             =   390
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   2328
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDropMode     =   1
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483635
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Station name"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Count"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Last exchange time"
         Object.Width           =   2011
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   -90
      Top             =   1710
   End
   Begin VB.Label lblInterval 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Polling interval"
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
      Left            =   225
      TabIndex        =   14
      Top             =   2940
      Width           =   1365
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   675
      Left            =   195
      TabIndex        =   13
      Top             =   3360
      Width           =   3075
      WordWrap        =   -1  'True
   End
   Begin VB.Shape sh2b 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   2580
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Shape sh1b 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   2190
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Polling database changes"
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
      Left            =   225
      TabIndex        =   12
      Top             =   2220
      Width           =   2475
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Polling P.O.S. workstations"
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
      Left            =   225
      TabIndex        =   11
      Top             =   2580
      Width           =   2475
   End
   Begin VB.Shape sh2 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2850
      Shape           =   3  'Circle
      Top             =   2580
      Width           =   285
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Exchanges from sales terminals"
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
      Left            =   135
      TabIndex        =   10
      Top             =   120
      Width           =   2790
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
      Left            =   2655
      TabIndex        =   9
      Top             =   -30
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Exchanges"
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
      Left            =   570
      TabIndex        =   8
      Top             =   -30
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Updates from server"
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
      Left            =   570
      TabIndex        =   7
      Top             =   1800
      Width           =   1725
   End
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
      Left            =   2355
      TabIndex        =   6
      Top             =   1800
      Width           =   615
   End
   Begin VB.Shape sh1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2850
      Shape           =   3  'Circle
      Top             =   2190
      Width           =   285
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClose 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOPtions 
      Caption         =   "&Options"
      Begin VB.Menu mnuUtilities 
         Caption         =   "&Utilities"
      End
      Begin VB.Menu mnuAudit 
         Caption         =   "&Audit"
      End
      Begin VB.Menu mnuRecover 
         Caption         =   "Recover missing exchanges"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGetExchanges 
         Caption         =   "Fetch missing exhange by number"
      End
      Begin VB.Menu mnuSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerifyPOSinventory 
         Caption         =   "Verify inventory"
      End
      Begin VB.Menu mnuSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "Debug"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSeo03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSB 
         Caption         =   "View service broker log"
      End
      Begin VB.Menu mnuViewSQLServerLog 
         Caption         =   "View SQL server error log"
      End
      Begin VB.Menu mnuViewPapyrusErrors 
         Caption         =   "View Papyrus II error log"
      End
   End
End
Attribute VB_Name = "frmPosServerMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bSysTrayLoaded As Boolean
Private nid As NOTIFYICONDATA
Public WithEvents oPSvs As Z_PollingServices
Attribute oPSvs.VB_VarHelpID = -1
Dim bCloseApp As Boolean

Private Sub cboTimerInterval_Click()
    On Error GoTo errHandler
    PollingInterval = CLng(cboTimerInterval) * 1000
    Timer1.Interval = PollingInterval
    SaveSetting "PBKS", "POSSvr", "PollingInterval", CStr(PollingInterval)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.cboTimerInterval_Click"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo errHandler
Dim CloseMode As String

    Select Case UnloadMode
    Case 0
        CloseMode = "User has closed application from the Control box"
    Case 1
        CloseMode = "Application has closed under program control"
    Case 2
        CloseMode = "Application is closing because Windows is closing"
    Case 3
        CloseMode = "Application has closed from task manager"
    Case 4
        CloseMode = "Application has closed because MDI owner form is closing"
    Case 5
        CloseMode = "Application has closed because owner form is closing"
    End Select
    LogSaveToFile "POSSvr is closing: in QueryUnload. Unload mode = " & CloseMode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.Form_QueryUnload(Cancel,UnloadMode)", Array(Cancel, UnloadMode)
End Sub

Private Sub mnuDebug_Click()
    mnuDebug.Checked = Not mnuDebug.Checked
    bDebug = mnuDebug.Checked
End Sub

Private Sub mnuGetExchanges_Click()
    On Error GoTo errHandler
Dim frm As New frmRecoverExchangesMissing
Dim i As Integer
Dim msg As String

    frm.Show vbModal
    i = 0
    Do
        msg = frm.GetMsg(i)
        i = i + 1
        If msg > "" Then oPSvs.ReprocessMessage msg, ""
    Loop Until msg = ""
    Unload frm
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.mnuGetExchanges_Click"
End Sub

Private Sub mnuRecover_Click()
    On Error GoTo errHandler
Dim frm As frmFindExchangeFiles

    Set frm = New frmFindExchangeFiles
    frm.Show vbModal
    If Not frm.Cancelled Then
        ReprocessPOSMsg frm.Filenames
    End If
    
    Unload frm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.mnuRecover_Click"
End Sub
Private Sub GetExchangeFromFD()

End Sub
Private Sub ReprocessPOSMsg(strFilenames As String)
    On Error GoTo errHandler
Dim i As Integer
Dim oTF As New z_TextFile
Dim strMsg As String
Dim ar() As String
Dim fs As New FileSystemObject
Dim strPath As String

    If strFilenames = "" Then Exit Sub
    ar = Split(strFilenames, Chr(0))
    If UBound(ar) = 0 Then 'There is only one file
        strPath = fs.GetParentFolderName(ar(0))
        oTF.OpenTextFileToRead ar(0)
        strMsg = oTF.ReadWholeFile
        oTF.CloseTextFile
        oPSvs.ReprocessMessage strMsg, ar(0)
    Else
        strPath = ar(0)
        For i = 1 To UBound(ar)
            oTF.OpenTextFileToRead strPath & "\" & ar(i)
            strMsg = oTF.ReadWholeFile
            oTF.CloseTextFile
            oPSvs.ReprocessMessage strMsg, strPath & "\" & ar(i)
        Next i
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.ReprocessPOSMsg(strFilenames)", strFilenames
End Sub

Private Sub mnuVerifyPOSinventory_Click()
    On Error GoTo errHandler
Dim F As New frmVerifyStock
    F.Show
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.mnuVerifyPOSinventory_Click"
End Sub


Private Sub oPSvs_DBStatus(val As Boolean)
    sh2.Visible = Not sh2.Visible
    sh2b.Visible = Not sh2b.Visible
End Sub

Private Sub oPSvs_DBStatusMsg(val As String)
    Me.lblStatus.Caption = val
End Sub

Private Sub oPSvs_QPOSTriggerStatus(pOn As Boolean)
    On Error GoTo errHandler
    Timer1.Enabled = pOn
    sh2.Visible = Not sh2.Visible
    sh2b.Visible = Not sh2b.Visible
    sh2.Refresh
    sh2b.Refresh

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.oPSvs_QPOSTriggerStatus(pOn)", pOn
End Sub
Private Sub oPSvs_POSMessageReceived(strStationName As String)
    On Error GoTo errHandler
Dim itm As ListItem

    On Error Resume Next
    Set itm = lvwClients.ListItems(UCase(strStationName))
    If Err = 35601 Then
        MsgBox "A P.O.S. workstation is labelled '" & strStationName & "', but this is not known on this application. " & vbCrLf & "Check the workstations set for this server by using menu : Options>Utilities.", vbOKOnly, "Problem"
    Else
        itm.SubItems(1) = FNN(itm.SubItems(1)) + 1
        itm.SubItems(2) = CStr(Format(Now(), "dd HH:NN am/pm"))
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.oPSvs_POSMessageReceived(strStationName)", strStationName, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdMinimize_Click()
    On Error GoTo errHandler
    Me.InitSysTray
    Me.WindowState = vbMinimized
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdMinimize_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwClients_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub mnuAudit_Click()
Dim frm As New frmMissingExchanges
    frm.Show
End Sub

Private Sub mnuClose_Click()
    If MsgBox("You should not exit this application unless instructed. " & vbCrLf & "You should rather minimize it using the button provided." & vbCrLf & "Do you want to close?", vbYesNo, "Warning") = vbYes Then
        bCloseApp = True
        Screen.MousePointer = vbHourglass
'        If Timer1.Enabled Then
'            cmdclose_Click
'        Else
'            If MsgBox("The application may be updating the database, you will see numbers changing if this is so, if so it is best not to close it now. " & vbCrLf & "Do you still want to close the application?", vbQuestion + vbOKCancel, "Warning") <> vbCancel Then
'                cmdclose_Click
'            End If
'        End If
    End If
End Sub

Private Sub mnuUtilities_Click()
    cmdUtilities_Click
End Sub

Private Sub oPSvs_ClientStatusChange()
    LoadClientListview
End Sub
Private Sub cmdUtilities_Click()
    On Error GoTo errHandler
Dim frm As New frmPOSSVRUtilities
    Timer1.Enabled = False
    frm.SetPollingServicesRef = oPSvs
    frm.Show vbModal
    oMS.LoadarClientList
   ' LoadClientListview

    Timer1.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.cmdUtilities_Click"
End Sub
'Private Sub oPSvs_SwitchPolling(bOn As Boolean)
'    On Error GoTo errHandler
'    If bOn Then
'       optSTart = True
'    Else
'        optStop = True
'    End If
'    Me.Refresh
'    DoEvents
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPosServerMain.oPSvs_SwitchPolling(bOn)", bOn
'End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
Dim i As Integer
Dim strPos As String

    
    bCloseApp = False
    bDebug = IIf(GetSetting("PBKS", "Debug", "POSSVR", "FALSE") = "TRUE", True, False)
    mnuDebug.Checked = bDebug

    Set oMS = New z_ManageStations
    Set oPSvs = New Z_PollingServices
    Set oMF = New Z_ManageFolders
    oMS.LoadarClientList
    LoadClientListview
    If oPC.CustomerTPID = 1 Then
        MsgBox "There is no cash sale account recognized by the system." & vbCrLf & "Use the main Papyrus application, ensure you have a customer with the account code CASH01," & vbCrLf & " go to settings/configuration, go to the main tab and click 'Locate cash sales customer'." & vbCrLf & "Restart this application", vbOKOnly, "Can't continue"
        Exit Sub
    End If
    If UBound(arCOMMANDLINE) > 0 Then
        If arCOMMANDLINE(1) <> "N" Then
            Me.BackColor = vbRed
        End If
        If UBound(arCOMMANDLINE) > 1 Then
            If arCOMMANDLINE(2) <> "" Then
                PollingInterval = CLng(arCOMMANDLINE(2))
            End If
        End If
    End If

    Me.KeyPreview = True
    Timer1.Interval = GetSetting("PBKS", "POSSvr", "PollingInterval", "5000")

    Timer1.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSServer.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set oMS = Nothing
    Set oPSvs = Nothing
    Set oMF = Nothing
    SaveSetting "PBKS", "Debug", "POSSVR", IIf(bDebug, "TRUE", "FALSE")
    oPC.DisconnectDBShort
        
    Set oPC = Nothing
    'Delete Icon from SysTray
    If bSysTrayLoaded Then
        UnloadSysTray
        Timer1.Enabled = False
        DoEvents
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub LoadClientListview()
    On Error GoTo errHandler
Dim i As Integer
Dim lstItem As ListItem
    lvwClients.ListItems.Clear
    For i = 0 To oMS.arClientListCount
        Set lstItem = lvwClients.ListItems.Add(, UCase(oMS.ClientName(i)))
        With lstItem
            .Text = oMS.ClientName(i)
        End With
    Next

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.LoadClientListview"
End Sub

Private Sub oPSvs_DataSent()
    lblUpdates = val(lblUpdates) + 1
End Sub
Private Sub Timer1_Timer()
On Error GoTo errHandler
Dim tim As Date

    If bCloseApp Then
        Timer1.Enabled = False
        Screen.MousePointer = vbDefault
        Unload Me
        Exit Sub
    Else
        Timer1.Enabled = False
        sh1.Visible = Not sh1.Visible
        sh1b.Visible = Not sh1b.Visible
        sh1.Refresh
        sh1b.Refresh
        oPSvs.SetQPOSTrigger
    End If
    DoEvents
START:
    If oPSvs Is Nothing Then
        Exit Sub
    End If
    
    tim = Now()
    If Hour(tim) < 22 And Hour(tim) > 7 Then  'we don't want to do this during the period at night when batcher operations take place
        If oPSvs.PutMasterDataInOutbox = False Then
    
        End If
    End If
    Timer1.Enabled = True

    Exit Sub

errHandler:
   ' If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.Timer1_Timer", , EA_NORERAISE
    HandleError
  '  Unload Me
End Sub


Private Sub optStart_Click()
    On Error GoTo errHandler
    If optSTart = True Then
      Me.lblRun.Caption = "Running"
      Me.Timer1.Enabled = True
      oPSvs.OpenQPOS
      ChangeSysTray True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.optStart_Click"
End Sub
Private Sub optStop_Click()
    On Error GoTo errHandler
    If optStop = True Then
       Me.lblRun.Caption = "Stopped"
       Me.Timer1.Enabled = False
       oPSvs.CloseQPOS
       ChangeSysTray False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.optStop_Click"
End Sub

Private Sub cmdclose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub menClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.menClose_Click", , EA_NORERAISE
    HandleError
End Sub


'Stuff for System Tray functionality~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Sub InitSysTray()
    On Error GoTo errHandler
Dim Res As Boolean
Dim cnt As Integer
    'the form must be fully visible before calling Shell_NotifyIcon
    SysTrayText = "P.O.S. Server: Running" & vbNullChar
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
    DoEvents
'    WindowState = vbMinimized
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.InitSysTray"
End Sub

Private Sub ChangeSysTray(IsRunning As Boolean)
    On Error GoTo errHandler
Dim Res As Boolean
Dim cnt As Integer
    If IsRunning Then
        SysTrayText = "P.O.S. Server: Running" & vbNullChar
    Else
        SysTrayText = "P.O.S. Server: Stopped" & vbNullChar
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
        Shell_NotifyIcon NIM_DELETE, nid
        Res = Shell_NotifyIcon(NIM_MODIFY, nid)
        If Res = False Then
            EventPause 1
            cnt = cnt + 1
        End If
    Loop
    DoEvents
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.ChangeSysTray(IsRunning)", IsRunning
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
    DoEvents
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosServerMain.UnloadSysTray"
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As _
                          Single, Y As Single)
    On Error GoTo errHandler
    'this procedure receives the callbacks from the System Tray icon.
Dim Result As Long
Dim msg As Long

    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        msg = x
    Else
        msg = x / Screen.TwipsPerPixelX
    End If
    Select Case msg
       ' Case WM_LBUTTONUP        '514 restore form window
           ' Me.WindowState = vbNormal
           ' Result = SetForegroundWindow(Me.hwnd)
           ' On Error Resume Next
           ' Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
            Me.WindowState = vbNormal
            Me.Show
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
           ' Refresh
           ' DoEvents
'        Case WM_RBUTTONUP        '517 display popup menu
'            Result = SetForegroundWindow(Me.hwnd)
'            Me.PopupMenu Me.menPopup
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.Form_MouseMove(Button,Shift,X,Y)", Array(Button, Shift, x, Y), _
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
    Me.Show
    Me.WindowState = vbNormal
    r = SetForegroundWindow(Me.hwnd)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPosListener.mPopRestore_Click", , EA_NORERAISE
    HandleError
End Sub
