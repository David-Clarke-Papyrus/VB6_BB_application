VERSION 5.00
Object = "{C1740A22-225F-11D1-86A2-006097B34438}#1.0#0"; "MTrayX.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMainScheduler 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Papyrus scheduler"
   ClientHeight    =   2520
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5220
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   630
      Top             =   1185
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":015A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":06F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":0C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":0DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":0F42
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":109C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":11F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":1350
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":14AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":1604
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":175E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":18B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":1BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":1EEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":2206
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":2520
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":283A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainScheduler.frx":2B54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer objT 
      Enabled         =   0   'False
      Interval        =   65000
      Left            =   1335
      Top             =   1320
   End
   Begin MTRAYXLibCtl.TrayX TrayX1 
      Left            =   0
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      ToolTipText     =   "Papyrus console"
      Icon            =   "frmMainScheduler.frx":2E6E
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2145
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8644
            MinWidth        =   8644
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1005
      Left            =   75
      TabIndex        =   1
      Top             =   60
      Width           =   3225
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMainScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bStartscheduler As Boolean
Dim dteNominalNextUPdate As Date
Dim fWait As Boolean
Public m_LastUpdateDate As Date
Public m_NextNominalUpdate As Date
Public m_UpdateWindowStart As Date
Public m_UpdateWindowEnd As Date
Public m_DaysInWeek As Integer

Public Sub Component(pbStartscheduler As Boolean)
Dim rs As ADODB.Recordset
Dim strStatus As String

    Set rs = New ADODB.Recordset
    rs.Open "SELECT CF_LASTUPDATEDATE,CF_NEXTNOMINALUPDATE,CF_UPDATEWINDOWSTART,CF_UPDATEWINDOWEND,CF_DAYSINWEEK FROM tCONFIGURATION", cnPapy, adOpenKeyset, adLockOptimistic
    m_LastUpdateDate = rs.Fields("CF_LASTUPDATEDATE")
    m_NextNominalUpdate = rs.Fields("CF_NEXTNOMINALUPDATE")
    m_UpdateWindowStart = rs.Fields("CF_UPDATEWINDOWSTART")
    m_UpdateWindowEnd = rs.Fields("CF_UPDATEWINDOWEND")
    m_DaysInWeek = rs.Fields("CF_DAYSINWEEK")
    rs.Close
    Set rs = Nothing
    strStatus = "Next daily update for: " & Format(m_NextNominalUpdate, "dd/mm/yyyy")
    strStatus = strStatus & vbCrLf & "Will start after " & Format(m_UpdateWindowStart, "dd/mm/yyyy Hh:Nn")
    lblStatus.Caption = strStatus
    bStartscheduler = pbStartscheduler
    If bStartscheduler Then
        Me.SB1.Panels(1).Text = "Scheduler activated and waiting . . ."
    Else
        Me.SB1.Panels(1).Text = "Scheduler inactive"
    End If
    ResetSchedulerdates
    DisconnectDB
End Sub
Private Sub UpdateScreenData(pdteEffectiveDate)
    Me.SB1.Panels(1).Text = "Last day-end for trading day: " & Format(pdteEffectiveDate, "dd/mm/yyyy")
    Me.SB1.Refresh
    MousePointer = vbDefault
End Sub



Private Sub Form_Load()
        
    Me.SB1.Panels(1).Text = "Last day-end for trading day: " & Format(m_LastUpdateDate, "dd/mm/yyyy")
End Sub



Private Sub mnuExit_Click()
    DisconnectDB
    Unload Me
End Sub


Private Sub objT_Timer()
Dim lngResult As Long
Dim lngPosition As Long
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

    objT.Enabled = False
    If InWindow And DatePart("y", m_LastUpdateDate) < DatePart("y", m_NextNominalUpdate) Or DatePart("yyyy", m_LastUpdateDate) < DatePart("yyyy", m_NextNominalUpdate) Then
        lngResult = OpenDB()
        MsgBox "Update started: &        " & m_NextNominalUpdate & "     Scheduler"
        lblStatus.Caption = "Daily update running . . ." & vbCrLf & vbCrLf & "   Please wait!"
        cnPapy.Execute "BACKUP DATABASE PJ to disk = 'c:\PBKS\BU\PJ.BAK' WITH INIT"
        cnPapy.Execute "BACKUP LOG PJ to disk = 'c:\PBKS\BU\PJLOG.BAK' WITH INIT"
        Set cmd = New ADODB.Command
        Set cmd.ActiveConnection = oPC.CO
        cmd.CommandText = "sp_DAYEND"
        cmd.CommandType = adCmdStoredProc
        Set prm = New ADODB.Parameter
        prm.Type = adVarChar
        prm.Size = 10
        prm.Direction = adParamInput
        prm.Value = ReverseDate(pDate)
        cmd.Parameters.Append prm
        Set prm = Nothing
        Set prm = New ADODB.Parameter
        prm.Type = adInteger
        prm.Direction = adParamInput
        prm.Value = pContactID
        cmd.Parameters.Append prm
        Set prm = Nothing
        Set prm = New ADODB.Parameter
        prm.Type = adInteger
        prm.Direction = adParamOutput
        prm.Value = lngResult
        cmd.Parameters.Append prm
        Set prm = Nothing
        Set prm = New ADODB.Parameter
        prm.Type = adInteger
        prm.Direction = adParamOutput
        prm.Value = lngPosition
        cmd.Parameters.Append prm
        cmd.Execute
        If lngResult <> 0 Then
            GoTo ERRH
        End If
        DisconnectDB
        Me.SB1.Panels(1).Text = "Last day-end for trading day: " & Format(m_NextNominalUpdate, "dd/mm/yyyy")
        Me.Refresh
    Else
        If bStartscheduler Then
            SB1.Panels(1).Text = "Scheduler activated and waiting . . ."
        Else
            SB1.Panels(1).Text = "Scheduler inactive"
        End If
    End If
    ResetSchedulerdates
    objT.Enabled = True
ERRH:
    DisconnectDB
    Exit Sub
    Resume
End Sub
Private Sub Form_Unload(Cancel As Integer)
        If MsgBox("Stopping the dayend scheduler?", vbExclamation + vbYesNo + vbDefaultButton2, "Warning") = vbNo Then
            Cancel = True
        Else
            Me.TrayX1.IconVisible = False
        End If
End Sub

Private Sub objBatch_Status(pMsg As String)
    If pMsg = "<none>" Then
        Me.MousePointer = vbDefault
    Else
        Me.MousePointer = vbHourglass
    End If
    Me.SB1.Panels(2).Text = "Current action: " & pMsg
    Me.Refresh
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        If bStartscheduler Then objT.Enabled = True
        Me.Visible = False
    Else
        objT.Enabled = False
        Me.Height = 2000
        Me.Width = 3800
        Me.Visible = True
    End If
End Sub


Private Sub TrayX1_DblClick()
Dim Result As Long
       Me.WindowState = vbNormal
       Result = SetForegroundWindow(Me.hwnd)
       Me.Show
End Sub

Public Sub StartScheduler(pSTart As Boolean)
    bStartscheduler = pSTart
End Sub
Private Sub ResetSchedulerdates()
    Dim iPresenthour As Integer
Dim dteLastGoodUpdate As Date
    fWait = True
    dteLastGoodUpdate = m_LastUpdateDate
    iPresenthour = DatePart("h", Now())
    If iPresenthour >= 12 And iPresenthour < 24 _
    And DatePart("y", dteLastGoodUpdate) < DatePart("y", Now) _
    And WorkingDay(Now) = True Then
                dteNominalNextUPdate = Now
                fWait = False
    ElseIf iPresenthour >= 12 And iPresenthour < 24 _
    And DatePart("y", dteLastGoodUpdate) < DatePart("y", Now) _
    And WorkingDay(Now) = False Then
        dteNominalNextUPdate = GetPreviousWorkingDay(m_DaysInWeek, Now)
                fWait = False
    ElseIf iPresenthour >= 0 And iPresenthour < 12 _
    And DatePart("y", dteLastGoodUpdate) < DatePart("y", GetPreviousWorkingDay(m_DaysInWeek, Now)) _
    And WorkingDay(Now) = True Then
        dteNominalNextUPdate = GetPreviousWorkingDay(m_DaysInWeek, Now)
                fWait = False
    ElseIf iPresenthour >= 0 And iPresenthour < 12 _
    And DatePart("y", dteLastGoodUpdate) < DatePart("y", Now) _
    And WorkingDay(Now) = False Then
        dteNominalNextUPdate = GetPreviousWorkingDay(m_DaysInWeek, Now)
                fWait = False
    End If
    cnPapy.Execute "UPDATE tCONFIGURATION SET CF_NEXTNOMINALUPDATE = " & Format(dteNominalNextUPdate, "yyyy-mm-dd")
End Sub
Private Function GetActualUpdateTime(pNextUPdate) As Date
Dim iHoursToAdd As Integer
Dim iMinsToAdd As Integer

    iHoursToAdd = DatePart("h", m_UpdateWindowStart)
    iMinsToAdd = DatePart("n", m_UpdateWindowStart)
    If DatePart("h", m_UpdateWindowStart) < 12 Then
        iHoursToAdd = 12 + iHoursToAdd
    End If
    pNextUPdate = DateAdd("h", iHoursToAdd, CDate(DatePart("yyyy", pNextUPdate) & "-" & DatePart("m", pNextUPdate) & "-" & DatePart("d", pNextUPdate)))
    GetActualUpdateTime = DateAdd("n", iMinsToAdd, pNextUPdate)
End Function
Private Function GetNextWorkingDay(DIW As Integer, pLastDate As Date) As Date
    Select Case DIW
    Case 5
        If Weekday(pLastDate, vbMonday) = 5 Then
            GetNextWorkingDay = DateAdd("d", 3, pLastDate)
        Else
            GetNextWorkingDay = DateAdd("d", 1, pLastDate)
        End If
    Case 6
        If Weekday(pLastDate, vbMonday) = 6 Then
            GetNextWorkingDay = DateAdd("d", 2, pLastDate)
        Else
            GetNextWorkingDay = DateAdd("d", 1, pLastDate)
        End If
    Case 7
            GetNextWorkingDay = DateAdd("d", 1, pLastDate)
    End Select
End Function
Private Function GetPreviousWorkingDay(DIW As Integer, pLastDate As Date) As Date
    Select Case DIW
    Case 5
        If Weekday(pLastDate, vbMonday) = 1 Then
            GetPreviousWorkingDay = DateAdd("d", -3, pLastDate)
        ElseIf Weekday(pLastDate, vbMonday) = 7 Then
            GetPreviousWorkingDay = DateAdd("d", -2, pLastDate)
        Else
            GetPreviousWorkingDay = DateAdd("d", -1, pLastDate)
        End If
    Case 6
        If Weekday(pLastDate, vbMonday) = 1 Then
            GetPreviousWorkingDay = DateAdd("d", 2, pLastDate)
        Else
            GetPreviousWorkingDay = DateAdd("d", -1, pLastDate)
        End If
    Case 7
            GetPreviousWorkingDay = DateAdd("d", -1, pLastDate)
    End Select
End Function
Private Function WorkingDay(pDate As Date)
    If Weekday(pDate, vbMonday) >= 1 And Weekday(pDate, vbMonday) <= m_DaysInWeek Then
        WorkingDay = True
    Else
        WorkingDay = False
    End If
End Function

Function InWindow() As Boolean
Dim dteFrom As Date
Dim dteTo As Date

    If fWait = True Then
        InWindow = False
        Exit Function
    End If
    InWindow = False
    dteFrom = m_UpdateWindowStart
    dteTo = m_UpdateWindowEnd
    If dteFrom > dteTo Then  'the times are from night to morning
        If DatePart("h", Now()) * 60 + DatePart("n", Now()) > DatePart("h", dteFrom) * 60 + DatePart("n", dteFrom) _
        Or DatePart("h", Now()) * 60 + DatePart("n", Now()) < DatePart("h", dteTo) * 60 + DatePart("n", dteTo) Then
            InWindow = True
        End If
    Else
        If DatePart("h", Now()) * 60 + DatePart("n", Now()) > DatePart("h", dteFrom) * 60 + DatePart("n", dteFrom) _
        And DatePart("h", Now()) * 60 + DatePart("n", Now()) < DatePart("h", dteTo) * 60 + DatePart("n", dteTo) Then
            InWindow = True
        End If
    End If
End Function

