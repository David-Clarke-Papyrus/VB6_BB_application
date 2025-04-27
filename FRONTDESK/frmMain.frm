VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00552619&
   Caption         =   "Papyrus II - Front desk"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   2385
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6000
      Top             =   4815
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8400
      Top             =   885
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Height          =   300
      Left            =   7380
      ScaleHeight     =   240
      ScaleWidth      =   285
      TabIndex        =   13
      Top             =   750
      Width           =   345
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00552619&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   405
      Left            =   5640
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   225
      Width           =   2280
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4470
      Top             =   4320
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   3360
      TabIndex        =   0
      Top             =   1155
      Width           =   2280
   End
   Begin VB.CommandButton cmdIssue 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Issue T/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8190
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Issues the current Transaction and starts a new transaction"
      Top             =   4230
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6000
      Top             =   4290
   End
   Begin MSComctlLib.ListView lstCSL 
      Height          =   2385
      Left            =   270
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1830
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4207
      SortKey         =   4
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483635
      BackColor       =   15790320
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "EAN/Code"
         Object.Width           =   2823
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   8714
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Price"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Time"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   6945
      Visible         =   0   'False
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12832
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "15/04/2009"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "20:07"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   6750
      Top             =   4230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00552619&
      Caption         =   "EAN, ISBN or code (manual entry)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   285
      TabIndex        =   11
      Top             =   1230
      Width           =   3015
   End
   Begin VB.Label lblSales 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   2790
      TabIndex        =   10
      Top             =   4710
      Width           =   2310
   End
   Begin VB.Label lblQty 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   2790
      TabIndex        =   9
      Top             =   4350
      Width           =   2310
   End
   Begin VB.Label lblStartedat 
      BackColor       =   &H00552619&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1635
      TabIndex        =   8
      Top             =   570
      Width           =   3330
   End
   Begin VB.Label lblCode 
      BackColor       =   &H00552619&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   240
      TabIndex        =   7
      Top             =   135
      Width           =   3330
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00552619&
      Caption         =   "Value of sales:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1170
      TabIndex        =   6
      Top             =   4710
      Width           =   1470
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00552619&
      Caption         =   "Number of books sold:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   435
      TabIndex        =   5
      Top             =   4350
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackColor       =   &H00552619&
      Caption         =   "Started at:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   540
      Left            =   255
      TabIndex        =   4
      Top             =   555
      Width           =   1245
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
      End
   End
   Begin VB.Menu mnuOp 
      Caption         =   "Operations"
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete selected row"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nRet         As Long
Private nMainhWnd    As Long
'Dim objPapyConn As PapyConn
Dim iFilenum As Integer
Dim strTextFilePath As String
Dim strTextFileName As String
Dim objCS As a_CS
Dim fClose As Boolean
Dim flgLoading As Boolean
Dim fIssued As Boolean
Dim fOffline As Boolean
Dim fExistingTA As Boolean
Dim objConfig As a_Configuration
Dim objError As a_Error
Dim strComputerName As String
Dim dblTotalValue As Double
Dim lngTotalQty As Long
Dim frmStartup As frmStartup
Dim objCSL As a_CSL
Dim objCSLs As c_CSL
Dim objTf As z_Logging
Attribute objTf.VB_VarHelpID = -1
Dim flgStatusVisible As Boolean
Dim objProduct As a_Product
Dim objProdcode As New z_ProdCode
Dim oQuickProduct As New z_QuickProduct
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Public Function CurrentCashSale() As a_CS
    On Error GoTo errHandler
    Set CurrentCashSale = objCS
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.CurrentCashSale"
End Function

Private Sub lstCSL_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.lstCSL_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub mnuDelete_Click()
    On Error GoTo errHandler
Dim lResult As Long
Dim strCode As String
Dim lID As Long
    If Me.lstCSL.SelectedItem.Index = -1 Then GoTo EXIT_Handler
    strCode = lstCSL.ListItems(lstCSL.SelectedItem.Key).Text
    oQuickProduct.DeleteRow val(lstCSL.SelectedItem.Key)
    lstCSL.ListItems.Remove (lstCSL.SelectedItem.Key)
   If objTf.WriteToDELLog(strCode) = False Then
    Beep
    MsgBox "Cannot write to Log file - call supervisor!", vbCritical, "Status"
End If

EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuDelete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub objTF_Status(pMsg As String)
    On Error GoTo errHandler
    If pMsg > "" Then
        MsgBox pMsg
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.objTF_Status(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub
Private Function NameOfPC(MachineName As String) As Long
    On Error GoTo errHandler


    Dim NameSize As Long
    Dim x As Long

    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
    MachineName = Left(MachineName, NameSize)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NameOfPC(MachineName)", MachineName
End Function
  
  
  


Private Sub cmdIssue_Click()
    On Error GoTo errHandler
Dim lngResult As Long
Dim OpenResult As Integer

    If MsgBox("Do you want to issue this transaction?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    objCS.Status = 3
    objCS.DateIssued = Date
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.Execute "UPDATE tTR SET TR_Status = 3,TR_Date= '" & ReverseDate(Date) & "' WHERE TR_ID = " & objCS.TRID
'-------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    fIssued = True
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdIssue_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdStart_Click()
    On Error GoTo errHandler
Dim strFname As String
Dim lngResult As Long

    objCS.DateStarted = Now()
    objCS.FrontDeskComputerName = strComputerName
    SaveSetting App.EXEName, "LOGFILE", "NAME", objTf.LogFileName
    objCS.TPID = objConfig
    objCS.ApplyEdit
    LoadControls
  '  DisableCOntrols
    Timer1.Enabled = True
    objCS.BeginEdit

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdStart_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub START()
    On Error GoTo errHandler
Dim strFname As String
Dim lngResult As Long
    objCS.DateStarted = Now()
    objCS.DateIssued = Date
    objCS.FrontDeskComputerName = strComputerName
    SaveSetting App.EXEName, "LOGFILE", "NAME", objTf.LogFileName
    SaveSetting App.EXEName, "DELLOGFILE", "NAME", objTf.LogDELFileName
    objCS.TPID = oPC.Configuration.CSCustomerID
    objCS.SetTILLID strComputerName
    objCS.ApplyEdit
    LoadControls
    Timer1.Enabled = True

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.START"
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
Dim objErr As New a_Error
Dim Retval
Dim strDBLocation As String
Dim lngResult As Long
Dim sFile As String
Dim iCom As Integer
    GetThunder
    Set objProduct = New a_Product
    Set objProdcode = New z_ProdCode
    SetLvw
    flgLoading = True
    Retval = NameOfPC(strComputerName)
    Screen.MousePointer = vbHourglass
    Set objTf = New z_Logging
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    strCom = GetSetting(App.Title, "Settings", "COMPort", "COM2:")
    
 '   Dim Result As Integer
    If Not fClose Then
        Timer1.Enabled = False
        Me.Timer1.Interval = oPC.TimerInterval
        Timer1.Enabled = True
        
        bSendsCRLF = oPC.ScannerSendsCR
        MSComm1.Settings = oPC.COMPORTSettings
        MSComm1.CommPort = oPC.COMMPORTNumber
        If MSComm1.PortOpen = False Then
           MSComm1.PortOpen = True
        End If

        If fOffline Then
            objTf.OpenOffLineLog
            Me.sbStatusBar.Panels(1).Text = "Appending to . . . " & objTf.OfflineLogName
     '       DisableCOntrols
            Me.cmdIssue.Enabled = False
            Timer1.Enabled = True
            Me.txtStatus = "Offline"
        Else
            Set objCS = New a_CS
            objCS.SetTILLID strComputerName
            objCS.LoadExistingTA lngResult, Date, strComputerName
            If lngResult = 0 Then 'An un-issued TA is found
                If Format(objCS.CaptureDate, "yymmdd") <> Format(Date, "yymmdd") Then  'It is not for today
                    If MsgBox("There is an un-issued transaction dated: " & Format(objCS.DateStarted, "dd/mm/yyyy") & "." & Chr(10) & Chr(13) _
                    & "Do you want to issue that one and start a new transaction (recommended).?", vbQuestion + vbYesNo, "Status") = vbNo Then
                    'Append to existing TA (not started today
                        LoadControls
                        LoadList
                        objTf.OpenExistingLog GetSetting(App.EXEName, "LOGFILE", "NAME", "")
                        fExistingTA = True
                        Timer1.Enabled = True
                        objCS.BeginEdit
                    Else
                        objCS.BeginEdit
                        objCS.Status = 3
                        objCS.ApplyEdit
                        fExistingTA = False
                        objTf.OpenNewLog
                        Set objCS = New a_CS
                        objCS.BeginEdit
                        START
                    End If
                Else 'existing TA and it is for today
                    LoadControls
                    LoadList
                    objTf.OpenExistingLog GetSetting(App.EXEName, "LOGFILE", "NAME", "")
                    fExistingTA = True
                    Timer1.Enabled = True
                    objCS.BeginEdit
                End If
            Else
                'Start new transaction
                fExistingTA = False
                objTf.OpenNewLog
                objCS.BeginEdit
                START
            End If
            Me.sbStatusBar.Panels(1).Text = "Appending to text file:" & objTf.LogFileName
        End If
        
    Else
        Unload Me
        GoTo EXIT_Handler
    End If
    Me.lblQty = lngTotalQty
    Me.lblSales = Format(dblTotalValue / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
   ' Me.lblSales = Format(dblTotalValue, "R#,##0.00")
    flgLoading = False
    Screen.MousePointer = vbDefault
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
    Me.lblCode.Caption = objCS.DOCCode
    Me.lblStartedat.Caption = Format(objCS.DateStarted, "ddd, dd/mm/yyyy  hh:mm ampm")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.LoadControls"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Dim i As Integer
    Dim strMessage As String
    If Not fIssued Then
        strMessage = "Confirm Quit without issuing? (Y/N)"
        If MsgBox(strMessage, vbQuestion + vbYesNo, "Quitting") = vbNo Then
            Cancel = True
            GoTo EXIT_Handler
        End If
    End If
    If objCS.IsEditing Then objCS.ApplyEdit
    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    Close iFilenum
    If MSComm1.PortOpen = True Then
       MSComm1.PortOpen = False
    End If
    Set objConfig = Nothing
    Set objCS = Nothing
    Set objCSL = Nothing
    Set objProduct = Nothing
    Set objProdcode = Nothing
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub mnuErrors_Click()
    On Error GoTo errHandler
Dim frmErrorList As New frmErrorList
    frmErrorList.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuErrors_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuFileExit_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuFileExit_Click", , EA_NORERAISE
    HandleError
End Sub


'Private Sub mnuViewOptions_Click()
'    On Error GoTo errHandler
'Dim frmSerial As frmSerial
'Dim Result
'    Set frmSerial = New frmSerial
'    frmSerial.Show vbModal
'    Result = IO1.Close
'    Result = IO1.Open(strCom, "baud=9600 parity=N data=7 stop=1")  'Set up scanner
'    strCom = GetSetting(App.Title, "Settings", "COMPort", "COM2:")
'    iTimer1Interval = GetSetting(App.Title, "Settings", "Timer1Interval", "250")
'    Timer1.Enabled = False
'    Me.Timer1.Interval = iTimer1Interval
'    Timer1.Enabled = True
'    bSendsCRLF = GetSetting(App.Title, "Settings", "CRLF", "COM2:") = "Y"
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuViewOptions_Click", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub mnuViewStatusBar_Click()
    On Error GoTo errHandler
    flgStatusVisible = Not flgStatusVisible
    Me.sbStatusBar.Visible = flgStatusVisible
    Me.mnuViewStatusBar.Checked = flgStatusVisible
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuViewStatusBar_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Timer2_Timer()
    On Error GoTo errHandler
    fOffline = Not oPC.Connected 'set fOffline to true if false isreturned
    If fOffline = True Then
        objTf.OpenOffLineLog
        Me.sbStatusBar.Panels(1).Text = "Appending to . . . " & objTf.OfflineLogName
'        DisableCOntrols
        Me.cmdIssue.Enabled = False
        Me.txtStatus = "Offline"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Timer2_Timer", , EA_NORERAISE
    HandleError
End Sub

Private Sub Timer1_Timer()
Dim iMax As Integer
Dim strIn As String
Dim Started As Boolean

    On Error GoTo errHandler
'    If bSendsCRLF Then
'        iMax = 14
'    Else
'        iMax = 12
'    End If
    
    Timer1.Enabled = False
    Started = False
    Do While MSComm1.InBufferCount > 0
        Started = True
        DoEvents
        strIn = strIn & MSComm1.Input
        Timer3.Interval = 100
        Timer3.Enabled = True
        Do While Timer3.Enabled = True
            DoEvents
        Loop
    Loop
    If strIn > "" Then
        HandleInput strIn
    End If
    Timer1.Enabled = True
   
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Timer1_Timer", , EA_NORERAISE
    HandleError
End Sub
Private Sub Timer3_Timer()
   Timer3.Enabled = False
End Sub

Private Function ConvertString(pIn As String) As String
Dim i As Integer
Dim c As String
Dim out As String

For i = 1 To Len(pIn)
    c = Asc(Mid(pIn, i, 1))
    out = out & c & " "
Next i
    ConvertString = out
End Function
Private Sub HandleInput(pString As String)
    On Error GoTo errHandler
Dim iResult As Integer
Dim iFound As Long
Dim strPID As String
Dim iNumFound As Integer
Dim liListItem As ListItem
Dim lngResult As Long
Dim strCOdeF As String
Dim strPriceF As String

Dim i
Dim j
Dim strEAN As String, strCode As String, strTitle As String, lngPrice As Long
Dim lCSLID As Long
'Dim objProdcode As z_ProdCode
Dim tmpCSL As a_CSL

'    MsgBox "in HandleInput 1:   " & pString

    If DebugMode Then
        sbStatusBar.Panels(1).Text = pString & "    " & ConvertString(pString)
    End If
    
'    MsgBox "in HandleInput 2:   " & pString
    
    If Right(pString, 2) = vbCrLf Then
        pString = Left(pString, Len(pString) - 2)
    End If
 '   MsgBox "in HandleInput 3:   " & pString
    
START:
'    MsgBox "ISISBN10  " & IsISBN10(txtCode)
'    MsgBox "IsHashCode   " & IsHashCode(txtCode)
'    MsgBox "IsPrivateCode   " & IsPrivateCode(txtCode)
 '   If Not (IsISBN13(txtCode) Or IsISBN10(txtCode) Or IsHashCode(txtCode) Or IsPrivateCode(txtCode)) Then
    strCode = stripCRLF(pString)
    
 '   MsgBox "in HandleInput 4:   " & strCode
    
 '   MsgBox "Length of Code:" & Len(strCode) & "    " & strCode
   ' MsgBox "ISISBN13  " & IsISBN13(strCode)
    If Not (IsISBN13(strCode) Or IsISBN10(strCode) Or IsHashCode(strCode) Or IsPrivateCode(strCode)) Then
        MsgBox "This is an invalid code, retry.", vbInformation, "Warning"
        GoTo EXIT_Handler
    End If
'    MsgBox "After Validation"

    If fOffline Then
        If objTf.WriteToOffLineLog(strCode) = False Then
            Beep
            MsgBox "Cannot write to Log file - call supervisor!", vbCritical, "Status"
        End If
        Set liListItem = lstCSL.ListItems.Add
        With liListItem
            .Text = strCode
            .SubItems(4) = Now()
            lngTotalQty = lngTotalQty + 1
        End With
        lstCSL.SelectedItem = lstCSL.ListItems(1)
        lstCSL.Refresh
        Me.lblQty = lngTotalQty
    Else   ' Normal Situation
  '  MsgBox "Before Lookup"
        
            iResult = oQuickProduct.HandleCode(strPID, objCS.TRID, Trim(strCode), strTitle, lngPrice, lCSLID, strCOdeF, strPriceF)
            If iResult <> 0 Then  'not found anywhere
                
                Dim frmAdHoc As frmAdHocProduct
                Set frmAdHoc = New frmAdHocProduct
                frmAdHoc.Component pString
                frmAdHoc.Show vbModal
                pString = frmAdHoc.code
                Unload frmAdHoc
                Set frmAdHoc = Nothing
                GoTo START
                
            End If
   ' MsgBox "After Lookup"

        Set liListItem = lstCSL.ListItems.Add(Index:=1, Key:=lCSLID & "k")
        With liListItem
          '  If strCode = "" Then
                .Text = strCOdeF
          '  Else
          '      .Text = strCOdeF
          '  End If
            .SubItems(1) = strTitle
           ' .SubItems(2) = Format(objProduct.SPF / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
            .SubItems(2) = strPriceF
            .SubItems(3) = Format(Now(), "ddd, h:mm::ss A/P")
            .SubItems(4) = Format(Now(), "yyyy-mm-dd hh:mm:ss")
            dblTotalValue = dblTotalValue + lngPrice
            lngTotalQty = lngTotalQty + 1
        End With
        lstCSL.SelectedItem = lstCSL.ListItems(1)
        lstCSL.SortKey = 4
        Me.lblQty = lngTotalQty
        Me.lblSales = Format(dblTotalValue / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
        If objTf.WriteToLog(pString) = False Then
            Beep
            MsgBox "Cannot write to Log file - call supervisor!", vbCritical, "Status"
        End If
    End If
    Set tmpCSL = Nothing
    Set objProduct = Nothing
EXIT_Handler:
'ERR_Handler:
'    MsgBox "Error getting scanner code:" & pString
'    oError.SetError Err, Error, Time(), "frmMain", "HandleInput", App.EXEName
'    Exit Sub
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.HandleInput(pString)", pString
    Exit Sub
    Resume
End Sub

Private Sub LoadList()
    On Error GoTo errHandler
Dim objItem As a_CSL
Dim itmList As ListItem
Dim lngIndex As Long
    dblTotalValue = 0
    lngTotalQty = 0

    lstCSL.ListItems.Clear
    If objCS.CSLines.Count = 0 Then GoTo EXIT_Handler
    For lngIndex = 1 To objCS.CSLines.Count
        With objItem
            Set objItem = objCS.CSLines.Item(lngIndex)
            Set itmList = lstCSL.ListItems.Add(Key:=objItem.Key)
            With itmList
                .Text = objItem.CodeF
                .SubItems(1) = objItem.Title
                .SubItems(2) = objItem.PriceF
               ' .SubItems(3) = Format(objItem.DateTime, "ddd, dd mmm hh:mm AMPM")
                .SubItems(3) = objItem.DateTime
         '       .SubItems(4) = CDbl(CDate(objItem.DateTime))
                .SubItems(4) = objItem.DateTimeForSort
                dblTotalValue = dblTotalValue + objItem.Price
                lngTotalQty = lngTotalQty + 1

            End With
        End With
    Next
    lstCSL.SelectedItem = lstCSL.ListItems(1)
    Me.lblQty = lngTotalQty
  '  Me.lblSales = Format(dblTotalValue, "currency")
    Me.lblSales = Format(dblTotalValue / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
EXIT_Handler:
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.LoadList"
End Sub
Private Sub cboSP_Click()
    On Error GoTo errHandler
'    If flgLoading Then Exit Sub
'    If Not cboSP.ListIndex = -1 Then
'        objCS.OperatorID = objCS.operators.Key(cboSP)
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cboSP_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = vbKeyReturn Then
        HandleInput (CStr(Me.txtCode))
        Me.txtCode = ""
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.txtCode_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub SetLvw()
    On Error GoTo errHandler
Dim style As Long
Dim hHeader As Long
   
  'get the handle to the listview header
   hHeader = SendMessage(lstCSL.hwnd, LVM_GETHEADER, 0, ByVal 0&)
   
  'get the current style attributes for the header
   style = GetWindowLong(hHeader, GWL_STYLE)
   
  'modify the style by toggling the HDS_BUTTONS style
   style = style Xor HDS_BUTTONS
   
  'set the new style and redraw the listview
   If style Then
      Call SetWindowLong(hHeader, GWL_STYLE, style)
      Call SetWindowPos(lstCSL.hwnd, Me.hwnd, 0, 0, 0, 0, SWP_FLAGS)
   End If


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.SetLvw"
End Sub
''''''''''''''''''''''''''''''''''

Private Sub GetThunder()
    On Error GoTo errHandler
Dim hIcon As Long
    
    nRet = GetWindowLong(Me.hwnd, GWL_HWNDPARENT)
    Do While nRet
       nMainhWnd = nRet
       nRet = GetWindowLong(nMainhWnd, GWL_HWNDPARENT)
    Loop
    ' set the icon
    Set Me.Icon = Picture1.Picture
    ' get a handle to ICON_BIG
    hIcon = SendMessage(Me.hwnd, WM_GETICON, ICON_BIG, ByVal 0)
    ' send ICON_BIG to the main window
    SendMessage nMainhWnd, WM_SETICON, ICON_BIG, ByVal hIcon

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.GetThunder"
End Sub
''''''''''''''''''''''''''''''''''''''

