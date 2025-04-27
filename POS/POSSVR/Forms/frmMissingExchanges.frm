VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMissingExchanges 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Missing exchanges"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00D5C5C1&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   10740
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMissingExchanges.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Close the purchase order"
      Top             =   90
      Width           =   885
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   6525
      Left            =   135
      TabIndex        =   2
      Top             =   930
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   11509
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483638
      ForeColor       =   8404992
      TabCaption(0)   =   "Exchanges on till but missing on database"
      TabPicture(0)   =   "frmMissingExchanges.frx":038A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "dteSince"
      Tab(0).Control(2)=   "cmdFetch"
      Tab(0).Control(3)=   "txtMissing"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Product differences"
      TabPicture(1)   =   "frmMissingExchanges.frx":03A6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DiffGrid"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdFetchDiffs"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdPrint"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdExport"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdExport 
         BackColor       =   &H00D5C5C1&
         Caption         =   "Export to spreadsheet"
         Height          =   405
         Left            =   90
         MaskColor       =   &H00D5C5C1&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   375
         Width           =   1710
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00D5C5C1&
         Caption         =   "Update all to front"
         Height          =   405
         Left            =   1905
         MaskColor       =   &H00D5C5C1&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   375
         Width           =   1605
      End
      Begin VB.CommandButton cmdFetchDiffs 
         BackColor       =   &H00D5C5C1&
         Caption         =   "Fetch"
         Height          =   405
         Left            =   10365
         MaskColor       =   &H00D5C5C1&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   390
         Width           =   945
      End
      Begin VB.TextBox txtMissing 
         Height          =   5775
         Left            =   -71895
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "frmMissingExchanges.frx":03C2
         Top             =   555
         Width           =   2145
      End
      Begin VB.CommandButton cmdFetch 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Fetch"
         Height          =   405
         Left            =   -72930
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1275
         Width           =   945
      End
      Begin MSComCtl2.DTPicker dteSince 
         Height          =   315
         Left            =   -73425
         TabIndex        =   5
         Top             =   840
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   556
         _Version        =   393216
         Format          =   99352577
         CurrentDate     =   38834
      End
      Begin TrueOleDBGrid60.TDBGrid DiffGrid 
         Height          =   5625
         Left            =   15
         OleObjectBlob   =   "frmMissingExchanges.frx":03C8
         TabIndex        =   7
         Top             =   840
         Width           =   11415
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "(Max. 1000 records fetched at a time)"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   6600
         TabIndex        =   12
         Top             =   585
         Width           =   2910
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Since"
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
         Height          =   435
         Left            =   -74865
         TabIndex        =   6
         Top             =   870
         Width           =   1305
      End
   End
   Begin VB.ComboBox cboTills 
      Height          =   315
      Left            =   1170
      TabIndex        =   1
      Text            =   "cboTills"
      Top             =   285
      Width           =   2145
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tillpoint"
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
      Height          =   345
      Left            =   345
      TabIndex        =   0
      Top             =   315
      Width           =   705
   End
End
Attribute VB_Name = "frmMissingExchanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strMissing
Dim tlTills As z_TextList
Dim rsDiff As ADODB.Recordset
Dim XA As XArrayDB
Private Sub FetchMissing()
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter


    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "CheckSkippedExchanges"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@SINCE", adDate, adParamInput, , dteSince)
    cmd.Parameters.Append par
    par.Value = dteSince
    Set par = cmd.CreateParameter("@TILLPOINT", adVarChar, adParamInput, 20, cboTills)
    cmd.Parameters.Append par
    par.Value = cboTills
    Set par = cmd.CreateParameter("@MISSING", adVarChar, adParamOutput, 3000)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.Execute
    
    strMissing = cmd.Parameters("@Missing").Value
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMissingExchanges.FetchMissing"
End Sub

Private Sub cmdclose_Click()
    SaveLayout Me.DiffGrid, Me.Name & DiffGrid.Name
    SaveFormSize Me.Name, Me.Height, Me.Width

    Unload Me
End Sub

Private Sub cmdExport_Click()
    On Error GoTo errHandler
      Dim fs As New Scripting.FileSystemObject
      Dim sFile As String
      Dim strExecutable As String
          
20        Screen.MousePointer = vbHourglass
          
30        If Not fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
40            fs.CreateFolder oPC.SharedFolderRoot & "\TEMP"
50        End If
60        sFile = oPC.SharedFolderRoot & "\TEMP\FrontBackComparison" & Format(Now(), "DDMMYYHHNN") & ".HTML"
70        If fs.FileExists(sFile) Then
80            fs.DeleteFile sFile, True
90        End If
100       Me.DiffGrid.ExportToFile sFile, True
          
110       Screen.MousePointer = vbDefault
120       If MsgBox("Spreadsheet file saved in: " & sFile & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
                OpenFileWithApplication sFile, enExcel
'130           strExecutable = GetPDFExecutable(oPC.SharedFolderRoot & "\TEMPLATES\DUMMY.XLS")
'140           Shell strExecutable & " " & sFile, vbNormalFocus
150       End If

    Exit Sub
errHandler:
    ErrPreserve
    Screen.MousePointer = vbDefault
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMissingExchanges.cmdExport_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFetch_Click()
    On Error GoTo errHandler
    FetchMissing
    txtMissing = strMissing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMissingExchanges.cmdFetch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFetchDiffs_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    LoadDifferences
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMissingExchanges.cmdFetchDiffs_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdPrint_Click()
Dim PID As String
Dim i As Long
    If XA.Count(1) = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    For i = 0 To XA.UpperBound(1)
        PID = FNS(XA(i, 10))
        TouchRecord PID
    Next i
    Screen.MousePointer = vbDefault
    MsgBox "The products have been updated." & vbCrLf & "Wait a few moments and fetch again to check the updates have been received at the POS station.", vbInformation + vbOKOnly, "Status"
End Sub

Private Sub DiffGrid_ButtonClick(ByVal ColIndex As Integer)
Dim PID As String
    PID = FNS(XA(DiffGrid.Bookmark, 10))
    TouchRecord PID
End Sub

Private Sub TouchRecord(pPID As String)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim SQL As String
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    SQL = "INSERT INTO tPRODUPDATES(PRU_LOG_TYPE,PRU_P_ID,PRU_Code,PRU_EAN," _
            & "PRU_Publisher,PRU_SeriesTitle,PRU_MainAuthor,PRU_Title,PRU_SP,PRU_VATRATE,PRU_SSP,PRU_NDA,PRU_LoyaltyRATE," _
            & "PRU_PTID,PRU_SECID,PRU_MULTIBUYCODE) " _
            & "SELECT 'NEW',P_ID,P_CODE," & "P_EAN,P_PUBLISHER,P_SERIESTITLE,P_MAINAUTHOR," _
            & "P_TITLE,P_SP,dbo.VATRATETOUSE(P_SpecialVat,P_VatRate),P_Special,P_NDA,P_LoyaltyRATE, P_ProductType_ID, vSectionMaster.PSEC_SEC_ID,P_MultibuyCode " _
            & " FROM tPRODUCT LEFT JOIN vSectionMaster ON P_ID = vSectionMaster.PSEC_P_ID   LEFT JOIN vMultibuyCode ON P_ID = vMultibuyCode.PSEC_P_ID" _
            & " WHERE P_ID = '" & pPID & "'"

    oPC.COShort.Execute SQL
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.TouchRecord(pPID)", pPID
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    dteSince = DateAdd("d", -3, Date)
    Set tlTills = New z_TextList
    tlTills.Load ltTillpoint
    LoadCombo cboTills, tlTills
    SetGridLayout Me.DiffGrid, Me.Name & DiffGrid.Name
    SetFormSize Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMissingExchanges.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
    resize
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMissingExchanges.Form_Resize", , EA_NORERAISE
    HandleError
End Sub
Private Sub resize()
    On Error GoTo errHandler
    SSTab.Width = NonNegative_Lng(Me.Width - 400)
    SSTab.Height = NonNegative_Lng(Me.Height - 1700)
    cmdclose.Left = NonNegative_Lng(SSTab.Width - 800)
    cmdclose.TOP = 100
    If SSTab.Tab = 0 Then
        Me.txtMissing.Left = 4000
        Me.txtMissing.TOP = 470
        Me.txtMissing.Width = 3000
        Me.txtMissing.Height = NonNegative_Lng(SSTab.Height - 800)
    Else
        Me.cmdFetchDiffs.Left = NonNegative_Lng(SSTab.Left + SSTab.Width - 2000)
        DiffGrid.Left = SSTab.Left + 90
        DiffGrid.Width = NonNegative_Lng(SSTab.Width - 400)
        DiffGrid.Height = NonNegative_Lng(SSTab.Height - 1200)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMissingExchanges.resize"
End Sub
Private Sub SSTab_Click(PreviousTab As Integer)
    On Error GoTo errHandler
    resize
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMissingExchanges.SSTab_Click(PreviousTab)", PreviousTab, EA_NORERAISE
    HandleError
End Sub

Private Sub LoadDifferences()
    On Error GoTo errHandler
Dim SQL As String
Dim strRemoteServerName As String
Dim OpenResult As Integer
Dim i As Integer

    strRemoteServerName = tlTills.KeyChar(Me.cboTills) & "\PBKSINSTANCE2"
    SQL = " SELECT TOP 1000 a.P_ID PID,dbo.CodeF('',a.P_EAN,0) EAN,a.P_TITLE Description,dbo.CUrrFOrmat(a.P_SP) MainPrice,dbo.CurrFOrmat(b.P_SAPrice) POSPrice,a.P_MultiBuyCode MainMBCode , " _
            & " b.P_MultiBuyCode POSMBCode, a.P_NDA MainNDA,b.P_NDA POSNDA FROM tPRODUCT a LEFT JOIN " _
            & " OPENROWSET('SQLOLEDB','Driver={SQL SERVER};SERVER=" & strRemoteServerName & ";UID=sa;PWD=car', " _
                    & " ' SELECT P_ID,P_EAN,P_TITLE,P_SAPrice,P_NDA,P_MultiBuyCode FROM PBKSFD.dbo.tPRODUCT') as b ON a.P_ID = b.P_ID  " _
                    & " WHERE a.P_SP <> ISNULL(b.P_SAPrice,0) OR a.P_MultiBuyCode  COLLATE DATABASE_DEFAULT <> ISNULL(b.P_MultiBuyCode,'')  COLLATE DATABASE_DEFAULT or a.P_NDA <> ISNULL(b.P_NDA,0) ORDER BY a.P_Title"
    LogSaveToFile SQL
    Set rsDiff = New ADODB.Recordset
    rsDiff.CursorLocation = adUseClient
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If Not XA Is Nothing Then
        XA.Clear
        Set XA = Nothing
    End If
    Set XA = New XArrayDB
    i = 0
    Screen.MousePointer = vbHourglass
    oPC.COShort.CommandTimeout = 600
    rsDiff.Open SQL, oPC.COShort, adOpenStatic, adLockOptimistic
    If rsDiff.RecordCount = 0 Then
            XA.ReDim 0, 1, 1, 1
            XA(0, 1) = "No mismatches found"
    Else
        If rsDiff.RecordCount = 2000 Then
            MsgBox "2000 records found. There are probably more. This tool only fetches 1000 at a time.", vbOKOnly, "Status"
        Else
            MsgBox CStr(rsDiff.RecordCount) & " records found.", vbOKOnly, "Status"
        End If
        Do While Not rsDiff.EOF
            XA.ReDim 0, i, 1, 10
            XA(i, 1) = FNS(rsDiff.Fields("EAN"))
            XA(i, 2) = FNS(rsDiff.Fields("Description"))
            XA(i, 3) = FNS(rsDiff.Fields("MainPrice"))
            XA(i, 4) = FNS(rsDiff.Fields("POSPrice"))
            XA(i, 5) = FNS(rsDiff.Fields("MainMBCode"))
            XA(i, 6) = FNS(rsDiff.Fields("POSMBCode"))
            XA(i, 7) = FNS(rsDiff.Fields("MainNDA"))
            XA(i, 8) = FNS(rsDiff.Fields("POSNDA"))
            XA(i, 9) = ">>>"
            XA(i, 10) = FNS(rsDiff.Fields("PID"))
            i = i + 1
            rsDiff.MoveNext
        Loop
    End If
    Me.DiffGrid.Array = XA
    DiffGrid.ReBind
        Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMissingExchanges.LoadDifferences"
End Sub

