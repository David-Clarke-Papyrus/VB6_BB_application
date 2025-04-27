VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Begin VB.Form frmHistory 
   Caption         =   "Customer history"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   10545
   Begin TabDlg.SSTab SSTab1 
      Height          =   6165
      Left            =   120
      TabIndex        =   0
      Top             =   195
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   10874
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Orders"
      TabPicture(0)   =   "frmCustomerHistory.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCount1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "G1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdExcelExport1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Sales"
      TabPicture(1)   =   "frmCustomerHistory.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdExcelExport2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "G2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblCount2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Appros"
      TabPicture(2)   =   "frmCustomerHistory.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdExcelExport3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "G3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblCount3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.CommandButton cmdExcelExport1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Spreadsheet"
         Height          =   570
         Left            =   8205
         Picture         =   "frmCustomerHistory.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1170
      End
      Begin VB.CommandButton cmdExcelExport2 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Spreadsheet"
         Height          =   570
         Left            =   -66810
         Picture         =   "frmCustomerHistory.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   5475
         Width           =   1170
      End
      Begin VB.CommandButton cmdExcelExport3 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Spreadsheet"
         Height          =   570
         Left            =   -66885
         Picture         =   "frmCustomerHistory.frx":0768
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   5490
         Width           =   1170
      End
      Begin VB.Frame Frame2 
         ForeColor       =   &H8000000D&
         Height          =   810
         Left            =   -74835
         TabIndex        =   22
         Top             =   465
         Width           =   9105
         Begin VB.PictureBox Picture 
            Height          =   600
            Left            =   60
            ScaleHeight     =   540
            ScaleWidth      =   4680
            TabIndex        =   30
            Top             =   150
            Width           =   4740
            Begin VB.OptionButton optAPP12 
               Caption         =   "Last 12 months"
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   2985
               TabIndex        =   33
               Top             =   150
               Width           =   1515
            End
            Begin VB.OptionButton optAPP3 
               Caption         =   "Last 3 months"
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   1470
               TabIndex        =   32
               Top             =   150
               Width           =   1365
            End
            Begin VB.OptionButton optAPP1 
               Caption         =   "Last month"
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   240
               TabIndex        =   31
               Top             =   150
               Value           =   -1  'True
               Width           =   1080
            End
         End
         Begin VB.CommandButton cmdFetch3 
            BackColor       =   &H00D3D3CB&
            Height          =   435
            Left            =   8310
            Picture         =   "frmCustomerHistory.frx":0AF2
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   210
            Width           =   495
         End
         Begin VB.CheckBox chkOS3 
            Caption         =   "OS only"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   6975
            TabIndex        =   23
            Top             =   225
            Width           =   1200
         End
         Begin MSComCtl2.DTPicker DTP3 
            Height          =   300
            Left            =   5310
            TabIndex        =   25
            Top             =   225
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   244973569
            CurrentDate     =   39627
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Since"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   4575
            TabIndex        =   26
            Top             =   285
            Width           =   645
         End
      End
      Begin VB.Frame Frame1 
         ForeColor       =   &H8000000D&
         Height          =   810
         Left            =   -74835
         TabIndex        =   10
         Top             =   465
         Width           =   9105
         Begin VB.OptionButton optSAL12 
            Caption         =   "Last 12 months"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   2865
            TabIndex        =   14
            Top             =   225
            Width           =   1515
         End
         Begin VB.CommandButton cmdFetch2 
            BackColor       =   &H00D3D3CB&
            Height          =   435
            Left            =   8325
            Picture         =   "frmCustomerHistory.frx":0E7C
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   210
            Width           =   495
         End
         Begin VB.OptionButton optSAL3 
            Caption         =   "Last 3 months"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1350
            TabIndex        =   12
            Top             =   225
            Width           =   1365
         End
         Begin VB.OptionButton optSal1 
            Caption         =   "Last month"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   225
            Value           =   -1  'True
            Width           =   1080
         End
         Begin MSComCtl2.DTPicker DTP2 
            Height          =   300
            Left            =   5310
            TabIndex        =   15
            Top             =   225
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   244973569
            CurrentDate     =   39627
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Since"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   4575
            TabIndex        =   16
            Top             =   285
            Width           =   645
         End
      End
      Begin VB.Frame Frame3 
         ForeColor       =   &H8000000D&
         Height          =   810
         Left            =   150
         TabIndex        =   1
         Top             =   465
         Width           =   9210
         Begin VB.CheckBox chkOS 
            Caption         =   "OS only"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   7005
            TabIndex        =   18
            Top             =   225
            Width           =   1200
         End
         Begin VB.OptionButton optCO1 
            Caption         =   "Last month"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   255
            Value           =   -1  'True
            Width           =   1080
         End
         Begin VB.OptionButton optCO3 
            Caption         =   "Last 3 months"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1365
            TabIndex        =   4
            Top             =   240
            Width           =   1365
         End
         Begin VB.CommandButton cmdFetch1 
            BackColor       =   &H00D3D3CB&
            Height          =   435
            Left            =   8325
            Picture         =   "frmCustomerHistory.frx":1206
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   210
            Width           =   495
         End
         Begin VB.OptionButton optCO12 
            Caption         =   "Last 12 months"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   2865
            TabIndex        =   2
            Top             =   240
            Width           =   1515
         End
         Begin MSComCtl2.DTPicker DTP1 
            Height          =   300
            Left            =   5280
            TabIndex        =   6
            Top             =   225
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   244973569
            CurrentDate     =   39627
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Since"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   4290
            TabIndex        =   7
            Top             =   285
            Width           =   900
         End
      End
      Begin TrueOleDBGrid60.TDBGrid G1 
         Height          =   4125
         Left            =   180
         OleObjectBlob   =   "frmCustomerHistory.frx":1590
         TabIndex        =   8
         Top             =   1335
         Width           =   9165
      End
      Begin TrueOleDBGrid60.TDBGrid G2 
         Height          =   4080
         Left            =   -74820
         OleObjectBlob   =   "frmCustomerHistory.frx":58A7
         TabIndex        =   17
         Top             =   1350
         Width           =   9165
      End
      Begin TrueOleDBGrid60.TDBGrid G3 
         Height          =   4125
         Left            =   -74835
         OleObjectBlob   =   "frmCustomerHistory.frx":9BBE
         TabIndex        =   21
         Top             =   1320
         Width           =   9165
      End
      Begin VB.Label lblCount3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   -74850
         TabIndex        =   20
         Top             =   5520
         Width           =   3645
      End
      Begin VB.Label lblCount2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   -74790
         TabIndex        =   19
         Top             =   5475
         Width           =   3645
      End
      Begin VB.Label lblCount1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   240
         TabIndex        =   9
         Top             =   5490
         Width           =   3645
      End
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngCUSTID As Long
Dim dteSelected As Date
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim XA As New XArrayDB
Dim XB As New XArrayDB
Dim XC As New XArrayDB
Dim bOSOnly As Boolean
Dim bOSOnly2 As Boolean
Dim bOSOnly3 As Boolean

Public Sub component(CustID As Long)
    On Error GoTo errHandler
    lngCUSTID = CustID
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.component(CustID)", CustID
End Sub
Private Sub LoadOrders()
    On Error GoTo errHandler
Dim OpenResult As Integer
    '--------------
    OpenResult = oPC.OpenDBSHort
    '--------------

    Screen.MousePointer = vbHourglass

    Set rs1 = New ADODB.Recordset
    rs1.CursorLocation = adUseClient
    
    If optCO1 = True Then
        If bOSOnly = True Then
            rs1.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,COL_QTY,COL_PRICE FROM tCOL JOIN tTR ON COL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID JOIN tPRODUCT ON COL_P_ID = P_ID" _
                & " WHERE ISNULL(COL_FULFILLED,'OS') IN ('OS') AND TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & "" _
                & " AND (TR_Date >= DATEADD(m,-1,dbo.startOfDay(Getdate()))) AND (dbo.tCOL.COL_DateReplaced IS NULL)", oPC.COShort, adOpenForwardOnly, adLockOptimistic
        Else
            rs1.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,COL_QTY,COL_PRICE FROM tCOL JOIN tTR ON COL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID JOIN tPRODUCT ON COL_P_ID = P_ID" _
                & " WHERE  TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & " AND (TR_Date >= DATEADD(m,-1,dbo.startOfDay(Getdate()))) AND (dbo.tCOL.COL_DateReplaced IS NULL)", oPC.COShort, adOpenForwardOnly, adLockOptimistic
        End If
    Else
    If Me.optCO3 = True Then
        If bOSOnly = True Then
            rs1.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,COL_QTY,COL_PRICE FROM tCOL JOIN tTR ON COL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID  JOIN tPRODUCT ON COL_P_ID = P_ID" _
                & " WHERE  ISNULL(COL_FULFILLED,'OS') IN ('OS') AND TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & " AND (TR_Date >= DATEADD(m,-3,dbo.startOfDay(Getdate())))", oPC.COShort
        Else
            rs1.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,COL_QTY,COL_PRICE FROM tCOL JOIN tTR ON COL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID  JOIN tPRODUCT ON COL_P_ID = P_ID" _
                & " WHERE  TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & " AND (TR_Date >= DATEADD(m,-3,dbo.startOfDay(Getdate())))", oPC.COShort
        End If
    Else
    If Me.optCO12 = True Then
        If bOSOnly = True Then
            rs1.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,COL_QTY,COL_PRICE FROM tCOL JOIN tTR ON COL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID  JOIN tPRODUCT ON COL_P_ID = P_ID" _
                & " WHERE  ISNULL(COL_FULFILLED,'OS') IN ('OS') AND TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & " AND (TR_Date >= DATEADD(m,-12,dbo.startOfDay(Getdate())))", oPC.COShort
        Else
            rs1.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,COL_QTY,COL_PRICE FROM tCOL JOIN tTR ON COL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID  JOIN tPRODUCT ON COL_P_ID = P_ID" _
                & " WHERE  TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & " AND (TR_Date >= DATEADD(m,-12,dbo.startOfDay(Getdate())))", oPC.COShort
        End If
    Else
        If bOSOnly = True Then
            rs1.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,COL_QTY,COL_PRICE FROM tCOL JOIN tTR ON COL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID  JOIN tPRODUCT ON COL_P_ID = P_ID" _
                & " WHERE  ISNULL(COL_FULFILLED,'OS') IN ('OS') AND TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & " AND (TR_Date >= '" & ReverseDate(DTP2) & "')  AND (dbo.tCOL.COL_DateReplaced IS NULL)", oPC.COShort
        Else
            rs1.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,COL_QTY,COL_PRICE FROM tCOL JOIN tTR ON COL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID  JOIN tPRODUCT ON COL_P_ID = P_ID" _
                & " WHERE  TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & " AND (TR_Date >= '" & ReverseDate(DTP2) & "')  AND (dbo.tCOL.COL_DateReplaced IS NULL)", oPC.COShort
        End If
    End If
    End If
    End If
    If rs1.eof Then
        Screen.MousePointer = vbDefault
        MsgBox "No records found", vbInformation, "Status"
        '    --------------
            If OpenResult = 0 Then oPC.DisconnectDBShort  'if the recent open command actually opened a connection then close it
        '    --------------
        
        Exit Sub
    End If
    LoadGrid1
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 1, 0, XTYPE_STRING
    G1.Array = XA
    G1.ReBind
    G1.Refresh
    lblCount1.Caption = CStr(rs1.RecordCount) & " records"
    Screen.MousePointer = vbDefault
'    --------------
    If OpenResult = 0 Then oPC.DisconnectDBShort  'if the recent open command actually opened a connection then close it
'    --------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.LoadOrders"
End Sub

Private Sub LoadGrid1()
    On Error GoTo errHandler
Dim lngIndex As Long
Dim tmp As String
Dim i As Integer
Dim iRecs As Long
Dim lngArrayRows As Long

    i = 0
    Set XA = Nothing
    Set XA = New XArrayDB
    XA.Clear
    iRecs = i
    lngIndex = 1
    lngArrayRows = rs1.RecordCount
    XA.ReDim 1, lngArrayRows, 1, 10
    If Not rs1.eof Then
        rs1.MoveFirst
        Do While Not rs1.eof
                XA.Value(lngIndex, 1) = Format(FNS(rs1.fields("TR_DATE")), "dd-mm-yyyy")
                XA.Value(lngIndex, 2) = FNS(rs1.fields("TR_CODE"))
                XA.Value(lngIndex, 3) = FNS(rs1.fields("CodeF"))
                XA.Value(lngIndex, 4) = FNS(rs1.fields("P_TITLE"))
                XA.Value(lngIndex, 5) = FNN(rs1.fields("COL_QTY"))
                XA.Value(lngIndex, 6) = Format(FNN(rs1.fields("COL_PRICE")) / oPC.Configuration.DefaultCurrency.Divisor, "R ###,##0.00")
                lngIndex = lngIndex + 1
                rs1.MoveNext
        Loop
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.LoadGrid1"
End Sub



Private Sub chkOS_Click()
    On Error GoTo errHandler
    bOSOnly = (chkOS = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.chkOS_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdFetch1_Click()
    On Error GoTo errHandler
    LoadOrders
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.cmdFetch1_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub DTP1_GotFocus()
    On Error GoTo errHandler
    Me.optCO1 = False
    Me.optCO3 = False
    Me.optCO12 = False
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.DTP1_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub G1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    If XA.Count(1) > 0 Then
       XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType1(ColIndex + 1)
    End If
    
    G1.Refresh
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType1(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1
            GetRowType1 = XTYPE_DATE
        Case 2, 3, 4, 5
            GetRowType1 = XTYPE_STRING
        Case 6
            GetRowType1 = XTYPE_NUMBER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.GetRowType1(ColIndex)", ColIndex
End Function
Private Sub cmdExcelExport1_Click()
    On Error GoTo errHandler
Dim xls As New ActiveReportsExcelExport.ARExportExcel
Dim sFile As String
Dim bSave As Boolean
Dim fs As New FileSystemObject
Dim rpt As New arCustomerOrders_ForExcel
Dim i As Long
Dim strExecutable As String

    If XA.UpperBound(1) = 0 Then
  MsgBox "There are no lines to print.", vbOKOnly, "Can't do this"
  Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    rpt.component XA, "Customer orders"
    rpt.Run False
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
  fs.CreateFolder oPC.LocalFolder & "\TEMP"
    End If
    sFile = oPC.LocalFolder & "\TEMP\CustomerOrders.XLS"
    If fs.FileExists(sFile) Then
  fs.DeleteFile sFile, True
    End If
    xls.FileName = sFile
    If rpt.Pages.Count > 0 Then
  xls.Export rpt.Pages
    End If
    Screen.MousePointer = vbDefault
    If MsgBox("Spreadsheet file saved in: " & sFile & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
      strExecutable = GetPDFExecutable(sFile)
    If strExecutable = "" Then
        MsgBox "There is no application set on this computer to open the file: " & sFile & ". The document cannot be displayed", vbOKOnly, "Can't do this"
    Else
      F_7_AB_1_ShellAndWaitSimple strExecutable & " " & sFile, vbNormalFocus, 10000
    End If
    End If

    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmHistory: cmdExcelExport1_Click"  'unknown source
        If errRepeat < 5 Then
            Err.Clear
            Exit Sub
        Else
            LogSaveToFile "Access violation in frmHistory: cmdExcelExport1_Click after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If

    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.cmdExcelExport1_Click", , EA_NORERAISE
    HandleError
End Sub

'=============================================================================
'S A L E S ---------------------------------------------------
Private Sub cmdFetch2_Click()
    On Error GoTo errHandler
    LoadSales
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.cmdFetch2_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub DTP2_GotFocus()
    On Error GoTo errHandler
    Me.optSal1 = False
    Me.optSAL3 = False
    Me.optSAL12 = False
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.DTP2_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadSales()
    On Error GoTo errHandler
Dim OpenResult As Integer
    '--------------
    OpenResult = oPC.OpenDBSHort
    '--------------

    Screen.MousePointer = vbHourglass

    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    
    If optSal1 = True Then
            rs2.open "Select * FROM vSalesForCustomer WHERE TPID = " & CStr(lngCUSTID) & " AND (DTE >= DATEADD(m,-1,dbo.startOfDay(Getdate())))", oPC.COShort, adOpenForwardOnly, adLockOptimistic
    Else
    If Me.optSAL3 = True Then
            rs2.open "Select * FROM vSalesForCustomer WHERE TPID = " & CStr(lngCUSTID) & " AND (DTE >= DATEADD(m,-3,dbo.startOfDay(Getdate())))", oPC.COShort, adOpenForwardOnly, adLockOptimistic
    Else
    If Me.optSAL12 = True Then
            rs2.open "Select * FROM vSalesForCustomer WHERE TPID = " & CStr(lngCUSTID) & " AND (DTE >= DATEADD(m,-12,dbo.startOfDay(Getdate())))", oPC.COShort, adOpenForwardOnly, adLockOptimistic
    Else
            rs2.open "Select * FROM vSalesForCustomer WHERE TPID = " & CStr(lngCUSTID) & " AND (DTE >=  '" & ReverseDate(DTP2) & "')", oPC.COShort, adOpenForwardOnly, adLockOptimistic
    End If
    End If
    End If
    If rs2.eof Then
        Screen.MousePointer = vbDefault
        MsgBox "No records found", vbInformation, "Status"
        '    --------------
            If OpenResult = 0 Then oPC.DisconnectDBShort  'if the recent open command actually opened a connection then close it
        '    --------------
        
        Exit Sub
    End If
    LoadGrid2
    XB.QuickSort XB.LowerBound(1), XB.UpperBound(1), 1, 0, XTYPE_DATE, 2, 0, XTYPE_STRING
    G2.Array = XB
    G2.ReBind
    G2.Refresh
    lblCount2.Caption = CStr(rs2.RecordCount) & " records"
    Screen.MousePointer = vbDefault
'    --------------
    If OpenResult = 0 Then oPC.DisconnectDBShort  'if the recent open command actually opened a connection then close it
'    --------------

    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.LoadSales"
End Sub

Private Sub LoadGrid2()
    On Error GoTo errHandler
Dim lngIndex As Long
Dim tmp As String
Dim i As Integer
Dim iRecs As Long
Dim lngArrayRows As Long

    i = 0
    Set XB = Nothing
    Set XB = New XArrayDB
    XB.Clear
    iRecs = i
    lngIndex = 1
    lngArrayRows = rs2.RecordCount
    XB.ReDim 1, lngArrayRows, 1, 10
    If Not rs2.eof Then
        rs2.MoveFirst
        Do While Not rs2.eof
                XB.Value(lngIndex, 1) = Format(FNS(rs2.fields("DTE")), "dd-mm-yyyy")
                XB.Value(lngIndex, 2) = FNS(rs2.fields("DOCCODE"))
                XB.Value(lngIndex, 3) = FNS(rs2.fields("CodeF"))
                XB.Value(lngIndex, 4) = FNS(rs2.fields("P_TITLE"))
                XB.Value(lngIndex, 5) = FNN(rs2.fields("QTY"))
                XB.Value(lngIndex, 6) = Format(FNN(rs2.fields("VALUE")), "R ###,##0.00")
                lngIndex = lngIndex + 1
                rs2.MoveNext
        Loop
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.LoadGrid2"
End Sub

Private Sub G2_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XB.QuickSort XB.LowerBound(1), XB.UpperBound(1), ColIndex + 1, Direction, GetRowType3(ColIndex + 1)
    
    G2.Refresh
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.G2_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType2(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1
            GetRowType2 = XTYPE_DATE
        Case 2, 3, 4, 5
            GetRowType2 = XTYPE_STRING
        Case 6
            GetRowType2 = XTYPE_NUMBER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.GetRowType2(ColIndex)", ColIndex
End Function
Private Sub cmdExcelExport2_Click()
    On Error GoTo errHandler
Dim xls As New ActiveReportsExcelExport.ARExportExcel
Dim sFile As String
Dim bSave As Boolean
Dim fs As New FileSystemObject
Dim rpt As New arCustomerSales_ForExcel
Dim i As Long
Dim strExecutable As String

    If XB.UpperBound(1) = 0 Then
        MsgBox "There are no lines to print.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    rpt.component XB, "Customer appros"
    rpt.Run False
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder oPC.LocalFolder & "\TEMP"
    End If
    sFile = oPC.LocalFolder & "\TEMP\CustomerSales.XLS"
    If fs.FileExists(sFile) Then
        fs.DeleteFile sFile, True
    End If
    xls.FileName = sFile
    If rpt.Pages.Count > 0 Then
        xls.Export rpt.Pages
    End If
    Screen.MousePointer = vbDefault
    If MsgBox("Spreadsheet file saved in: " & sFile & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
            strExecutable = GetPDFExecutable(sFile)
          If strExecutable = "" Then
              MsgBox "There is no application set on this computer to open the file: " & sFile & ". The document cannot be displayed", vbOKOnly, "Can't do this"
          Else
            F_7_AB_1_ShellAndWaitSimple strExecutable & " " & sFile, vbNormalFocus, 10000
          End If
    End If

    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmHistory: cmdExcelExport2_Click"  'unknown source
        If errRepeat < 5 Then
            Err.Clear
            Exit Sub
        Else
            LogSaveToFile "Access violation in frmHistory: cmdExcelExport2_Click after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.cmdExcelExport2_Click", , EA_NORERAISE
    HandleError
End Sub

'=============================================================================
'A P P R O S -------------------------------------------------
Private Sub cmdFetch3_Click()
    On Error GoTo errHandler
    LoadAppros
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.cmdFetch3_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkOS3_Click()
    On Error GoTo errHandler
    bOSOnly3 = (chkOS3 = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.chkOS3_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub DTP3_GotFocus()
    On Error GoTo errHandler
    Me.optAPP1 = False
    Me.optAPP3 = False
    Me.optAPP12 = False
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.DTP3_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadAppros()
    On Error GoTo errHandler
Dim OpenResult As Integer
    '--------------
    OpenResult = oPC.OpenDBSHort
    '--------------

    Screen.MousePointer = vbHourglass

    Set rs3 = New ADODB.Recordset
    rs3.CursorLocation = adUseClient
    
    If optAPP1 = True Then
        If bOSOnly3 = True Then
            rs3.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,(ISNULL(APPL_QTY,0) -ISNULL(APPL_QTYRETURNED,0)) as QTY,APPL_PRICE FROM tAPPL JOIN tTR ON APPL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID JOIN tPRODUCT ON APPL_P_ID = P_ID" _
                & " WHERE ISNULL(APPL_FULFILLED,'OS') IN ('OS') AND TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & "" _
                & " AND (TR_Date >= DATEADD(m,-1,dbo.startOfDay(Getdate())))", oPC.COShort, adOpenForwardOnly, adLockOptimistic
        Else
            rs3.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,(ISNULL(APPL_QTY,0) -ISNULL(APPL_QTYRETURNED,0)) as QTY,APPL_PRICE FROM tAPPL JOIN tTR ON APPL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID JOIN tPRODUCT ON APPL_P_ID = P_ID" _
                & " WHERE  TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & " AND (TR_Date >= DATEADD(m,-1,dbo.startOfDay(Getdate())))", oPC.COShort, adOpenForwardOnly, adLockOptimistic
        End If
    Else
    If Me.optAPP3 = True Then
        If bOSOnly3 = True Then
            rs3.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,(ISNULL(APPL_QTY,0) -ISNULL(APPL_QTYRETURNED,0)) as QTY,APPL_PRICE FROM tAPPL JOIN tTR ON APPL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID  JOIN tPRODUCT ON APPL_P_ID = P_ID" _
                & " WHERE  ISNULL(APPL_FULFILLED,'OS') IN ('OS') AND TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & " AND (TR_Date >= DATEADD(m,-3,dbo.startOfDay(Getdate())))", oPC.COShort
        Else
            rs3.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,(ISNULL(APPL_QTY,0) -ISNULL(APPL_QTYRETURNED,0)) as QTY,APPL_PRICE FROM tAPPL JOIN tTR ON APPL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID  JOIN tPRODUCT ON APPL_P_ID = P_ID" _
                & " WHERE  TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & " AND (TR_Date >= DATEADD(m,-3,dbo.startOfDay(Getdate())))", oPC.COShort
        End If
    Else
    If Me.optAPP12 = True Then
        If bOSOnly3 = True Then
            rs3.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,(ISNULL(APPL_QTY,0) -ISNULL(APPL_QTYRETURNED,0)) as QTY,APPL_PRICE FROM tAPPL JOIN tTR ON APPL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID  JOIN tPRODUCT ON APPL_P_ID = P_ID" _
                & " WHERE  ISNULL(APPL_FULFILLED,'OS') IN ('OS') AND TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & " AND (TR_Date >= DATEADD(m,-12,dbo.startOfDay(Getdate())))", oPC.COShort
        Else
            rs3.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,(ISNULL(APPL_QTY,0) -ISNULL(APPL_QTYRETURNED,0)) as QTY,APPL_PRICE FROM tAPPL JOIN tTR ON APPL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID  JOIN tPRODUCT ON APPL_P_ID = P_ID" _
                & " WHERE  TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & " AND (TR_Date >= DATEADD(m,-12,dbo.startOfDay(Getdate())))", oPC.COShort
        End If
    Else
        If bOSOnly3 = True Then
            rs3.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,(ISNULL(APPL_QTY,0) -ISNULL(APPL_QTYRETURNED,0)) as QTY,APPL_PRICE FROM tAPPL JOIN tTR ON APPL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID  JOIN tPRODUCT ON APPL_P_ID = P_ID" _
                & " WHERE  ISNULL(APPL_FULFILLED,'OS') IN ('OS') AND TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & " AND (TR_Date >= '" & ReverseDate(DTP2) & "')", oPC.COShort
        Else
            rs3.open "Select TR_DATE,TR_CODE,dbo.CODEF(P_CODE,P_EAN,0) as CodeF,P_TITLE,(ISNULL(APPL_QTY,0) -ISNULL(APPL_QTYRETURNED,0)) as QTY,APPL_PRICE FROM tAPPL JOIN tTR ON APPL_TR_ID = TR_ID " _
                        & " JOIN tTP ON TR_TP_ID = TP_ID  JOIN tPRODUCT ON APPL_P_ID = P_ID" _
                & " WHERE  TR_STATUS IN (3,4) AND TP_ID = " & CStr(lngCUSTID) & " AND (TR_Date >= '" & ReverseDate(DTP2) & "')", oPC.COShort
        End If
    End If
    End If
    End If
    If rs3.eof Then
        Screen.MousePointer = vbDefault
        MsgBox "No records found", vbInformation, "Status"
        '    --------------
            If OpenResult = 0 Then oPC.DisconnectDBShort  'if the recent open command actually opened a connection then close it
        '    --------------
        
        Exit Sub
    End If
    LoadGrid3
    XC.QuickSort XC.LowerBound(1), XC.UpperBound(1), 1, 0, XTYPE_DATE, 2, 0, XTYPE_STRING
    G3.Array = XC
    G3.ReBind
    G3.Refresh
    lblCount3.Caption = CStr(rs3.RecordCount) & " records"
    Screen.MousePointer = vbDefault
'    --------------
    If OpenResult = 0 Then oPC.DisconnectDBShort  'if the recent open command actually opened a connection then close it
'    --------------

    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.LoadAppros"
End Sub

Private Sub LoadGrid3()
    On Error GoTo errHandler
Dim lngIndex As Long
Dim tmp As String
Dim i As Integer
Dim iRecs As Long
Dim lngArrayRows As Long

    i = 0
    Set XC = Nothing
    Set XC = New XArrayDB
    XC.Clear
    iRecs = i
    lngIndex = 1
    lngArrayRows = rs3.RecordCount
    XC.ReDim 1, lngArrayRows, 1, 10
    If Not rs3.eof Then
        rs3.MoveFirst
        Do While Not rs3.eof
                XC.Value(lngIndex, 1) = Format(FNS(rs3.fields("TR_DATE")), "dd-mm-yyyy")
                XC.Value(lngIndex, 2) = FNS(rs3.fields("TR_CODE"))
                XC.Value(lngIndex, 3) = FNS(rs3.fields("CodeF"))
                XC.Value(lngIndex, 4) = FNS(rs3.fields("P_TITLE"))
                XC.Value(lngIndex, 5) = FNN(rs3.fields("QTY"))
                XC.Value(lngIndex, 6) = Format(FNN(rs3.fields("APPL_PRICE")) / oPC.Configuration.DefaultCurrency.Divisor, "R ###,##0.00")
                lngIndex = lngIndex + 1
                rs3.MoveNext
        Loop
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.LoadGrid3"
End Sub

Private Sub G3_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XC.QuickSort XC.LowerBound(1), XC.UpperBound(1), ColIndex + 1, Direction, GetRowType3(ColIndex + 1)
    
    G3.Refresh
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.G3_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType3(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1
            GetRowType3 = XTYPE_DATE
        Case 2, 3, 4, 5
            GetRowType3 = XTYPE_STRING
        Case 6
            GetRowType3 = XTYPE_NUMBER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.GetRowType3(ColIndex)", ColIndex
End Function
Private Sub cmdExcelExport3_Click()
    On Error GoTo errHandler
Dim xls As New ActiveReportsExcelExport.ARExportExcel
Dim sFile As String
Dim bSave As Boolean
Dim fs As New FileSystemObject
Dim rpt As New arCustomerAppros_ForExcel
Dim i As Long
Dim strExecutable As String

    If XC.UpperBound(1) = 0 Then
        MsgBox "There are no lines to print.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    rpt.component XC, "Customer appros"
    rpt.Run False
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder oPC.LocalFolder & "\TEMP"
    End If
    sFile = oPC.LocalFolder & "\TEMP\CustomerAppros.XLS"
    If fs.FileExists(sFile) Then
        fs.DeleteFile sFile, True
    End If
    xls.FileName = sFile
    If rpt.Pages.Count > 0 Then
        xls.Export rpt.Pages
    End If
    Screen.MousePointer = vbDefault
    If MsgBox("Spreadsheet file saved in: " & sFile & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
            strExecutable = GetPDFExecutable(sFile)
          If strExecutable = "" Then
              MsgBox "There is no application set on this computer to open the file: " & sFile & ". The document cannot be displayed", vbOKOnly, "Can't do this"
          Else
            F_7_AB_1_ShellAndWaitSimple strExecutable & " " & sFile, vbNormalFocus, 10000
          End If
    End If

    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmHistory: cmdExcelExport3_Click"  'unknown source
        If errRepeat < 5 Then
            Err.Clear
            Exit Sub
        Else
            LogSaveToFile "Access violation in frmHistory: cmdExcelExport3_Click after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.cmdExcelExport3_Click", , EA_NORERAISE
    HandleError
End Sub


'===================================================================================

Private Sub Form_Load()
    On Error GoTo errHandler
Dim i As Integer

    Width = 10100
    Height = 7000
    SSTab1.Tab = 0
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "A", CStr(i), G1.Columns(i - 1).Width)
    Next
    For i = 1 To G2.Columns.Count
        G2.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "B", CStr(i), G2.Columns(i - 1).Width)
    Next
    For i = 1 To G3.Columns.Count
        G3.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "C", CStr(i), G3.Columns(i - 1).Width)
    Next
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.G1, Me.Name & "A"
    SaveLayout Me.G2, Me.Name & "B"
    SaveLayout Me.G3, Me.Name & "C"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.mnuSaveLayout"
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = False
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuFulfil.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Forms(0).mnuCopyLines.Enabled = False
    Forms(0).mnuPastelines.Enabled = False
    Forms(0).mnuPastelinestoNEW = False
   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHistory.SetMenu"
End Sub

