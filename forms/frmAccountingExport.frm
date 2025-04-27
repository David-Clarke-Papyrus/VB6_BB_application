VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmAccountingExportManagement 
   Caption         =   "Documents for export to accounting "
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00CED0BF&
      Caption         =   "Save column widths"
      Height          =   285
      Left            =   8460
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   90
      Width           =   1620
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6585
      Left            =   75
      TabIndex        =   0
      Top             =   135
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   11615
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Documents "
      TabPicture(0)   =   "frmAccountingExport.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCount"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "G"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdExport"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdUnsent"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Days"
      TabPicture(1)   =   "frmAccountingExport.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblCount2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "G2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame3 
         Caption         =   "Show transmission log"
         ForeColor       =   &H8000000D&
         Height          =   975
         Left            =   -74745
         TabIndex        =   25
         Top             =   525
         Width           =   9105
         Begin VB.OptionButton optG2DBY 
            Caption         =   "Day before yesterday"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   2355
            TabIndex        =   32
            Top             =   375
            Width           =   2280
         End
         Begin VB.CommandButton cmdFetch2 
            BackColor       =   &H00D3D3CB&
            Height          =   435
            Left            =   8325
            Picture         =   "frmAccountingExport.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   315
            Width           =   495
         End
         Begin VB.OptionButton optG2Yesterday 
            Caption         =   "Yesterday"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1080
            TabIndex        =   27
            Top             =   360
            Width           =   1050
         End
         Begin VB.OptionButton optG2Today 
            Caption         =   "Today"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Value           =   -1  'True
            Width           =   825
         End
         Begin MSComCtl2.DTPicker DTP2 
            Height          =   300
            Left            =   5895
            TabIndex        =   29
            Top             =   330
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   51249153
            CurrentDate     =   39627
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Select day"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   4905
            TabIndex        =   30
            Top             =   390
            Width           =   900
         End
      End
      Begin VB.CommandButton cmdUnsent 
         BackColor       =   &H00D3D3CB&
         Height          =   390
         Left            =   1950
         Picture         =   "frmAccountingExport.frx":03C2
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1815
         Width           =   795
      End
      Begin VB.Frame Frame1 
         Caption         =   "Show documents since"
         ForeColor       =   &H8000000D&
         Height          =   1185
         Left            =   225
         TabIndex        =   8
         Top             =   495
         Width           =   9105
         Begin VB.OptionButton optSince 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   5445
            TabIndex        =   35
            Top             =   240
            Width           =   345
         End
         Begin VB.OptionButton opttoday 
            Caption         =   "Today"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   225
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.OptionButton optYesterday 
            Caption         =   "Yesterday"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1080
            TabIndex        =   15
            Top             =   225
            Width           =   1050
         End
         Begin VB.OptionButton opt7 
            Caption         =   "last 7 days"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   2355
            TabIndex        =   14
            Top             =   225
            Width           =   1050
         End
         Begin VB.OptionButton opt14 
            Caption         =   "last 14 days"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   3645
            TabIndex        =   13
            Top             =   225
            Width           =   1140
         End
         Begin VB.CommandButton cmdFetch 
            BackColor       =   &H00D3D3CB&
            Height          =   765
            Left            =   8085
            Picture         =   "frmAccountingExport.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   195
            Width           =   915
         End
         Begin VB.Frame V 
            Height          =   555
            Left            =   75
            TabIndex        =   9
            Top             =   510
            Width           =   3120
            Begin VB.OptionButton optByIssuedDate 
               Caption         =   "By issued date"
               ForeColor       =   &H8000000D&
               Height          =   360
               Left            =   1725
               TabIndex        =   11
               Top             =   135
               Width           =   1350
            End
            Begin VB.OptionButton optByDocDate 
               Caption         =   "By document date"
               ForeColor       =   &H8000000D&
               Height          =   360
               Left            =   105
               TabIndex        =   10
               Top             =   120
               Value           =   -1  'True
               Width           =   1590
            End
         End
         Begin MSComCtl2.DTPicker dtSince 
            Height          =   300
            Left            =   6345
            TabIndex        =   17
            Top             =   195
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   51249153
            CurrentDate     =   39627
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Since"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   5610
            TabIndex        =   18
            Top             =   255
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Height          =   555
         Left            =   4965
         TabIndex        =   2
         Top             =   1725
         Width           =   4380
         Begin VB.OptionButton optALL 
            Caption         =   "All"
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   3615
            TabIndex        =   7
            Top             =   150
            Value           =   -1  'True
            Width           =   600
         End
         Begin VB.OptionButton optCN 
            Caption         =   "CN"
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   2805
            TabIndex        =   6
            Top             =   150
            Width           =   570
         End
         Begin VB.OptionButton optINV 
            Caption         =   "Inv."
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   1995
            TabIndex        =   5
            Top             =   150
            Width           =   615
         End
         Begin VB.OptionButton optRET 
            Caption         =   "Retn."
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   1050
            TabIndex        =   4
            Top             =   150
            Width           =   690
         End
         Begin VB.OptionButton optSI 
            Caption         =   "Supp.Inv"
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   45
            TabIndex        =   3
            Top             =   150
            Width           =   945
         End
      End
      Begin VB.CommandButton cmdExport 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Export to Head office"
         Height          =   540
         Left            =   6675
         MaskColor       =   &H00D3D3CB&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5715
         Width           =   2700
      End
      Begin TrueOleDBGrid60.TDBGrid G 
         Height          =   3180
         Left            =   210
         OleObjectBlob   =   "frmAccountingExport.frx":0AD6
         TabIndex        =   19
         Top             =   2415
         Width           =   9165
      End
      Begin TrueOleDBGrid60.TDBGrid G2 
         Height          =   4125
         Left            =   -74745
         OleObjectBlob   =   "frmAccountingExport.frx":5304
         TabIndex        =   24
         Top             =   1920
         Width           =   9165
      End
      Begin VB.Label lblCount2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -74715
         TabIndex        =   33
         Top             =   6120
         Width           =   2445
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Day reflects most recently transmitted date (in the case of re-transmissions)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74430
         TabIndex        =   31
         Top             =   1695
         Width           =   5910
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "All un-sent documents"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   255
         TabIndex        =   23
         Top             =   1905
         Width           =   1560
      End
      Begin VB.Label lblCount 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   540
         TabIndex        =   21
         Top             =   5790
         Width           =   2445
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Filter"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4455
         TabIndex        =   20
         Top             =   1950
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmAccountingExportManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strBy As String
Dim dteSince As Date
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim dteSelected As Date

Private Sub cmdExport_Click()
Dim oSQL As New z_SQL
    oSQL.ExportToHO
End Sub

Private Sub cmdFetch2_Click()
    Screen.MousePointer = vbHourglass

    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    
    If optG2Today = True Then
            dteSelected = Date
            rs2.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
                & " WHERE (TR_DateToPastel >= dbo.startOfDay(Getdate()) and TR_DateToPastel < dbo.endofday(GetDate()))", oPC.CO
    Else
    If Me.optG2Yesterday = True Then
            dteSelected = DateAdd("d", -1, Date)
            rs2.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
                & " WHERE (TR_DateToPastel >= dbo.startOfDay(DATEADD(d,-1,Getdate())) and TR_DateToPastel < dbo.endofday(DATEADD(d,-1,Getdate())))", oPC.CO
    Else
    If Me.optG2DBY = True Then
            dteSelected = DateAdd("d", -2, Date)
            rs2.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
                & " WHERE (TR_DateToPastel >= dbo.startOfDay(DATEADD(d,-2,Getdate())) and TR_DateToPastel < dbo.endofday(DATEADD(d,-2,Getdate())))", oPC.CO
    Else
            dteSelected = DTP2
            rs2.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
                & " WHERE (TR_DateToPastel >= dbo.startOfDay('" & Me.DTP2 & " ') and TR_DateToPastel < dbo.endofday('" & Me.DTP2 & " '))", oPC.CO
    End If
    End If
    End If
    If rs2.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "No records found", vbInformation, "Status"
        Exit Sub
    End If
    LoadGrid2
    XB.QuickSort XB.LowerBound(1), XB.UpperBound(1), 1, 0, XTYPE_STRING
    G2.Array = XB
    G2.ReBind
    G2.Refresh
    lblCount2.Caption = CStr(rs2.RecordCount) & " records"
    Screen.MousePointer = vbDefault

End Sub


Private Sub cmdFetch_Click()
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    If opttoday = True Then
        If Me.optByDocDate = True Then
            rs.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
                & " WHERE ((TR_TYPE IN (4,8) AND TR_STATUS IN (3,4)) OR (TR_TYPE in (3,11) AND TR_STATUS =4)) AND TR_CAPTUREDATE >= dbo.startOfDay(Getdate())", oPC.CO
        Else
            rs.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
                & " WHERE ((TR_TYPE IN (4,8) AND TR_STATUS IN (3,4)) OR (TR_TYPE in (3,11) AND TR_STATUS =4)) AND TR_ProcessingDATE >= dbo.startOfDay(Getdate())", oPC.CO
        End If
    Else
    If Me.optYesterday = True Then
        If Me.optByDocDate = True Then
            rs.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
                & " WHERE ((TR_TYPE IN (4,8) AND TR_STATUS IN (3,4)) OR (TR_TYPE in (3,11) AND TR_STATUS =4))  AND  TR_CAPTUREDATE >= dbo.startOfDay(DATEADD(d,-1,Getdate()))", oPC.CO
        Else
            rs.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
                & " WHERE ((TR_TYPE IN (4,8) AND TR_STATUS IN (3,4)) OR (TR_TYPE in (3,11) AND TR_STATUS =4)) AND  TR_ProcessingDATE >= dbo.startOfDay(DATEADD(d,-1,Getdate()))", oPC.CO
        End If
    Else
    If Me.opt7 = True Then
        If Me.optByDocDate = True Then
            rs.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
                & " WHERE ((TR_TYPE IN (4,8) AND TR_STATUS IN (3,4)) OR (TR_TYPE in (3,11) AND TR_STATUS =4)) AND  TR_CAPTUREDATE >= dbo.startOfDay(DATEADD(d,-7,Getdate()))", oPC.CO
        Else
            rs.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
                & " WHERE ((TR_TYPE IN (4,8) AND TR_STATUS IN (3,4)) OR (TR_TYPE in (3,11) AND TR_STATUS =4))  AND  TR_ProcessingDATE >= dbo.startOfDay(DATEADD(d,-7,Getdate()))", oPC.CO
        End If
    Else
    If opt14 = True Then
        If Me.optByDocDate = True Then
            rs.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
                & " WHERE ((TR_TYPE IN (4,8) AND TR_STATUS IN (3,4)) OR (TR_TYPE in (3,11) AND TR_STATUS =4)) AND  TR_CAPTUREDATE >= dbo.startOfDay(DATEADD(d,-14,Getdate()))", oPC.CO
        Else
            rs.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
                & " WHERE ((TR_TYPE IN (4,8) AND TR_STATUS IN (3,4)) OR (TR_TYPE in (3,11) AND TR_STATUS =4)) AND  TR_ProcessingDATE >= dbo.startOfDay(DATEADD(d,-14,Getdate()))", oPC.CO
        End If
    Else
        If Me.optByDocDate = True Then
            rs.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
                & " WHERE ((TR_TYPE IN (4,8) AND TR_STATUS IN (3,4)) OR (TR_TYPE in (3,11) AND TR_STATUS =4)) AND  TR_CAPTUREDATE >=dbo.startofDay('" & ReverseDate(Me.dtSince) & "')", oPC.CO
        Else
            rs.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
                & " WHERE ((TR_TYPE IN (4,8) AND TR_STATUS IN (3,4)) OR (TR_TYPE in (3,11) AND TR_STATUS =4)) AND  TR_ProcessingDATE >=dbo.startofDay('" & ReverseDate(Me.dtSince) & "')", oPC.CO
        End If
    End If
    End If
    End If
    End If
    If rs.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "No records found", vbInformation, "Status"
        Exit Sub
    End If
    LoadGrid
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 1, 0, XTYPE_STRING
    G.Array = XA
    G.ReBind
    G.Refresh
    lblCount.Caption = CStr(rs.RecordCount) & " records"
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdUnsent_Click()
    Screen.MousePointer = vbHourglass

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
        rs.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
            & " WHERE ((TR_TYPE IN (4,8) AND TR_STATUS IN (3,4)) OR (TR_TYPE in (3,11) AND TR_STATUS =4))  AND ISNULL(TR_DATETOPASTEL,0) < '2000-01-01'", oPC.CO

    If rs.EOF Then
        Screen.MousePointer = vbDefault
        MsgBox "No records found", vbInformation, "Status"
        Exit Sub
    End If
    
    LoadGrid
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 1, 0, XTYPE_STRING
    G.Array = XA
    G.ReBind
    G.Refresh
    lblCount = CStr(rs.RecordCount) & " records"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_Click()
    mnuSaveLayout
End Sub

Private Sub Form_Load()
Dim i As Integer

    SSTab1.Tab = 0
    For i = 1 To G.Columns.Count
        G.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "A", CStr(i), G.Columns(i - 1).Width)
    Next
    For i = 1 To G2.Columns.Count
        G2.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "B", CStr(i), G2.Columns(i - 1).Width)
    Next
    
End Sub

Private Sub G_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    G.Refresh
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAccountingExportManagement.G_HeadClick(ColIndex)", ColIndex
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    Select Case ColIndex
        Case 1
            GetRowType = XTYPE_STRING
        Case 2, 3, 4
            GetRowType = XTYPE_DATE
    End Select
End Function
Private Sub LoadGrid()
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
    lngArrayRows = rs.RecordCount
    XA.ReDim 1, lngArrayRows, 1, 10
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
                XA.Value(lngIndex, 1) = FNS(rs.Fields("TR_CODE"))
                XA.Value(lngIndex, 2) = TranslateDocType(FNS(rs.Fields("TR_TYPE")))
                XA.Value(lngIndex, 3) = FNS(rs.Fields("TP_ACNO")) & " - " & FNS(rs.Fields("TP_Name"))
                XA.Value(lngIndex, 4) = FNS(rs.Fields("TR_CAPTUREDATE"))
                XA.Value(lngIndex, 5) = FNS(rs.Fields("TR_PROCESSINGDATE"))
                XA.Value(lngIndex, 6) = IIf(FND(rs.Fields("TR_DATETOPASTEL")) < "2000-01-01", "", FNS(rs.Fields("TR_DATETOPASTEL")))
                XA.Value(lngIndex, 7) = FNS("Send again")
                XA.Value(lngIndex, 10) = FNN(rs.Fields("TR_ID"))
                lngIndex = lngIndex + 1
                rs.MoveNext
        Loop
    End If
   ' XA.QuickSort 1, lngArrayRows, 1, XORDER_ASCEND, XTYPE_STRING, 5, XORDER_ASCEND, XTYPE_DATE, 3, XORDER_ASCEND, XTYPE_STRING
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAccountingExportManagement.LoadGrid"
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
    If Not rs2.EOF Then
        rs2.MoveFirst
        Do While Not rs2.EOF
                XB.Value(lngIndex, 1) = Format(dteSelected, "dd-mm-yyyy")
                XB.Value(lngIndex, 2) = FNS(rs2.Fields("TR_CODE"))
                XB.Value(lngIndex, 3) = TranslateDocType(FNS(rs2.Fields("TR_TYPE")))
                XB.Value(lngIndex, 4) = FNS(rs2.Fields("TP_ACNO")) & " - " & FNS(rs2.Fields("TP_Name"))
                XB.Value(lngIndex, 5) = FNS(rs2.Fields("TR_CAPTUREDATE"))
                XB.Value(lngIndex, 6) = FNS(rs2.Fields("TR_PROCESSINGDATE"))
                XB.Value(lngIndex, 7) = IIf(FND(rs2.Fields("TR_DATETOPASTEL")) < "2000-01-01", "", FNS(rs2.Fields("TR_DATETOPASTEL")))
                XB.Value(lngIndex, 10) = FNN(rs2.Fields("TR_ID"))
                lngIndex = lngIndex + 1
                rs2.MoveNext
        Loop
    End If
   ' XA.QuickSort 1, lngArrayRows, 1, XORDER_ASCEND, XTYPE_STRING, 5, XORDER_ASCEND, XTYPE_DATE, 3, XORDER_ASCEND, XTYPE_STRING
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAccountingExportManagement.LoadGrid2"
End Sub

Private Function TranslateDocType(i As Integer) As String
    Select Case i
        Case 3
            TranslateDocType = "INV"
        Case 8
            TranslateDocType = "CN"
        Case 4
            TranslateDocType = "PUR"
        Case 11
            TranslateDocType = "RET"
    End Select
End Function
Private Sub dtSince_GotFocus()
    Me.opt14 = False
    Me.opt7 = False
    Me.optYesterday = False
    Me.opttoday = False
End Sub


Private Sub optall_Click()
    If optALL Then
        rs.Filter = ""
        LoadGrid
        G.Array = XA
        G.ReBind
        lblCount = CStr(rs.RecordCount) & " records"
    End If
    
End Sub

Private Sub optRET_Click()
    If optRET Then
        rs.Filter = "TR_TYPE = 11"
        'rs.Requery
        LoadGrid
        G.Array = XA
        G.ReBind
        lblCount = CStr(rs.RecordCount) & " records"
    End If
End Sub

Private Sub optSI_Click()
    If optSI Then
        rs.Filter = "TR_TYPE = 4"
        rs.Requery
        LoadGrid
        G.Array = XA
        G.ReBind
        lblCount = CStr(rs.RecordCount) & " records"
    End If
End Sub
Private Sub optINV_Click()
    If optINV Then
        rs.Filter = "TR_TYPE = 3"
        'rs.Requery
        LoadGrid
        G.Array = XA
        G.ReBind
        lblCount = CStr(rs.RecordCount) & " records"
    End If
End Sub

Private Sub optCN_Click()
    If optCN Then
        rs.Filter = "TR_TYPE = 8"
        rs.Requery
        LoadGrid
        G.Array = XA
        G.ReBind
        lblCount = CStr(rs.RecordCount) & " records"
    End If
End Sub
Private Sub G_ButtonClick(ByVal ColIndex As Integer)
Dim i As Integer
    i = ColIndex + 1
    If i = 7 Then   'checkbox
        oPC.CO.Execute "UPDATE tTR SET TR_DATETOPASTEL = NULL WHERE TR_ID =  " & XA(G.Bookmark, 10)
        'cmdFetch_Click
        XA(G.Bookmark, 6) = ""
        G.RefetchRow
    End If
End Sub
Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayout Me.G, Me.Name & "A"
    SaveLayout Me.G2, Me.Name & "B"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
    HandleError
End Sub

