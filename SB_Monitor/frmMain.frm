VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Service broker monitor"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Connect to . . ."
      Height          =   885
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   3660
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   450
         Left            =   2580
         TabIndex        =   2
         Top             =   255
         Width           =   915
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1065
         TabIndex        =   1
         Top             =   300
         Width           =   1395
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2805
      Left            =   105
      TabIndex        =   13
      Top             =   1605
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4948
      SortKey         =   1
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "EAN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date added"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Supplier"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.Label lblLastLoyaltymasterReceived 
      Caption         =   "Unknown"
      Height          =   255
      Left            =   7140
      TabIndex        =   12
      Top             =   1275
      Width           =   1665
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Last loyalty received from Central:"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   1275
      Width           =   2385
   End
   Begin VB.Label lblLastLoyaltySent 
      Caption         =   "Unknown"
      Height          =   255
      Left            =   7140
      TabIndex        =   10
      Top             =   990
      Width           =   1665
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Last loyalty sent:"
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   990
      Width           =   2385
   End
   Begin VB.Label lblLastHubSent 
      Caption         =   "Unknown"
      Height          =   255
      Left            =   7140
      TabIndex        =   8
      Top             =   690
      Width           =   1665
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Hub sent:"
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   675
      Width           =   2385
   End
   Begin VB.Label lblLastSalesSent 
      Caption         =   "Unknown"
      Height          =   255
      Left            =   7140
      TabIndex        =   6
      Top             =   390
      Width           =   1665
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Last sales sent:"
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   390
      Width           =   2385
   End
   Begin VB.Label lblLastTimer 
      Caption         =   "Unknown"
      Height          =   255
      Left            =   7140
      TabIndex        =   4
      Top             =   90
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Last timer:"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   90
      Width           =   2385
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset

Private Sub cmdConnect_Click()
    On Error GoTo errHandler
    ConnectToDB
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM _tTimerStat", oCnn.Connection, adOpenForwardOnly, adLockOptimistic
    
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.Open "SELECT TOP 150 * FROM _tSBLog Order By SBL_DATE DESC", oCnn.Connection, adOpenForwardOnly, adLockOptimistic
    
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdConnect_Click"
End Sub

Private Sub ConnectToDB()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    
    On Error Resume Next
       oCnn.Connection.Close
    On Error GoTo errHandler
    
    oCnn.SetBranchCode Me.txtIP
    oCnn.InitializeSettings
    oCnn.SetConnectionString
    oCnn.OpenDB
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ConnectToDB"
End Sub


Private Sub LoadControls()
    Me.lblLastLoyaltymasterReceived.Caption = Format(rs.Fields("T_LastLoyaltyUpdateFromCentral"), "DD-MM-YYYY Hh:Nn:Ss")
    Me.lblLastHubSent.Caption = Format(rs.Fields("T_LastHubTransmission"), "DD-MM-YYYY Hh:Nn:Ss")
    Me.lblLastLoyaltySent.Caption = Format(rs.Fields("T_LastLoyaltyTransmission"), "DD-MM-YYYY Hh:Nn:Ss")
    Me.lblLastSalesSent.Caption = Format(rs.Fields("T_LastSalesTransmission"), "DD-MM-YYYY Hh:Nn:Ss")
    Me.lblLastTimer.Caption = Format(rs.Fields("T_DATE"), "DD-MM-YYYY Hh:Nn:Ss")
    LoadListView
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oCnn.CloseDB
End Sub

Private Sub LoadListView()
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvw.ListItems.Clear
    For i = 1 To rs2.RecordCount
        Set objItm = Me.lvw.ListItems.Add
        With objItm
            .Text = Format(rs2.Fields("SBL_DATE"), "yyyy-mm-dd Hh:Nn")
            .SubItems(1) = rs2.Fields("SBL_Msg")
            .SubItems(2) = rs2.Fields("SBL_PROC")
        End With
        rs2.MoveNext
    Next i
End Sub

