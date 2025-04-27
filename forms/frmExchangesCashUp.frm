VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmExchangesCashUP 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Cash-up"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   5490
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Print daily summary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2790
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3630
      Width           =   1995
   End
   Begin VB.CommandButton cmdCashup 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Cash up"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2865
      Width           =   1830
   End
   Begin VB.CommandButton cmdZSession 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Print &Z ession"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2790
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2850
      Width           =   1980
   End
   Begin TrueOleDBGrid60.TDBGrid GZ 
      Height          =   2295
      Left            =   150
      OleObjectBlob   =   "frmExchangesCashUp.frx":0000
      TabIndex        =   0
      Top             =   435
      Width           =   4605
   End
   Begin TrueOleDBGrid60.TDBGrid GO 
      Height          =   2295
      Left            =   180
      OleObjectBlob   =   "frmExchangesCashUp.frx":460F
      TabIndex        =   2
      Top             =   5715
      Width           =   4545
   End
   Begin VB.Label lblTimer 
      BackStyle       =   0  'Transparent
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
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   7245
      Width           =   3405
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Operator sessions"
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
      Height          =   315
      Left            =   285
      TabIndex        =   3
      Top             =   5430
      Width           =   2640
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Z sessions"
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
      Height          =   315
      Left            =   270
      TabIndex        =   1
      Top             =   135
      Width           =   2640
   End
End
Attribute VB_Name = "frmExchangesCashUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XE As XArrayDB
Dim XO As XArrayDB
Dim XZ As XArrayDB
Dim XCSL As XArrayDB
Dim XPAY As XArrayDB
Dim rs As ADODB.Recordset
Dim rsZ As ADODB.Recordset
Dim OPSID As Variant
Dim ocZ As c_ZSession
Dim ocCS As c_CSs
Dim ocEX As c_Exchanges

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub G1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuReserveList   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.G1_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, X, Y), EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo errHandler
    LoadZSessions
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.cmdRefresh_Click"
End Sub

Private Sub cmdZSession_Click()
Dim ar As arZSession

End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    top = 310
    left = 120
    Width = 6600
    Height = 8000
    LoadZSessions
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadZGrid()
    On Error GoTo errHandler
Dim objItem As d_ZSession
Dim itmList As ListItem
Dim lngIndex As Long
Dim lngArrayRows As Long
Dim i As Integer
    Set XZ = New XArrayDB
    XZ.Clear
    XZ.ReDim 1, ocZ.Count, 1, 8
    For i = 1 To ocZ.Count
        XZ.Value(i, 1) = ocZ.Item(i).NominalDateF
        XZ.Value(i, 2) = ocZ.Item(i).TillPoint
        XZ.Value(i, 3) = ocZ.Item(i).supervisorName
        XZ.Value(i, 4) = ocZ.Item(i).StartDateF
        XZ.Value(i, 5) = ocZ.Item(i).EndDateF
        XZ.Value(i, 6) = ocZ.Item(i).ID
        XZ.Value(i, 7) = ocZ.Item(i).StartDateSort
        XZ.Value(i, 8) = ocZ.Item(i).EndDate
    Next
    XZ.QuickSort 1, XZ.UpperBound(1), 7, XORDER_DESCEND, XTYPE_STRING
    GZ.Array = XZ
    GZ.ReBind
'    GZ.Bookmark = 0
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.LoadZGrid"
End Sub

Private Sub LoadOpsGrid()
    On Error GoTo errHandler
Dim objItem As d_CS
Dim itmList As ListItem
Dim lngIndex As Long
Dim lngArrayRows As Long
Dim i As Integer
    Set XO = New XArrayDB
    XO.Clear
    XO.ReDim 1, ocCS.Count, 1, 6
    For i = 1 To ocCS.Count
        XO.Value(i, 1) = ocCS.Item(i).StaffName
        XO.Value(i, 2) = ocCS.Item(i).StartDateF
        XO.Value(i, 3) = ocCS.Item(i).EndDateF
        XO.Value(i, 4) = ocCS.Item(i).TRID
        XO.Value(i, 5) = ocCS.Item(i).StartDateSort
        XO.Value(i, 6) = ocCS.Item(i).CSGUID
    Next
    XO.QuickSort 1, XO.UpperBound(1), 5, XORDER_DESCEND, XTYPE_STRING
    GO.Array = XO
    GO.ReBind
    GO.Bookmark = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.LoadOpsGrid"
End Sub

Private Sub ClearOpsGrid()
    On Error GoTo errHandler
    If Not XO Is Nothing Then
        XO.Clear
        XO.ReDim 0, 0, 1, 6
    End If
    GO.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.ClearOpsGrid"
End Sub


Private Sub GZ_SelChange(Cancel As Integer)
    If XZ(GZ.Bookmark, 8) = CDate(0) Then ' the session has not ended
        TimerON True
    Else
        TimerON False
        RefreshOps
        RefreshExchanges
        RefreshDetails
    End If
End Sub


Private Sub GO_SelChange(Cancel As Integer)
    TimerON False
    RefreshExchanges
    RefreshDetails
End Sub
Private Sub LoadZSessions()
    On Error GoTo errHandler
    Set ocZ = New c_ZSession
    ocZ.Load DateAdd("yyyy", -1, Date)
    Screen.MousePointer = vbHourglass
    LoadZGrid
    RefreshOps
    RefreshExchanges
    RefreshDetails
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.LoadZSessions"
End Sub
Private Sub RefreshOps()
    On Error GoTo errHandler
    If Not ocCS Is Nothing Then ClearOpsGrid
    Set ocCS = New c_CSs
    If Not XZ(GZ.Bookmark, 6) = Empty Then
        ocCS.LoadByZID XZ(GZ.Bookmark, 6)
        LoadOpsGrid
    Else
        ClearOpsGrid
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.RefreshOps()", , EA_NORERAISE
    HandleError
End Sub
