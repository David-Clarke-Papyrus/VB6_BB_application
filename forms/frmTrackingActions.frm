VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Begin VB.Form frmTrackingActions 
   Caption         =   "Tracking actions"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4005
   ScaleWidth      =   11385
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8340
      Picture         =   "frmTrackingActions.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2925
      Width           =   1080
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Height          =   2760
      Left            =   90
      OleObjectBlob   =   "frmTrackingActions.frx":038A
      TabIndex        =   0
      Top             =   90
      Width           =   11025
   End
End
Attribute VB_Name = "frmTrackingActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim x As XArrayDB
Dim oSQL As New z_SQL

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = False
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuCancelINactive.Enabled = False
    Forms(0).mnuFulfil.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
    'Forms(0).mnuInvAdd.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.SetMenu"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActions.SetMenu"
End Sub
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.G, Me.Name, Me.Height, Me.Width
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActions.mnuSaveLayout"
End Sub

Public Sub component(PXMLArgs As String, CXMLArgs As String)
    On Error GoTo errHandler
    Set rs = oSQL.GetTrackingActions(PXMLArgs, CXMLArgs, 300)
    If rs Is Nothing Then
        Exit Sub
    End If
    If Not rs.eof Then
        rs.MoveFirst
    End If
        LoadGrid
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActions.component(PXMLArgs,CXMLArgs)", Array(PXMLArgs, CXMLArgs)
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set x = New XArrayDB
    x.Clear
    x.ReDim 1, rs.RecordCount, 1, 10
    For lngIndex = 1 To rs.RecordCount
        x.Value(lngIndex, 1) = Format(rs.fields("DTE"), "dd/mm/yyyy Hh:Nn AM/PM")
        x.Value(lngIndex, 2) = FormatType(FNS(rs.fields("TYP")))
        x.Value(lngIndex, 3) = FNS(rs.fields("WORKSTATION"))
        x.Value(lngIndex, 10) = FNS(rs.fields("PA_ID"))
        rs.MoveNext
    Next
    x.QuickSort 1, rs.RecordCount, 1, XORDER_DESCEND, XTYPE_DATE
    G.Array = x
    G.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActions.LoadGrid"
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActions.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Left = 500
    TOP = 900
    Me.Width = 11000
    Me.Height = 4300
    SetMenu
    SetGridLayout Me.G, Me.Name
    SetFormSize Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActions.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    SaveLayout G, Me.Name, Me.Height, Me.Width
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActions.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub G_ButtonClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim frm As frmTrackingActionsDetails
Dim oSM As z_StockManager
Dim strSM As String
Dim strCM As String
Dim strSignature As String
Dim bHasDocuments As Boolean
Dim frem As frmReminderstoProcess

    If ColIndex = 3 Then
        Set frem = New frmReminderstoProcess
        frem.component FNN(x.Value(G.Bookmark, 10))
        frem.Show vbModal
'        Screen.MousePointer = vbHourglass
'        Set oSM = New z_StockManager
'        p 1, "Before generateReminders"
'        bHasDocuments = oSM.GenerateReminders(FNN(x.Value(G.Bookmark, 10)))
'        p 2, "After generateReminders"
'        bHasDocuments = bHasDocuments Or oSM.GenerateCancellationReports(FNN(x.Value(G.Bookmark, 10)))
'        Screen.MousePointer = vbDefault
'        If bHasDocuments Then MsgBox "You will find PDF reminders in preview mode to be printed or emailed in your taskbar.", vbInformation, "Done"
    ElseIf ColIndex = 4 Then
        Screen.MousePointer = vbHourglass
        Set oSM = New z_StockManager
        bHasDocuments = oSM.GenerateBOStatusReports(FNN(x.Value(G.Bookmark, 10)))
        Screen.MousePointer = vbDefault
        If bHasDocuments Then MsgBox "You will find PDF Customer reports in preview mode to be printed or emailed in your taskbar.", vbInformation, "Done"
    ElseIf ColIndex = 5 Then
        Screen.MousePointer = vbHourglass
        Set oSM = New z_StockManager
        bHasDocuments = oSM.GenerateBOStatusReports(FNN(x.Value(G.Bookmark, 10)))
        Screen.MousePointer = vbDefault
        If bHasDocuments Then MsgBox "Transmission is complete.", vbInformation, "Done"
    ElseIf ColIndex = 6 Then
        Screen.MousePointer = vbHourglass
        Set frm = New frmTrackingActionsDetails
        Set rs = oSQL.GetTrackingActionsDetails(FNN(x.Value(G.Bookmark, 10)), strSM, strCM, strSignature)
        Screen.MousePointer = vbDefault
        frm.component rs, strSM, strCM
        frm.Show
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmTrackingActions.G_ButtonClick(ColIndex)", ColIndex, , , "strErrPos", Array(strErrPos)
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActions.G_ButtonClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Function FormatType(pType As String) As String
    On Error GoTo errHandler
    Select Case pType
    Case "P"
        FormatType = "P.O. tracking"
    Case "I"
        FormatType = "Supplier report"
    Case "C"
        FormatType = "C.O. tracking"
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActions.FormatType(pType)", pType
End Function

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    If Me.Width < 9800 Then Exit Sub
    G.Width = NonNegative_Lng(Me.Width - (G.Left + 300))
    lngDiff = G.Height
    G.Height = NonNegative_Lng(Me.Height - (G.TOP + 1250))
    lngDiff = (G.Height - lngDiff)
    cmdClose.TOP = cmdClose.TOP + lngDiff
    
    Me.cmdClose.Left = G.Width - 960

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActions.Form_Resize", , EA_NORERAISE
    HandleError
End Sub



