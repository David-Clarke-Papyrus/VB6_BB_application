VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Begin VB.Form frmReminderstoProcess 
   Caption         =   "Reminders and cancellations to view"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
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
      Left            =   7395
      Picture         =   "frmReminderstoProcess.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1080
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Height          =   2760
      Left            =   240
      OleObjectBlob   =   "frmReminderstoProcess.frx":038A
      TabIndex        =   0
      Top             =   285
      Width           =   8955
   End
End
Attribute VB_Name = "frmReminderstoProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim x As XArrayDB
Dim oSQL As New z_SQL
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
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

Public Sub component(pPAID As Long)
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "SELECT * FROM vRemindersToProcess_b WHERE PAID = " & CStr(pPAID) & " ORDER BY SupplierName"
    cmd.commandType = adCmdText
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open cmd, , adOpenStatic
    rs.ActiveConnection = Nothing
     If rs.eof Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
   
    Set cmd = Nothing

End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    If rs Is Nothing Then Exit Sub
    If rs.eof Then Exit Sub
    Set x = New XArrayDB
    x.Clear
    x.ReDim 1, rs.RecordCount, 1, 10
    For lngIndex = 1 To rs.RecordCount
        x.Value(lngIndex, 1) = FNS(rs.fields("Suppliername")) & " (" & FNS(rs.fields("SupplierACNo")) & ")"
        x.Value(lngIndex, 2) = IIf(FNB(rs.fields("POLA_IsCancelled")), "C", "R")
        x.Value(lngIndex, 3) = FNS(rs.fields("StaffMember"))
        x.Value(lngIndex, 4) = FNS(rs.fields("Email"))
        x.Value(lngIndex, 7) = FNB(rs.fields("POLA_IsCancelled"))
        x.Value(lngIndex, 8) = FNB(rs.fields("POLA_NeedReminder"))
        x.Value(lngIndex, 9) = FNN(rs.fields("PAID"))
        x.Value(lngIndex, 10) = FNN(rs.fields("TP_ID"))
        rs.MoveNext
    Next
    x.QuickSort 1, rs.RecordCount, 1, XORDER_ASCEND, XTYPE_STRING
    G.Array = x
    G.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActions.LoadGrid"
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Activate()
SetMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveLayout G, Me.Name, Me.Height, Me.Width

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

    If ColIndex = 4 Then
        Screen.MousePointer = vbHourglass
        Set oSM = New z_StockManager
        bHasDocuments = oSM.PrintAReminderorCancellation(FNN(x.Value(G.Bookmark, 10)), FNN(x.Value(G.Bookmark, 9)), IIf(FNB(x.Value(G.Bookmark, 7)), "C", "R"))
        Screen.MousePointer = vbDefault
        If bHasDocuments Then MsgBox "PDF file created.", vbInformation, "Done"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActions.G_ButtonClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    SetGridLayout Me.G, Me.Name
    SetFormSize Me
    LoadGrid
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

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.G, Me.Name, Me.Height, Me.Width
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActions.mnuSaveLayout"
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    If Me.Width < 5800 Then Exit Sub
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

