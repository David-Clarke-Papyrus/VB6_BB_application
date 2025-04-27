VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Begin VB.Form frmTrackingActionsDetails 
   Caption         =   "Tracking actions details"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4005
   ScaleWidth      =   12870
   Begin VB.TextBox txtCustomerMessage 
      Height          =   435
      Left            =   7035
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   15
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.TextBox txtSupplierMessage 
      Height          =   435
      Left            =   1725
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   15
      Visible         =   0   'False
      Width           =   3525
   End
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
      Left            =   9675
      Picture         =   "frmTrackingActionsDetails.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3420
      Width           =   1080
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Height          =   2970
      Left            =   135
      OleObjectBlob   =   "frmTrackingActionsDetails.frx":038A
      TabIndex        =   0
      Top             =   360
      Width           =   12585
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer report"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   5715
      TabIndex        =   3
      Top             =   15
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier message"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   390
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1260
   End
End
Attribute VB_Name = "frmTrackingActionsDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim x As XArrayDB
Dim flgLoading As Boolean

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.G, Me.Name, Me.Height, Me.Width
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActionsDetails.mnuSaveLayout"
End Sub
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
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.SetMenu"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActionsDetails.SetMenu"
End Sub

Public Sub component(pRs As ADODB.Recordset, pSupplierMessage As String, pCustomerMessage As String)
    On Error GoTo errHandler
    Set rs = pRs
    If rs Is Nothing Then
        Exit Sub
    End If
    If Not rs.eof Then
        rs.MoveFirst
        LoadGrid
    End If
    Me.txtSupplierMessage = pSupplierMessage
    Me.txtCustomerMessage = pCustomerMessage
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActionsDetails.component(pRS,pSupplierMessage,pCustomerMessage)", Array(pRs, _
         pSupplierMessage, pCustomerMessage)
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set x = New XArrayDB
    x.Clear
    x.ReDim 1, rs.RecordCount, 1, 11
    For lngIndex = 1 To rs.RecordCount
        x.Value(lngIndex, 1) = FNS(rs.fields("Number"))
        x.Value(lngIndex, 2) = FNS(rs.fields("TPNAME"))
        x.Value(lngIndex, 3) = FNS(rs.fields("EAN"))
        x.Value(lngIndex, 4) = FNS(rs.fields("Description"))
        If FND(rs.fields("DiarizeDate")) > "2000-01-01" Then
            x.Value(lngIndex, 5) = FND(rs.fields("DiarizeDate"))
        Else
            x.Value(lngIndex, 5) = "n/a"
        End If
        x.Value(lngIndex, 6) = FNB(rs.fields("IsCancellation")) 'IIf(FNB(rs.Fields("IsCancellation")), 1, 0)
        If FNB(rs.fields("NeedAction")) = True Then
            If FNS(rs.fields("ActionType")) = "POL" Then
                x.Value(lngIndex, 7) = "Reminder"
            Else
                x.Value(lngIndex, 7) = "Cust rep."
            End If
        End If
      '  x.Value(lngIndex, 8) = oPC.configuration.ProductStatus.Item(FNS(rs.Fields("NewStatus")))
        x.Value(lngIndex, 8) = FNS(rs.fields("SUPPMSG"))
        x.Value(lngIndex, 11) = FNS(rs.fields("PID"))
        rs.MoveNext
    Next
    x.QuickSort 1, rs.RecordCount, 1, XORDER_DESCEND, XTYPE_DATE
    G.Array = x
    G.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActionsDetails.LoadGrid"
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActionsDetails.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    Me.Width = 11000
    Me.Height = 4400
    SetMenu
    SetGridLayout Me.G, Me.Name
    SetFormSize Me
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActionsDetails.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    G.Width = NonNegative_Lng(Me.Width - (G.Left + 310))
    lngDiff = G.Height
    G.Height = NonNegative_Lng(Me.Height - (G.TOP + 1250))
    lngDiff = (G.Height - lngDiff)
    cmdClose.TOP = NonNegative_Lng(cmdClose.TOP + lngDiff)
    
    Me.cmdClose.Left = G.Width - 950

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActionsDetails.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveLayout Me.G, Me.Name, Me.Height, Me.Width
End Sub

Private Sub G_ButtonClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
    If ColIndex = 3 Then
        MsgBox "Supplier reminder"
    ElseIf ColIndex = 4 Then
        MsgBox "Customer report"
    ElseIf ColIndex = 5 Then
        MsgBox "Review details"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActionsDetails.G_ButtonClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Function FormatType(pType As String) As String
    On Error GoTo errHandler
    Select Case pType
    Case "P"
        FormatType = "P.O. tracking"
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActionsDetails.FormatType(pType)", pType
End Function

Private Sub G_DblClick()
    On Error GoTo errHandler
Dim strPID As String
Dim frm As frmProductPrev
Dim oProd As a_Product

    strPID = x.Value(G.Bookmark, 11)
    If strPID > "" Then
        Set oProd = New a_Product
        oProd.Load strPID, 0
        Set frm = New frmProductPrev
        frm.component oProd
        frm.Show
    End If

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmTrackingActionsDetails.G_DblClick"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActionsDetails.G_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub G_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant
    If flgLoading Then Exit Sub

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        x.QuickSort x.LowerBound(1), x.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    G.Refresh
    Screen.MousePointer = vbDefault

    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActionsDetails.G_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 3, 4, 7, 8
            GetRowType = XTYPE_STRING
        Case 5
            GetRowType = XTYPE_DATE
        Case 6
            GetRowType = XTYPE_BOOLEAN
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTrackingActionsDetails.GetRowType(ColIndex)", ColIndex
End Function



