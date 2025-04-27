VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmExchanges 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Front desk activity"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   14430
   Begin VB.CommandButton cmdGet 
      BackColor       =   &H00C4BCA4&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3120
      Picture         =   "frmExchanges.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   -15
      Width           =   765
   End
   Begin VB.TextBox txtArg1 
      Height          =   285
      Left            =   1935
      TabIndex        =   0
      Top             =   75
      Width           =   1035
   End
   Begin VB.CommandButton cmdFix 
      BackColor       =   &H00C4BCA4&
      Caption         =   "CreateInvoice(Davidonly)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   13710
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1890
      Width           =   1065
   End
   Begin VB.CommandButton cmdRefreshZ 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Refresh day sessions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13005
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3030
      Width           =   2535
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Refresh exchanges"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5310
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3285
      Width           =   1815
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
      Left            =   13185
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6420
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Frame frm1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Exchange details"
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
      Height          =   3255
      Left            =   4140
      TabIndex        =   9
      Top             =   3750
      Width           =   8685
      Begin VB.CommandButton cmdPrintSale 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Print &exchange"
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
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2475
         Width           =   1830
      End
      Begin TrueOleDBGrid60.TDBGrid GPAY 
         Height          =   1140
         Left            =   240
         OleObjectBlob   =   "frmExchanges.frx":038A
         TabIndex        =   10
         Top             =   1950
         Width           =   6030
      End
      Begin TrueOleDBGrid60.TDBGrid GCSL 
         Height          =   1470
         Left            =   225
         OleObjectBlob   =   "frmExchanges.frx":460F
         TabIndex        =   14
         Top             =   345
         Width           =   8205
      End
   End
   Begin VB.CommandButton cmdPrintList 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print exchanges list"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10740
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3315
      Width           =   2100
   End
   Begin TrueOleDBGrid60.TDBGrid GE 
      Height          =   2970
      Left            =   5310
      OleObjectBlob   =   "frmExchanges.frx":A4B8
      TabIndex        =   2
      Top             =   300
      Width           =   7515
   End
   Begin TrueOleDBGrid60.TDBGrid GZ 
      Height          =   2235
      Left            =   15
      OleObjectBlob   =   "frmExchanges.frx":10373
      TabIndex        =   3
      Top             =   780
      Width           =   5160
   End
   Begin TrueOleDBGrid60.TDBGrid GO 
      Height          =   1245
      Left            =   30
      OleObjectBlob   =   "frmExchanges.frx":156F6
      TabIndex        =   7
      Top             =   3390
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Exchange no. or date"
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
      Height          =   420
      Left            =   30
      TabIndex        =   18
      Top             =   75
      Width           =   1950
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Notes"
      ForeColor       =   &H8000000D&
      Height          =   2175
      Left            =   30
      TabIndex        =   17
      Top             =   4800
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Operator sessions"
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
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   3150
      Width           =   2640
   End
   Begin VB.Label lblExchanges 
      BackStyle       =   0  'Transparent
      Caption         =   "Exchanges"
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
      Height          =   255
      Left            =   5310
      TabIndex        =   6
      Top             =   60
      Width           =   7380
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Day sessions"
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
      Height          =   315
      Left            =   45
      TabIndex        =   5
      Top             =   525
      Width           =   1185
   End
End
Attribute VB_Name = "frmExchanges"
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
Dim flgLoading As Boolean
Dim GESwitch As Boolean
Const strNotes As String = "Notes:" & vbCrLf & "1. Enter an exchange number or a date or leave blank for recent day-sessions and click ther tick button." & vbCrLf _
        & "2. Select the day-session you wish to examine." & vbCrLf _
        & "then . . ." & vbCrLf _
        & "3. Select the operator-session of the selected day-session." & vbCrLf _
        & "Make selections by clicking on the grey margin of the day-session and operator-session grids respectively. The mouse pointer will show a right-arrow while the data is being fetched."

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



Private Sub cmdFix_Click()
Dim oSM As New z_StockManager

    oSM.CreateInvoiceFromExchange XE(GE.Bookmark, 9)
End Sub

Private Sub cmdGet_Click()
    cmdRefreshZ_Click
End Sub

Private Sub cmdPrintSale_Click()
Dim ar As New arExchange


    If IsNull(GE.Bookmark) Then Exit Sub
    
    If XE(GE.Bookmark, 9) = Empty Then Exit Sub

    ar.Component XE(GE.Bookmark, 1), XZ(GZ.Bookmark, 2), ocEX.Item(GE.Bookmark).ExchangeDate2F, XE(GE.Bookmark, 3), ocEX(XE(GE.Bookmark, 9)).CSLS, ocEX(XE(GE.Bookmark, 9)).PAYS, XE(GE.Bookmark, 6), XE(GE.Bookmark, 7), XE(GE.Bookmark, 8), ocEX(XE(GE.Bookmark, 9)).voided
    ar.Show vbModal
    
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo errHandler
        Screen.MousePointer = vbHourglass
        
'
'    RefreshOps
'    RefreshExchanges
'    RefreshDetails
        
        
        DoEvents
        RefreshExchanges
        RefreshDetails
        
        Screen.MousePointer = vbDefault
Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.cmdRefresh_Click"
End Sub


Private Sub cmdPrintList_Click()
Dim ar As New arExchanges
ar.Printer.Orientation = ddOLandscape
    ar.Component XE, "Exchanges for X Session started: " & Format(XO.Value(GO.Bookmark, 2), "dd/mm/yyyy HH:NN AMPM")
    ar.Show
End Sub

Private Sub cmdRefreshZ_Click()
    LoadZSessions
End Sub

Private Sub cmdZSession_Click()
Dim ar As arZSession

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    top = 35
    left = 30
    Width = 13000
    Height = 7500
    Me.lblNotes.Caption = strNotes
    Screen.MousePointer = vbHourglass
    Me.cmdRefresh.Visible = True
 
    Set XZ = New XArrayDB
    XZ.ReDim 1, 1, 1, 8
    Set GZ.Array = XZ
    
    Set XO = New XArrayDB
    XO.ReDim 1, 1, 1, 6
    Set GO.Array = XO
    
    Set XE = New XArrayDB
    XE.ReDim 1, 1, 1, 11
    Set GE.Array = XE
    
    Set XCSL = New XArrayDB
    XCSL.ReDim 1, 1, 1, 7
    Set GCSL.Array = XCSL
    
    Set XPAY = New XArrayDB
    XPAY.ReDim 1, 1, 1, 3
    Set GPAY.Array = XPAY
    
   ' LoadZSessions
    
    flgLoading = False
        
    Screen.MousePointer = vbDefault
    
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
   ' Set XZ = New XArrayDB
    XZ.Clear
    XZ.ReDim 1, ocZ.Count, 1, 8
    For i = 1 To ocZ.Count
        XZ.Value(i, 1) = ocZ.Item(i).NominalDateF
        XZ.Value(i, 2) = ocZ.Item(i).TillPoint
        XZ.Value(i, 3) = ocZ.Item(i).SupervisorName
        XZ.Value(i, 4) = ocZ.Item(i).StartDateF
        XZ.Value(i, 5) = ocZ.Item(i).EndDateF
        XZ.Value(i, 6) = ocZ.Item(i).ID
        XZ.Value(i, 7) = ocZ.Item(i).StartDateSort
        XZ.Value(i, 8) = ocZ.Item(i).EndDate
    Next
    XZ.QuickSort 1, XZ.UpperBound(1), 7, XORDER_DESCEND, XTYPE_STRING
    'GZ.Array = XZ
    GZ.ReBind
    GZ.Bookmark = 0
    
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
'    Set XO = New XArrayDB
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
  '  GO.Array = XO
    GO.ReBind
    GO.Bookmark = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.LoadOpsGrid"
End Sub

Private Sub LoadExGrid()
    On Error GoTo errHandler
Dim objItem As d_Exchange
Dim itmList As ListItem
Dim lngIndex As Long
Dim lngArrayRows As Long
Dim i As Integer
  '  Set XE = New XArrayDB
    XE.Clear
    XE.ReDim 1, ocEX.Count, 1, 11
    For i = 1 To ocEX.Count
        XE.Value(i, 1) = ocEX.Item(i).ExchangeNumber
        XE.Value(i, 2) = ocEX.Item(i).ExchangeDateF
        XE.Value(i, 3) = ocEX.Item(i).SalesPersonName
        XE.Value(i, 4) = ocEX.Item(i).TotalPayableF
        XE.Value(i, 5) = ocEX.Item(i).TotalVATF
        XE.Value(i, 6) = ocEX.Item(i).ChangeGivenF
        XE.Value(i, 7) = ocEX.Item(i).ExchangeTypeF & IIf(ocEX.Item(i).voided = True, "(Voided)", "")
        If ocEX.Item(i).VOIDS > 0 Then
            XE.Value(i, 8) = "Voids - " & CStr(ocEX.Item(i).VOIDS)
        Else
            XE.Value(i, 8) = ocEX.Item(i).Note
        End If
        XE.Value(i, 9) = ocEX.Item(i).ID
        XE.Value(i, 10) = ocEX.Item(i).ExchangeDateSort
        XE.Value(i, 11) = ocEX.Item(i).voided
    Next
    XE.QuickSort 1, XE.UpperBound(1), 1, XORDER_DESCEND, XTYPE_DATE
  '  GE.Array = XE
    On Error Resume Next
    GE.ReBind
    GE.Bookmark = 0 'GE.FirstRow
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.LoadExGrid"
End Sub

Private Sub LoadCSLGrid(cCSL As Collection)
    On Error GoTo errHandler
Dim objItem As d_CSL
Dim itmList As ListItem
Dim lngIndex As Long
Dim lngArrayRows As Long
Dim i As Integer
'    Set XCSL = New XArrayDB
    XCSL.Clear
    XCSL.ReDim 1, cCSL.Count, 1, 7
    For i = 1 To cCSL.Count
        XCSL.Value(i, 1) = cCSL.Item(i).CodeF
        XCSL.Value(i, 2) = cCSL.Item(i).Title
        XCSL.Value(i, 3) = cCSL.Item(i).Qty
        XCSL.Value(i, 4) = cCSL.Item(i).PriceF
        XCSL.Value(i, 5) = cCSL.Item(i).DiscountCombinedF
        XCSL.Value(i, 6) = cCSL.Item(i).ExtensionF
        XCSL.Value(i, 7) = cCSL.Item(i).VATCombinedF
        
    Next
    XCSL.QuickSort 1, XCSL.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    Set GCSL.Array = XCSL
   ' On Error Resume Next
'GCSL.Refresh
    GCSL.ReBind
  ' GCSL.Refresh
    GCSL.Bookmark = 0 'GCSL.FirstRow
   GCSL.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.LoadCSLGrid"
End Sub
Private Sub LoadPAYGrid(cPAY As Collection)
    On Error GoTo errHandler
Dim objItem As d_Payment
Dim itmList As ListItem
Dim lngIndex As Long
Dim lngArrayRows As Long
Dim i As Integer
  '  Set XPAY = New XArrayDB
    XPAY.Clear
    XPAY.ReDim 1, cPAY.Count, 1, 3
    For i = 1 To cPAY.Count
        XPAY.Value(i, 1) = cPAY.Item(i).PaymentTypeF
        XPAY.Value(i, 2) = cPAY.Item(i).AmtF
        XPAY.Value(i, 3) = cPAY.Item(i).Reference
    Next
 '   XPAY.QuickSort 1, XPAY.UpperBound(1), 5, XORDER_DESCEND, XTYPE_STRING
  '  GPAY.Array = XPAY
    GPAY.ReBind
    GPAY.Bookmark = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.LoadPAYGrid"
End Sub
Private Sub ClearZGrid()
    On Error GoTo errHandler
    If Not XZ Is Nothing Then
        XZ.Clear
        XZ.ReDim 0, 0, 1, 8
    End If
    GZ.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.ClearZGrid"
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
Private Sub ClearExGrid()
    On Error GoTo errHandler
    If Not XE Is Nothing Then
        XE.Clear
        XE.ReDim 0, 0, 1, 6
    End If
    GE.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.ClearExGrid"
End Sub

Private Sub ClearCSLGrid()
    On Error GoTo errHandler
    If Not XCSL Is Nothing Then
        XCSL.Clear
        XCSL.ReDim 0, 0, 1, 5
    End If
'    GCSL.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.ClearCSLGrid"
End Sub

Private Sub ClearPAYGrid()
    On Error GoTo errHandler
    If Not XPAY Is Nothing Then
        XPAY.Clear
        XPAY.ReDim 0, 0, 1, 5
    End If
    GPAY.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.ClearPAYGrid"
End Sub






Private Sub GE_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If XE(Bookmark, 11) = True Then
        RowStyle.BackColor = &HC0C0C0
        RowStyle.Font.Strikethrough = True
    Else
        RowStyle.BackColor = &H80000018
    End If

End Sub

Private Sub GCSL_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If XCSL(Bookmark, 3) < 0 Then
        RowStyle.BackColor = RGB(181, 230, 234)
       ' RowStyle.Font.Strikethrough = True
    Else
        RowStyle.BackColor = &H80000018
    End If

End Sub
Private Sub GO_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If GESwitch = True Then
        GESwitch = False
    Else
        GE.Splits(0).ForeColor = COLOR_CANCELLED
    End If
End Sub

Private Sub GZ_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If GESwitch = True Then
        GESwitch = False
    Else
        GE.Splits(0).ForeColor = COLOR_CANCELLED
    End If
End Sub

Private Sub GZ_SelChange(Cancel As Integer)
    If flgLoading Then Exit Sub
    
    If GZ.Bookmark = 0 Then Exit Sub
    GZ.MousePointer = dbgMPHourglass
    RefreshOps
    RefreshExchanges
    RefreshDetails
    If IsNull(GO.Bookmark) Then
        Me.lblExchanges.Caption = "Exchanges for day-session: " & XZ.Value(GZ.Bookmark, 1)
    Else
        Me.lblExchanges.Caption = "Exchanges for day-session: " & XZ.Value(GZ.Bookmark, 1) & " and operator-session started: " & IIf(IsNull(GO.Bookmark), "", Format(XO.Value(GO.Bookmark, 2), "HH:NN"))  '
    End If
    GE.Splits(0).ForeColor = &H8000000D
    GESwitch = True
    GZ.MousePointer = dbgMPDefault
End Sub
Private Sub GO_SelChange(Cancel As Integer)
    If flgLoading Then Exit Sub
    If GO.Bookmark = 0 Then Exit Sub
    
    GO.MousePointer = dbgMPHourglass
        
    DoEvents
    RefreshExchanges
    RefreshDetails
    GE.Splits(0).ForeColor = &H8000000D
    GESwitch = True
    Me.lblExchanges.Caption = "Exchanges for day-session: " & XZ.Value(GZ.Bookmark, 1) & " and operator-session started: " & Format(XO.Value(GO.Bookmark, 2), "HH:NN")
    GO.MousePointer = dbgMPDefault

End Sub

Private Sub GE_SelChange(Cancel As Integer)

    If flgLoading Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    If GE.Bookmark = 0 Then Exit Sub
    
    RefreshDetails

    Screen.MousePointer = vbDefault
End Sub
Private Sub GE_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If flgLoading Then Exit Sub
    Screen.MousePointer = vbHourglass
    If GE.Bookmark = 0 Then Exit Sub
    RefreshDetails

    Screen.MousePointer = vbDefault

End Sub

Private Sub LoadZSessions()
    On Error GoTo errHandler
    
    Set ocZ = Nothing
    Set ocZ = New c_ZSession
    Screen.MousePointer = vbHourglass
    
    If txtArg1 > "" Then
        If IsDate(txtArg1) Then
            ocZ.Load CDate(0), CDate(txtArg1), 0
        ElseIf IsNumeric(txtArg1) Then
            ocZ.Load CDate(0), CDate(0), txtArg1
        Else
            ocZ.Load DateAdd("m", -2, Date), CDate(0), 0
        End If
    Else
        ocZ.Load DateAdd("m", -2, Date), CDate(0), 0
    End If
    
    
    ClearZGrid
    LoadZGrid
    RefreshOps
   
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.LoadZSessions"
End Sub
Private Sub RefreshOps()
    On Error GoTo errHandler
    
    flgLoading = True
    
    If Not ocCS Is Nothing Then ClearOpsGrid
    Set ocCS = New c_CSs
    If IsNull(GZ.Bookmark) Then Exit Sub
    If Not XZ(GZ.Bookmark, 6) = Empty Then
        ocCS.LoadByZID XZ(GZ.Bookmark, 6)
        LoadOpsGrid
    Else
        ClearOpsGrid
    End If
    
    
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.RefreshOps()", , EA_NORERAISE
    HandleError
End Sub
Private Sub RefreshExchanges()
    On Error GoTo errHandler
    
    flgLoading = True
    If IsNull(GO.Bookmark) Then Exit Sub
    If Not ocEX Is Nothing Then ClearExGrid
    
    Set ocEX = New c_Exchanges
    If Not XO(GO.Bookmark, 6) = Empty Then
        ocEX.Load XO(GO.Bookmark, 6), ""
        LoadExGrid
    
    Else
        ClearExGrid
    End If
    
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.RefreshExchanges()", , EA_NORERAISE
    HandleError
End Sub
Private Sub RefreshDetails()
    On Error GoTo errHandler
    If IsNull(GE.Bookmark) Then Exit Sub
    
    If Not XE(GE.Bookmark, 9) = Empty Then
        LoadCSLGrid ocEX(XE(GE.Bookmark, 9)).CSLS
        LoadPAYGrid ocEX(XE(GE.Bookmark, 9)).PAYS
    Else
        ClearCSLGrid
        ClearPAYGrid
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.RefreshDetails()", , EA_NORERAISE
    HandleError
End Sub
'Private Sub cmdAuto_Click()
'
'   GZ.Bookmark = 1
'   TimerON IIf(cmdAuto.Value = 1, True, False)
'
'End Sub

'Private Sub TimerON(pOn As Boolean)
'    Timer1.Enabled = pOn
'    cmdAuto.Value = IIf(pOn = True, 1, 0)
'End Sub

Private Sub txtArg1_Change()

End Sub

Private Sub txtArg1_DblClick()
    txtArg1 = ""
End Sub
