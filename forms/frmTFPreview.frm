VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Begin VB.Form frmTFPreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Transfer"
   ClientHeight    =   5835
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11100
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmTFPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLabels 
      BackColor       =   &H00D7D1BF&
      Caption         =   "Print product labels"
      Height          =   405
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4965
      Width           =   1755
   End
   Begin VB.CommandButton cmdHeader 
      BackColor       =   &H00FFC0C0&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5175
      Width           =   255
   End
   Begin VB.TextBox txtTPMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1185
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1935
      Visible         =   0   'False
      Width           =   5280
   End
   Begin VB.CommandButton cmdMemo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4755
      Width           =   255
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmTFPreview.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print the invoice"
      Top             =   4770
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1335
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmTFPreview.frx":2B2C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print or preview"
      Top             =   4770
      Width           =   855
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00C4BCA4&
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
      Height          =   615
      Left            =   2205
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmTFPreview.frx":2EB6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close the transfer"
      Top             =   4770
      Width           =   855
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   4455
      Left            =   90
      OleObjectBlob   =   "frmTFPreview.frx":3240
      TabIndex        =   2
      Top             =   195
      Width           =   10725
   End
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   9510
      X2              =   11070
      Y1              =   -45
      Y2              =   945
   End
   Begin VB.Label lblTotalCaption 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1140
      Left            =   5325
      TabIndex        =   1
      Top             =   4695
      Width           =   3495
   End
   Begin VB.Label lblTotalValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1140
      Left            =   8895
      TabIndex        =   0
      Top             =   4710
      Width           =   1845
   End
End
Attribute VB_Name = "frmTFPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cTF As c_TF
Dim oTF As a_TF
Dim dblTotal As Double
Dim XA As XArrayDB
Dim bMemoExpanded As Boolean
Dim mbShowMemo As Boolean

Dim PrintCommandButtonCTRLDown As Boolean
Private Sub Form_Initialize()
    PrintCommandButtonCTRLDown = False
End Sub
Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim ShiftTest As Integer
   PrintCommandButtonCTRLDown = False
   ShiftTest = Shift And 7
   Select Case ShiftTest
      Case 1 ' or vbShiftMask
      Case 2 ' or vbCtrlMask
         PrintCommandButtonCTRLDown = True
      End Select
End Sub

Private Sub cmdPrint_KeyUp(KeyCode As Integer, Shift As Integer)
        PrintCommandButtonCTRLDown = False
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oTF.Status = stInProcess And oTF.IsNew = False)
    Forms(0).mnuCancel.Enabled = (oTF.Status = stISSUED) Or (oTF.Status = stCOMPLETE)
'    Forms(0).mnuCancelLine.Enabled = (oTF.statusF = "ISSUED" And oTF.IsNew = False)
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = False
    'Forms(0).mnuInvAdd.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Forms(0).mnuCopyLines.Enabled = True
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.SetMenu"
End Sub

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.G1, Me.Name, Me.Height, Me.Width
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.mnuSaveLayout"
End Sub

Private Sub cmdHeader_Click()
    On Error GoTo errHandler
    Header
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.cmdHeader_Click", , EA_NORERAISE
    HandleError
End Sub
Public Sub mnuHeader()
    On Error GoTo errHandler
    Header
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.mnuHeader"
End Sub
Private Sub Header()
    On Error GoTo errHandler
Dim frm As New frmHeader_TFR
Dim strRef As String
Dim strMemo As String
    
    oTF.BeginEdit
    frm.component oTF
    frm.Show vbModal
    oTF.ApplyEdit
    
    Unload frm

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.Header"
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo errHandler

Dim blnEdit As Boolean
Dim frm As frmTFR2
Dim bCancel As Boolean
    WaitMsg "Loading . . .", True, Me
    Set frm = New frmTFR2
    blnEdit = True
    frm.component "", False, 0, oTF
    frm.Show
    WaitMsg "", False, Me

EXIT_Handler:
    Unload Me
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdLabels_Click()
    On Error GoTo errHandler
    PrintLabels
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.cmdLabels_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub PrintLabels()
    On Error GoTo errHandler
Dim frm As frmPrintLabels
    Set frm = New frmPrintLabels
    frm.component "T", oTF
    frm.Show vbModal
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmDELPreview.PrintLabels"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.PrintLabels"
End Sub

Private Sub cmdMemo_Click()
    On Error GoTo errHandler

    ShowMemo Not mbShowMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.cmdMemo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub ShowMemo(bON As Boolean)
            On Error Resume Next
        mbShowMemo = bON
        txtTPMemo.Visible = bON
        If bON Then txtTPMemo.SetFocus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.ShowMemo(bOn)", bON
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_TFR
Dim i As Long
Dim oDOC As a_DocumentControl
Dim qtyLinesToPrint As Integer
Dim Dummy As String
    If PrintCommandButtonCTRLDown Then
        PrintCommandButtonCTRLDown = False

        Screen.MousePointer = vbHourglass
        oTF.TFLines.SortLines enSequence, True

        Set oDOC = oPC.Configuration.DocumentControls.FindDC(oTF.constDOCCODE)
        If oDOC Is Nothing Then
            qtyLinesToPrint = 1
        Else
            qtyLinesToPrint = oPC.Configuration.DocumentControls.FindDC(oTF.constDOCCODE).QtyCopies
        End If

       If oTF.ExportToXML(enView, , , qtyLinesToPrint) = False Then
           Screen.MousePointer = vbDefault
           MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
       End If
       Screen.MousePointer = vbDefault
    Else
        Set frm = New frmPrintingOptions_TFR
        frm.ComponentObject oTF
        frm.Show vbModal
    End If
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Public Sub component(PID As Long)
    On Error GoTo errHandler
Dim lngID As Long
    lngID = PID
    Set oTF = New a_TF
    oTF.Load lngID
    Me.Caption = "Transfer: " & oTF.DOCCode & " :   " & IIf(oTF.InOut = "IN", " IN from  ", " OUT to  ") & oTF.DestinationName & "  " & oTF.StatusF & "  " & oTF.DocDateF & "(" & oTF.StaffName & ")"
    LoadControls
    SetMenu
    mbShowMemo = False
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.component(PID)", PID
End Sub
Public Sub ComponentObject(pCS As a_TF)
    On Error GoTo errHandler
    Set oTF = pCS
    Me.Caption = "Transfer: " & oTF.DOCCode & " :   " & IIf(oTF.InOut = "IN", " IN from  ", " OUT to  ") & oTF.DestinationName & "  " & oTF.StatusF & "  " & oTF.DocDateF & "(" & oTF.StaffName & ")"
    
    mbShowMemo = False
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.ComponentObject(pCS)", pCS
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
Dim strTotalCaption As String
Dim strTotalValues As String
        CancelLine.Visible = (oTF.Status = stCANCELLED Or oTF.Status = stVOID)
        If oTF.Status = stInProcess Then
            cmdEdit.Enabled = True
        Else
            cmdEdit.Enabled = False
        End If
        With oTF
            txtTPMemo = IIf(Len(.Memo) > 0, "Note:  " & Trim$(.Memo), "")
          '  txtTPMemo.Visible = (txtTPMemo > "")
            .DisplayTotals strTotalCaption, strTotalValues
        lblTotalCaption.Caption = strTotalCaption
        lblTotalValues.Caption = strTotalValues
        End With
        LoadListView
        cmdHeader.Visible = (oTF.InOut = "IN")
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.LoadControls"
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub LoadListView()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Integer
Dim currDeposit As Currency
Dim currPrice As Currency
Dim dblVAT As Double
Dim strSummaryDescription As String
Dim strSummary As String
Dim lngTotal As Long
Dim lngDepositTotal As Long

    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, oTF.TFLines.Count, 1, 12

    For i = 1 To oTF.TFLines.Count
        With oTF.TFLines(i)
            XA(i, 1) = .CodeF
            XA(i, 2) = .Title
            XA(i, 3) = .Qty
            XA(i, 4) = .PriceF
            XA(i, 5) = .DiscountF
            XA(i, 6) = .ExtLessDiscF
            XA(i, 7) = .ExtLessDiscExVATF
            XA(i, 8) = .Note
            XA(i, 9) = .ID & "k"
            XA(i, 10) = .PID
            XA(i, 11) = .code
            XA(i, 12) = .EAN
        End With
    Next i
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", "frmTFPreview", CStr(i), G1.Columns(i - 1).Width)
    Next
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    G1.Array = XA
    G1.ReBind
    
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.LoadListView"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = vbKey4 Then
        If MsgBox("Confirm close?", vbOKCancel, "Close form") = vbOK Then
            Unload Me
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.Form_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Me.TOP = 50
        Me.Left = 50
'        Me.Height = 6500
'        Me.Width = 11500
    End If
    Me.Width = GetSetting("PBKS", Me.Name, "width", 11500)
    Me.Height = GetSetting("PBKS", Me.Name, "height", 6500)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    G1.Width = NonNegative_Lng(Me.Width - (G1.Left + 550))
    lngDiff = G1.Height
    G1.Height = NonNegative_Lng(Me.Height - (G1.TOP + 1700))
    lngDiff = (G1.Height - lngDiff)
    cmdEdit.TOP = cmdEdit.TOP + lngDiff
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    cmdClose.TOP = cmdClose.TOP + lngDiff
    txtTPMemo.TOP = txtTPMemo.TOP + lngDiff
    lblTotalCaption.TOP = lblTotalCaption.TOP + lngDiff
    lblTotalValues.TOP = lblTotalValues.TOP + lngDiff
    cmdMemo.TOP = cmdMemo.TOP + lngDiff
    cmdHeader.TOP = cmdHeader.TOP + lngDiff
    cmdLabels.TOP = cmdLabels.TOP + lngDiff
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set oTF = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub G1_Click()
Dim str As String
On Error Resume Next
    If XA.UpperBound(1) = 0 Then Exit Sub
    If IsNull(G1.Bookmark) Then Exit Sub
    If Err Then Exit Sub
    
On Error GoTo errHandler
    str = IIf(FNS(XA.Value(G1.Bookmark, 12)) > "", FNS(XA.Value(G1.Bookmark, 12)), FNS(XA.Value(G1.Bookmark, 11)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.G1_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub G1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 12)) > "", FNS(XA.Value(G1.Bookmark, 12)), FNS(XA.Value(G1.Bookmark, 11)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.G1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub G1_SelChange(Cancel As Integer)
    On Error GoTo errHandler
Dim str As String

    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
    
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.G1_SelChange(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub G1_DblClick()
    On Error GoTo errHandler
Dim frmA As frmProductPrevAQ
Dim frm As frmProductPrev
Dim oP As a_Product
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    
    str = IIf(FNS(XA.Value(G1.Bookmark, 12)) > "", FNS(XA.Value(G1.Bookmark, 12)), FNS(XA.Value(G1.Bookmark, 11)))
    If str = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oP.Load FNS(XA.Value(G1.Bookmark, 10)), 0 'oDel.DeliveryLines.FindLineByID(val(Me.lvw.SelectedItem.Key)).pID, 0
    If oPC.Configuration.AntiquarianYN Then
        Set frmA = New frmProductPrevAQ
        frmA.component oP
        frmA.Show
    Else
        Set frm = New frmProductPrev
        frm.component oP
        frm.Show
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmTFPreview: G1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmTFPreview: G1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.G1_DblClick", , EA_NORERAISE
    HandleError
End Sub


Public Sub mnuCancel()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oSM.PostTransfer oTF.TRID, stCANCELLED, True
    RefreshData
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.mnuCancel"
End Sub

Public Sub RefreshData()
    On Error GoTo errHandler
    oTF.Reload
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.RefreshData"
End Sub


Public Sub mnuVoid()
    On Error GoTo errHandler
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oTF.VoidDocument
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.mnuVoid"
End Sub

Private Sub G1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1)
    
    G1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2
            GetRowType = XTYPE_STRING
        Case Else
            GetRowType = XTYPE_NUMBER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.GetRowType(ColIndex)", ColIndex
End Function

Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
Dim oSM As New z_StockManager
    ofrm.component oTF.Memo
    ofrm.Show vbModal
    oSM.SetMemo ofrm.Memo, oTF.TRID
    txtTPMemo.Visible = (ofrm.Memo > "")
    txtTPMemo = ofrm.Memo
    oTF.Memo = ofrm.Memo
    Unload ofrm
    Set ofrm = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.mnuMemo"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.mnuMemo"
End Sub

Private Sub txtTPMemo_Change()
    On Error GoTo errHandler
    txtTPMemo = HandleTextWithBites(txtTPMemo)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.txtTPMemo_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_DblClick()
    On Error GoTo errHandler
    If bMemoExpanded Then
        txtTPMemo.Height = txtTPMemo.Height - 800
        txtTPMemo.Width = txtTPMemo.Width - 800
        txtTPMemo.TOP = txtTPMemo.TOP + 800
        bMemoExpanded = False
        txtTPMemo.ZOrder 1
    Else
        bMemoExpanded = True
        txtTPMemo.Height = txtTPMemo.Height + 800
        txtTPMemo.Width = txtTPMemo.Width + 800
        txtTPMemo.TOP = txtTPMemo.TOP - 800
        txtTPMemo.ZOrder 0
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.txtTPMemo_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_LostFocus()
    On Error GoTo errHandler
    If bMemoExpanded Then
        txtTPMemo.Height = txtTPMemo.Height - 800
        txtTPMemo.Width = txtTPMemo.Width - 800
        txtTPMemo.TOP = txtTPMemo.TOP + 800
        bMemoExpanded = False
        txtTPMemo.ZOrder 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.txtTPMemo_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    txtTPMemo = HandleTextWithBites(txtTPMemo)

'    If InStr(1, txtTPMemo, Chr(13)) > 0 Then
'        If MsgBox("There are multiple lines in the memo you are saving.", vbExclamation + vbOKCancel, "Warning") = vbCancel Then
'            Cancel = True
'            Exit Sub
'        End If
'    End If
Dim oSM As New z_StockManager
    oSM.SetMemo txtTPMemo, oTF.TRID
    oTF.SetMemo txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.txtTPMemo_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_DragOver(Source As Control, x As Single, _
    Y As Single, State As Integer)
    On Error GoTo errHandler
    Dim picdocument As PictureBox
        ' Optionally move the cursor position so
        ' the user can see where the drop would happen.
        txtTPMemo.SelStart = TextBoxCursorPos(txtTPMemo, x, Y)
        txtTPMemo.SelLength = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.txtTPMemo_DragOver(Source,x,Y,State)", Array(Source, x, Y, State), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_DragDrop(Source As Control, x As Single, _
    Y As Single)
    On Error GoTo errHandler
    txtTPMemo.SelStart = TextBoxCursorPos(txtTPMemo, x, Y)
    txtTPMemo.SelLength = 0
    txtTPMemo.SelText = Source
Dim oSM As New z_StockManager
    oSM.SetMemo txtTPMemo, oTF.TRID
    oTF.SetMemo txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.txtTPMemo_DragDrop(Source,x,Y)", Array(Source, x, Y), EA_NORERAISE
    HandleError
End Sub

Public Sub mnuCopyLines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oLine As a_TFL
Dim fs As New FileSystemObject

    oPC.PrepareLinesClipboard
    Set rs = oPC.LinesClipboard
    rs.open
    For Each oLine In oTF.TFLines
        rs.AddNew
        rs.fields("GUID") = CreateGUID
        rs.fields("PID") = oLine.PID
        rs.fields("Qty") = oLine.Qty
        rs.fields("QtyFirm") = oLine.Qty
        rs.fields("QtySS") = 0
        rs.fields("Price") = oLine.Price
        rs.fields("DISCOUNTRATE") = oLine.Discount
        rs.fields("CODEF") = oLine.CodeF
        rs.fields("EANF") = oLine.EAN
        rs.fields("TITLE") = oLine.Title
        rs.fields("VATRATE") = oLine.VATRate
        rs.fields("REF") = oLine.Note
        rs.Update
    Next
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
        fs.CreateFolder (oPC.SharedFolderRoot & "\TEMP")
        If Err <> 0 Then
            MsgBox "Cannot create folder for Papyrus clipboard", vbInformation + vbOKOnly, "Can't do this"
        End If
    End If
    If fs.FileExists(oPC.SharedFolderRoot & "\TEMP\Clipboard.rs") Then
        fs.DeleteFile oPC.SharedFolderRoot & "\TEMP\Clipboard.rs"
    Else
        If fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
            rs.Save oPC.SharedFolderRoot & "\TEMP\Clipboard.rs"
        End If
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTFPreview.mnuCopyLines"
End Sub


