VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExchanges 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Front desk activity"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   13545
   Begin VB.TextBox txtBranchCode 
      Alignment       =   2  'Center
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
      Height          =   390
      Left            =   270
      TabIndex        =   15
      Top             =   330
      Width           =   930
   End
   Begin VB.TextBox txtExchangeNumber 
      Alignment       =   2  'Center
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
      Height          =   390
      Left            =   6270
      TabIndex        =   13
      Top             =   345
      Width           =   930
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   360
      Left            =   2415
      TabIndex        =   8
      Top             =   360
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   61276161
      CurrentDate     =   38644
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Refresh"
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
      Left            =   10965
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   270
      Width           =   945
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
      Left            =   12690
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2295
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
      Height          =   2580
      Left            =   240
      TabIndex        =   2
      Top             =   4845
      Width           =   12375
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
         Left            =   7830
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   1830
      End
      Begin TrueOleDBGrid60.TDBGrid GPAY 
         Height          =   1980
         Left            =   8895
         OleObjectBlob   =   "frmExchanges.frx":0000
         TabIndex        =   3
         Top             =   330
         Width           =   3165
      End
      Begin TrueOleDBGrid60.TDBGrid GCSL 
         Height          =   1950
         Left            =   225
         OleObjectBlob   =   "frmExchanges.frx":3C51
         TabIndex        =   7
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
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3105
      Width           =   2100
   End
   Begin TrueOleDBGrid60.TDBGrid GE 
      Height          =   3825
      Left            =   270
      OleObjectBlob   =   "frmExchanges.frx":9AEA
      TabIndex        =   0
      Top             =   915
      Width           =   12330
   End
   Begin MSComCtl2.DTPicker dtTO 
      Height          =   360
      Left            =   4005
      TabIndex        =   9
      Top             =   360
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   61276161
      CurrentDate     =   38644
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "# number"
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
      Height          =   285
      Left            =   6300
      TabIndex        =   14
      Top             =   60
      Width           =   1140
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Height          =   285
      Left            =   4095
      TabIndex        =   12
      Top             =   60
      Width           =   930
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Height          =   285
      Left            =   2445
      TabIndex        =   11
      Top             =   60
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Branch"
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
      Height          =   285
      Left            =   285
      TabIndex        =   10
      Top             =   90
      Width           =   930
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
'Dim ocZ As c_ZSession
'Dim ocCS As c_CSs
Dim ocEX As c_Exchanges
Dim flgLoading As Boolean


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
    LoadExchanges
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    top = 310
    left = 120
    Width = 13600
    Height = 8000
    Me.dtTO = DateAdd("d", 1, Date)
    Me.dtFrom = DateAdd("ww", -1, dtTO)
    
    Screen.MousePointer = vbHourglass
    Me.cmdRefresh.Visible = True
    
    Set XE = New XArrayDB
    XE.ReDim 1, 1, 1, 11
    Set GE.Array = XE
    
    Set XCSL = New XArrayDB
    XCSL.ReDim 1, 1, 1, 7
    Set GCSL.Array = XCSL
    
    Set XPAY = New XArrayDB
    XPAY.ReDim 1, 1, 1, 3
    Set GPAY.Array = XPAY
    
    
    flgLoading = False
        
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.Form_Load", , EA_NORERAISE
    HandleError
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
    XE.ReDim 1, ocEX.Count, 1, 13
    For i = 1 To ocEX.Count
        XE.Value(i, 1) = ocEX.Item(i).BranchName
        XE.Value(i, 2) = ocEX.Item(i).StationName
        XE.Value(i, 3) = ocEX.Item(i).ExchangeNumber
        XE.Value(i, 4) = ocEX.Item(i).ExchangeDateF
        XE.Value(i, 5) = ocEX.Item(i).SalesPersonName
        XE.Value(i, 6) = ocEX.Item(i).TotalPayableF
        XE.Value(i, 7) = ocEX.Item(i).TotalVATF
        XE.Value(i, 8) = ocEX.Item(i).ChangeGivenF
        XE.Value(i, 9) = ocEX.Item(i).ExchangeTypeF & IIf(ocEX.Item(i).VOIDED = True, "(Voided)", "")
'        If ocEX.Item(i).VOIDS > 0 Then
'            XE.Value(i, 10) = "Voids - " & CStr(ocEX.Item(i).VOIDS)
'        Else
            XE.Value(i, 10) = ocEX.Item(i).Note
'        End If
        XE.Value(i, 11) = ocEX.Item(i).ID
        XE.Value(i, 12) = ocEX.Item(i).ExchangeDateSort
        XE.Value(i, 13) = ocEX.Item(i).VOIDED
    Next
    XE.QuickSort 1, XE.UpperBound(1), 3, XORDER_DESCEND, XTYPE_NUMBER
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
    XCSL.Clear
    XCSL.ReDim 1, cCSL.Count, 1, 7
    For i = 1 To cCSL.Count
        XCSL.Value(i, 1) = cCSL.Item(i).CodeF
        XCSL.Value(i, 2) = cCSL.Item(i).Title
        XCSL.Value(i, 3) = cCSL.Item(i).Qty
        XCSL.Value(i, 4) = cCSL.Item(i).PriceF
        XCSL.Value(i, 5) = cCSL.Item(i).DiscountCombinedF
        XCSL.Value(i, 6) = cCSL.Item(i).ExtensionF
        XCSL.Value(i, 7) = cCSL.Item(i).vatratef
        
    Next
    XCSL.QuickSort 1, XCSL.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    Set GCSL.Array = XCSL
    GCSL.ReBind
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
    XPAY.Clear
    XPAY.ReDim 1, cPAY.Count, 1, 3
    For i = 1 To cPAY.Count
        XPAY.Value(i, 1) = cPAY.Item(i).PaymentTypeF
        XPAY.Value(i, 2) = cPAY.Item(i).AmtF
    Next
    GPAY.ReBind
    GPAY.Bookmark = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.LoadPAYGrid"
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

Private Sub GE_DblClick()
Dim oEX As New c_Exchanges
    oEX.LoadCSL XE(GE.Bookmark, 11)
    LoadCSLGrid oEX.colCSL
    oEX.LoadPay XE(GE.Bookmark, 11)
    LoadPAYGrid oEX.colPay
    
End Sub

Private Sub GE_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
On Error Resume Next
    If XE(Bookmark, 13) = True Then
        RowStyle.BackColor = &HC0C0C0
    Else
        RowStyle.BackColor = &H80000018
    End If

End Sub


Private Sub GE_HeadClick(ByVal ColIndex As Integer)
Static Direction As Variant
    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    
    Select Case ColIndex
    Case 1, 2, 3, 4
        XE.QuickSort XE.LowerBound(1), XE.UpperBound(1), ColIndex + 1, Direction, XTYPE_STRING
    Case Else
        XE.QuickSort XE.LowerBound(1), XE.UpperBound(1), ColIndex + 1, Direction, XTYPE_NUMBER
    End Select
    
    GE.Refresh
    Screen.MousePointer = vbDefault
End Sub

Private Sub GE_SelChange(Cancel As Integer)
    If flgLoading Then Exit Sub
    Screen.MousePointer = vbHourglass
    If GE.Bookmark = 0 Then Exit Sub
    RefreshDetails

    Screen.MousePointer = vbDefault
End Sub
'
Private Sub LoadExchanges()
    On Error GoTo errHandler

    Set ocEX = New c_Exchanges
    
    ocEX.Load Me.dtFrom, Me.dtTO, FNN(txtExchangeNumber), txtBranchCode

    Screen.MousePointer = vbHourglass

    LoadExGrid

    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges.LoadZSessions"
End Sub

Private Sub RefreshDetails()
    On Error GoTo errHandler
    If IsNull(GE.Bookmark) Then Exit Sub
    
    If Not XE(GE.Bookmark, 11) = Empty Then
        LoadCSLGrid ocEX(XE(GE.Bookmark, 11)).CSLS
        LoadPAYGrid ocEX(XE(GE.Bookmark, 11)).PAYS
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
