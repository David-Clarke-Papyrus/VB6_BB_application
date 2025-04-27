VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmExchanges1 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Front desk activity"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   14070
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
      Picture         =   "frmExchanges1.frx":0000
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
      Left            =   4695
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3315
      Width           =   1815
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
      Left            =   4470
      TabIndex        =   9
      Top             =   3825
      Width           =   8265
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
         Left            =   6315
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2490
         Width           =   1695
      End
      Begin TrueOleDBGrid60.TDBGrid GPAY 
         Height          =   1140
         Left            =   240
         OleObjectBlob   =   "frmExchanges1.frx":038A
         TabIndex        =   10
         Top             =   1950
         Width           =   6030
      End
      Begin TrueOleDBGrid60.TDBGrid GCSL 
         Height          =   1470
         Left            =   240
         OleObjectBlob   =   "frmExchanges1.frx":45F7
         TabIndex        =   13
         Top             =   345
         Width           =   7755
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
      Left            =   10620
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3345
      Width           =   2100
   End
   Begin TrueOleDBGrid60.TDBGrid GE 
      Height          =   2970
      Left            =   4455
      OleObjectBlob   =   "frmExchanges1.frx":A3AC
      TabIndex        =   2
      Top             =   285
      Width           =   8250
   End
   Begin TrueOleDBGrid60.TDBGrid GZ 
      Height          =   2235
      Left            =   15
      OleObjectBlob   =   "frmExchanges1.frx":10173
      TabIndex        =   3
      Top             =   765
      Width           =   4320
   End
   Begin TrueOleDBGrid60.TDBGrid GO 
      Height          =   1245
      Left            =   45
      OleObjectBlob   =   "frmExchanges1.frx":14FBA
      TabIndex        =   7
      Top             =   3390
      Width           =   4320
   End
   Begin VB.Label lblNotes 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Notes"
      ForeColor       =   &H8000000D&
      Height          =   2175
      Left            =   30
      TabIndex        =   15
      Top             =   4830
      Width           =   4305
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
      TabIndex        =   14
      Top             =   75
      Width           =   1950
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
      Left            =   4500
      TabIndex        =   6
      Top             =   45
      Width           =   1395
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
Attribute VB_Name = "frmExchanges1"
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
Dim OPSID As Variant
Dim ocZ As c_ZSession
Dim ocCS As c_CSs
Dim ocEX As c_Exchanges

Dim rsE As ADODB.Recordset
Dim SelectedExchID As String
Dim rsCSL As ADODB.Recordset
Dim rsPAY As ADODB.Recordset

Dim flgLoading As Boolean
Dim GESwitch As Boolean
Const strNotes As String = "Notes:" & vbCrLf & "1. Enter an exchange number or a date or leave blank for recent day-sessions and click ther tick button." & vbCrLf _
        & "2. Select the day-session you wish to examine." & vbCrLf _
        & "then . . ." & vbCrLf _
        & "3. Select the operator-session of the selected day-session." & vbCrLf _
        & "Make selections by clicking on the grey margin of the day-session and operator-session grids respectively. The mouse pointer will show a right-arrow while the data is being fetched."
Public Sub SaveFormLayout()
    On Error GoTo errHandler
    SaveLayout Me.GZ, Me.Name & "1", Me.Height, Me.Width
    SaveLayout Me.GO, Me.Name & "2"
    SaveLayout Me.GE, Me.Name & "3"
    SaveLayout Me.GCSL, Me.Name & "4"
    SaveLayout Me.GPAY, Me.Name & "5"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.SaveFormLayout"
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub G1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuReserveList   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.G1_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), EA_NORERAISE
    HandleError
End Sub



Private Sub cmdFix_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If IsNull(GE.Bookmark) Then Exit Sub
    oSM.CreateInvoiceFromExchange XE(GE.Bookmark, 9)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.cmdFix_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdGet_Click()
    On Error GoTo errHandler
    cmdRefreshZ_Click
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.cmdGet_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrintSale_Click()
    On Error GoTo errHandler
Dim ar As New arExchange


    If IsNull(GE.Bookmark) Then Exit Sub
    
    If XE(GE.Bookmark, 9) = Empty Then Exit Sub
    Set ocEX = New c_Exchanges
    ocEX.Load "", "", FNS(XE(GE.Bookmark, 9))
    ar.component XE(GE.Bookmark, 1), XZ(GZ.Bookmark, 3), ocEX.Item(1).ExchangeDate2F, XE(GE.Bookmark, 3), ocEX(1).CSLS, ocEX(1).PAYS, XE(GE.Bookmark, 6), XE(GE.Bookmark, 7), XE(GE.Bookmark, 8), ocEX.Item(1).VOIDED
    ar.Caption = "Exchange number: " & FNS(XE(GE.Bookmark, 1)) & " on station " & FNS(XZ(GZ.Bookmark, 3))
    ar.Show vbModal
    Set ocEX = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.cmdPrintSale_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo errHandler
        Screen.MousePointer = vbHourglass
        
        DoEvents
        RefreshExchanges
        RefreshDetails
        
        Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.cmdRefresh_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdPrintList_Click()
    On Error GoTo errHandler
Dim ar As New arExchanges
ar.Printer.Orientation = ddOLandscape
    
    ar.component XE, "Exchanges for X Session started: " & Format(XO.Value(GO.Bookmark, 2), "dd/mm/yyyy HH:NN AMPM")
    ar.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.cmdPrintList_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRefreshZ_Click()
    On Error GoTo errHandler
    LoadZSessions
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.cmdRefreshZ_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdZSession_Click()
    On Error GoTo errHandler
Dim ar As arZSession

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.cmdZSession_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Command1_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.Command1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim i As Integer
    flgLoading = True
    If Me.WindowState <> 2 Then
        TOP = 35
        Left = 30
        Width = 13000
        Height = 7500
    End If
    For i = 1 To GZ.Columns.Count
        GZ.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "1", CStr(i), GZ.Columns(i - 1).Width)
    Next
    Me.Height = GetSetting("PBKS", Me.Name, "Height", 7500)
    Me.Width = GetSetting("PBKS", Me.Name, "Width", 13000)
    
    For i = 1 To GO.Columns.Count
        GO.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "2", CStr(i), GO.Columns(i - 1).Width)
    Next
    For i = 1 To GE.Columns.Count
        GE.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "3", CStr(i), GE.Columns(i - 1).Width)
    Next
    For i = 1 To GCSL.Columns.Count
        GCSL.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "4", CStr(i), GCSL.Columns(i - 1).Width)
    Next
    For i = 1 To GPAY.Columns.Count
        GPAY.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "5", CStr(i), GPAY.Columns(i - 1).Width)
    Next
    
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
    ErrorIn "frmExchanges1.Form_Load", , EA_NORERAISE
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
        XZ.Value(i, 1) = ocZ.Item(i).StartDateF
        XZ.Value(i, 2) = ocZ.Item(i).EndDateF
        XZ.Value(i, 3) = ocZ.Item(i).TillPoint
        XZ.Value(i, 4) = ocZ.Item(i).SupervisorName
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
    ErrorIn "frmExchanges1.LoadZGrid"
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
    ErrorIn "frmExchanges1.LoadOpsGrid"
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
    XE.ReDim 1, rsE.RecordCount, 1, 11
    For i = 1 To rsE.RecordCount
        XE.Value(i, 1) = FNS(rsE.Fields(1))
        XE.Value(i, 2) = Format(FND(rsE.Fields(2)), "dd/mm Hh:Nn")
        XE.Value(i, 3) = IIf(FNS(rsE.Fields(3)) = "", "n.a.", FNS(rsE.Fields(3))) & IIf(FNS(rsE.Fields(16)) > "" And FNS(rsE.Fields(16)) <> FNS(rsE.Fields(3)), "/" & FNS(rsE.Fields(16)), "")
        XE.Value(i, 4) = Format(FNDBL(rsE.Fields(4)), "##0.00")
        XE.Value(i, 5) = Format(FNDBL(rsE.Fields(5)), "##0.00")
        XE.Value(i, 6) = Format(FNDBL(rsE.Fields(6)), "##0.00")
        XE.Value(i, 7) = FNS(rsE.Fields(7)) & IIf(FNB(rsE.Fields(8)) = True, "(Voided)", "")
        If (XE.Value(i, 7) = "OPEN DRAWER" Or XE.Value(i, 7) = "PETTY CASH") Then
            XE.Value(i, 8) = FNS(rsE.Fields(14))
        Else
            If FNN(rsE.Fields(9)) > 0 Then
                XE.Value(i, 8) = "Voids-" & CStr(FNN(rsE.Fields(9))) & " (" & FNS(rsE.Fields(14)) & ")"
            Else
                XE.Value(i, 8) = FNS(rsE.Fields(10))
            End If
        End If
        XE.Value(i, 9) = FNS(rsE.Fields(0))
'        XE.Value(i, 10) = ocEX.Item(i).ExchangeDateSort
'        XE.Value(i, 11) = ocEX.Item(i).voided
        rsE.MoveNext
    Next
   ' XE.QuickSort 1, XE.UpperBound(1), 1, XORDER_ASCEND, XTYPE_INTEGER
  '  GE.Array = XE
        On Error Resume Next
    GE.ReBind
    GE.Bookmark = 0 'GE.FirstRow
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.LoadExGrid"
End Sub

Private Sub LoadCSLGrid()
    On Error GoTo errHandler
Dim i As Integer
    XCSL.Clear
    If rsCSL.State = 0 Then
        GCSL.ReBind
        GCSL.Bookmark = 0
        GCSL.Refresh
        Exit Sub
    End If
    If rsCSL.BOF And rsCSL.eof Then
        GCSL.ReBind
        GCSL.Bookmark = 0
        GCSL.Refresh
        Exit Sub
    End If
    XCSL.ReDim 1, rsCSL.RecordCount, 1, 7
    rsCSL.MoveFirst
    For i = 1 To rsCSL.RecordCount
        XCSL.Value(i, 1) = FNS(rsCSL.Fields(13))
        XCSL.Value(i, 2) = FNS(rsCSL.Fields(2))
        XCSL.Value(i, 3) = FNS(rsCSL.Fields(8))
        XCSL.Value(i, 4) = Format(FNDBL(rsCSL.Fields(9)), "##0.00")
        If FNDBL(rsCSL.Fields(15)) > 0 Then
            XCSL.Value(i, 5) = Format(FNDBL(rsCSL.Fields(15)), "##0.00")
        Else
            XCSL.Value(i, 5) = ""
        End If
        XCSL.Value(i, 6) = Format(FNDBL(rsCSL.Fields(17)), "##0.00")
        XCSL.Value(i, 7) = Format(FNDBL(rsCSL.Fields(18)), "##0.00")
        rsCSL.MoveNext
    Next
    XCSL.QuickSort 1, XCSL.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    Set GCSL.Array = XCSL
    GCSL.ReBind
    GCSL.Bookmark = 0
    GCSL.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.LoadCSLGrid"
End Sub
Private Sub LoadPAYGrid()
    On Error GoTo errHandler
Dim i As Integer
    XPAY.Clear
    If rsPAY.State = 0 Then
        GPAY.ReBind
        GPAY.Bookmark = 0
        Exit Sub
    End If
    If rsPAY.BOF And rsPAY.eof Then
        GPAY.ReBind
        GPAY.Bookmark = 0
        Exit Sub
    End If
    XPAY.ReDim 1, rsPAY.RecordCount, 1, 3
    rsPAY.MoveFirst
    For i = 1 To rsPAY.RecordCount
        XPAY.Value(i, 1) = FNS(rsPAY.Fields(1))
        XPAY.Value(i, 2) = Format(FNS(rsPAY.Fields(4)), "##0.00")
        XPAY.Value(i, 3) = FNS(rsPAY.Fields(2))
        rsPAY.MoveNext
    Next
    GPAY.ReBind
    GPAY.Bookmark = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.LoadPAYGrid"
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
    ErrorIn "frmExchanges1.ClearZGrid"
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
    ErrorIn "frmExchanges1.ClearOpsGrid"
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
    ErrorIn "frmExchanges1.ClearExGrid"
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
    ErrorIn "frmExchanges1.ClearCSLGrid"
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
    ErrorIn "frmExchanges1.ClearPAYGrid"
End Sub






Private Sub Form_Resize()
    GE.Width = NonNegative_Lng(Me.Width - 5000)
    cmdPrintList.Left = NonNegative_Lng(Me.Width - 2630)
    frm1.Width = NonNegative_Lng(Me.Width - 5000)
    GCSL.Width = NonNegative_Lng(frm1.Width - 500)
   ' GPAY.Width = NonNegative_Lng(frm1.Width - 4400)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    SaveFormLayout
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub GE_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If IsNull(Bookmark) Then Exit Sub
    If Bookmark = 0 Then Exit Sub
    If XE(Bookmark, 11) = True Then
        RowStyle.BackColor = &HC0C0C0
        RowStyle.Font.Strikethrough = True
    Else
        RowStyle.BackColor = &H80000018
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.GE_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub GCSL_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If IsNull(Bookmark) Then Exit Sub
    If XCSL(Bookmark, 3) < 0 Then
        RowStyle.BackColor = RGB(181, 230, 234)
       ' RowStyle.Font.Strikethrough = True
    Else
        RowStyle.BackColor = &H80000018
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.GCSL_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub
Private Sub GO_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
    If GESwitch = True Then
        GESwitch = False
    Else
        GE.Splits(0).ForeColor = COLOR_CANCELLED
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.GO_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub GZ_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
    If GESwitch = True Then
        GESwitch = False
    Else
        GE.Splits(0).ForeColor = COLOR_CANCELLED
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.GZ_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub GZ_SelChange(Cancel As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    
    If IsNull(GZ.Bookmark) Then Exit Sub
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.GZ_SelChange(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub GO_SelChange(Cancel As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If IsNull(GO.Bookmark) Then Exit Sub
    If GO.Bookmark = 0 Then Exit Sub
    
    GO.MousePointer = dbgMPHourglass
        
    DoEvents
    RefreshExchanges
   ' RefreshDetails
    GE.Splits(0).ForeColor = &H8000000D
    GESwitch = True
    Me.lblExchanges.Caption = "Exchanges for day-session: " & XZ.Value(GZ.Bookmark, 1) & " and operator-session started: " & Format(XO.Value(GO.Bookmark, 2), "HH:NN")
    GO.MousePointer = dbgMPDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.GO_SelChange(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub GE_SelChange(Cancel As Integer)
    On Error GoTo errHandler

    If flgLoading Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    If IsNull(GE.Bookmark) Then Exit Sub
    If GE.Bookmark = 0 Then Exit Sub
    SelectedExchID = XE(GE.Bookmark, 9)
    RefreshDetails

    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.GE_SelChange(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub GE_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    Screen.MousePointer = vbHourglass
    If IsNull(GE.Bookmark) Then Exit Sub
    If GE.Bookmark = 0 Then Exit Sub
    SelectedExchID = XE(GE.Bookmark, 9)
    RefreshDetails

    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.GE_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
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
            ocZ.Load DateAdd("m", -6, Date), CDate(0), 0
        End If
    Else
        ocZ.Load DateAdd("m", -6, Date), CDate(0), 0
    End If
    
    
    ClearZGrid
    LoadZGrid
    RefreshOps
   
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.LoadZSessions"
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
    ErrorIn "frmExchanges1.RefreshOps"
End Sub
Private Sub RefreshExchanges()
    On Error GoTo errHandler
Dim OpenResult As Integer
    
    If IsNull(GO.Bookmark) Then Exit Sub
    If Not ocEX Is Nothing Then ClearExGrid
    flgLoading = True
    
    Set ocEX = New c_Exchanges
    If Not XO(GO.Bookmark, 6) = Empty Then

        Set rsE = New ADODB.Recordset
        rsE.CursorLocation = adUseClient
        Set rsCSL = New ADODB.Recordset
        rsCSL.CursorLocation = adUseClient
        Set rsPAY = New ADODB.Recordset
        rsPAY.CursorLocation = adUseClient
        '-------------------------------
        OpenResult = oPC.OpenDBSHort
        '-------------------------------
    
        rsE.Open "SELECT * FROM vExchangeBrowse WHERE OPID = '" & XO(GO.Bookmark, 6) & "' ORDER BY EXCHNUMBER DESC", oPC.COShort, adOpenKeyset, adLockReadOnly
        rsCSL.Open "SELECT * FROM vExchangeBrowseCSL2 WHERE  OPID = '" & FNS(XO(GO.Bookmark, 6)) & "'", oPC.COShort, adOpenKeyset, adLockReadOnly
        rsPAY.Open "SELECT * FROM vExchangeBrowsePAY WHERE  OPID = '" & FNS(XO(GO.Bookmark, 6)) & "'", oPC.COShort, adOpenKeyset, adLockReadOnly
        LoadExGrid
    Else
        ClearExGrid
    End If
    
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.RefreshExchanges"
End Sub
Private Sub RefreshDetails()
    On Error GoTo errHandler
    If IsNull(GE.Bookmark) Then Exit Sub
    If GE.Bookmark = 0 Then Exit Sub
    If Not XE(GE.Bookmark, 9) = Empty Then
        rsCSL.Filter = ""
        If SelectedExchID > "" Then rsCSL.Filter = "EXCHANGE_GUID = '" & SelectedExchID & "'"
        rsPAY.Filter = ""
        If SelectedExchID > "" Then rsPAY.Filter = "EXCHID = '" & SelectedExchID & "'"
        LoadCSLGrid
        LoadPAYGrid
'        LoadCSLGrid ocEX(XE(GE.Bookmark, 9)).CSLS
'        LoadPAYGrid ocEX(XE(GE.Bookmark, 9)).PAYS
    Else
        ClearCSLGrid
        ClearPAYGrid
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.RefreshDetails"
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
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.txtArg1_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtArg1_DblClick()
    On Error GoTo errHandler
    txtArg1 = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchanges1.txtArg1_DblClick", , EA_NORERAISE
    HandleError
End Sub

'''Sub LoadTree()
'''    'First we add four leveldef objects to gttree
'''    'to describe what each level will look like.
'''    'This allows us to set verious properties such
'''    'as colors, fonts, and heights on a per level
'''    'basis
'''    gt.LevelDefs.Add "Level0"
'''    gt.LevelDefs.Add "Level1"
'''    gt.LevelDefs.Add "Level2"
'''    gt.LevelDefs.Add "Level3"
'''    gt.LevelDefs.Add "Level4"
'''
'''    gt.LevelDefs(0).Font.Size = 12
'''
'''    gt.LevelDefs(1).Font.Size = 10
'''    gt.LevelDefs(1).Font.Bold = False
'''    gt.LevelDefs(1).ColumnDefs.ColumnCaptions = gtColumnCaptionsTrue
'''    gt.LevelDefs(1).ColumnDefs.Add , , "Tillpoint"
'''    gt.LevelDefs(1).ColumnDefs.Add , , "Start"
'''    gt.LevelDefs(1).ColumnDefs.Add , , "End"
'''    gt.LevelDefs(1).ColumnDefs.Add , , "Supervisor"
'''    gt.LevelDefs(1).ColumnDefs(0).Width = 600
'''    gt.LevelDefs(1).ColumnDefs(1).Width = 1000
'''    gt.LevelDefs(1).ColumnDefs(2).Width = 1000
'''    gt.LevelDefs(1).ColumnDefs(3).Width = 1000
'''
'''    gt.LevelDefs(2).Font.Size = 10
'''    gt.LevelDefs(2).Font.Bold = False
'''    gt.LevelDefs(2).ColumnDefs.ColumnCaptions = gtColumnCaptionsTrue
'''    gt.LevelDefs(2).ColumnDefs.Add , , "Start"
'''    gt.LevelDefs(2).ColumnDefs.Add , , "End"
'''    gt.LevelDefs(2).ColumnDefs.Add , , "Operator"
'''    gt.LevelDefs(2).ColumnDefs(0).Width = 600
'''    gt.LevelDefs(2).ColumnDefs(1).Width = 1000
'''    gt.LevelDefs(2).ColumnDefs(2).Width = 1000
'''
'''    gt.LevelDefs(3).Font.Size = 10
'''    gt.LevelDefs(3).Font.Bold = False
'''    gt.LevelDefs(3).ColumnDefs.ColumnCaptions = gtColumnCaptionsTrue
'''    gt.LevelDefs(3).ColumnDefs.Add , , "#"
'''    gt.LevelDefs(3).ColumnDefs.Add , , "Time"
'''    gt.LevelDefs(3).ColumnDefs.Add , , "Operator"
'''    gt.LevelDefs(3).ColumnDefs.Add , , "Sale value"
'''    gt.LevelDefs(3).ColumnDefs.Add , , "VAT"
'''    gt.LevelDefs(3).ColumnDefs.Add , , "Change"
'''    gt.LevelDefs(3).ColumnDefs.Add , , "Type"
'''    gt.LevelDefs(3).ColumnDefs.Add , , "Customer"
'''    gt.LevelDefs(3).ColumnDefs(0).Width = 600
'''    gt.LevelDefs(3).ColumnDefs(1).Width = 1000
'''    gt.LevelDefs(3).ColumnDefs(2).Width = 1000
'''    gt.LevelDefs(3).ColumnDefs(3).Width = 1000
'''    gt.LevelDefs(3).ColumnDefs(4).Width = 600
'''    gt.LevelDefs(3).ColumnDefs(5).Width = 1000
'''    gt.LevelDefs(3).ColumnDefs(6).Width = 1000
'''    gt.LevelDefs(3).ColumnDefs(7).Width = 1000
'''
'''
'''    'We now set the default width for each column.  This
'''    'width can be changed by the user by placing the
'''    'mouse pointer between two column captions, holding
'''    'the mouse button down, and moving left or right
'''    gt.LevelDefs(3).ColumnDefs(0).Width = 600
'''    gt.LevelDefs(3).ColumnDefs(1).Width = 1300
'''    gt.LevelDefs(3).ColumnDefs(2).Width = 1000
'''    gt.LevelDefs(3).ColumnDefs(3).Width = 700
'''
''''    'The alignment of the text in each column has to be adjusted
''''    'as well since the data are different types.  Ranks are
''''    'centered, and numbers (or in this case the score) are
''''    'right justified
''''    gt.LevelDefs(3).ColumnDefs(0).TextAlignment = gtTextAlignmentCenterMiddle
''''    gt.LevelDefs(3).ColumnDefs(3).TextAlignment = gtTextAlignmentRightMiddle
''''
''''    'The captions for the rank and the score column is aligned
''''    'to match the data that is listed below it.
''''    gt.LevelDefs(3).ColumnDefs(0).CaptionTextAlignment = gtTextAlignmentCenterTop
''''    gt.LevelDefs(3).ColumnDefs(3).CaptionTextAlignment = gtTextAlignmentRightTop
''''
''''    'The Pictures in the Rank Column are aligned to
''''    'the left
''''    gt.LevelDefs(3).ColumnDefs(0).PictureAlignment = gtPictureAlignmentLeftMiddle
'''
'''    'Add fifteen nodes.  We use constants defined by the control
'''    'to specify the relationship between nodes.  To get these
'''    'constants, hit the 'F2' key to bring up the object browser,
'''    'select datatree, and find the Relationship constants
'''    gt.Nodes.Add , , "TitleKey", "1996 Atlanta Olympics", 1, 1
'''    gt.Nodes.Add "TitleKey", gtRelationshipChild, "BicyclingKey", "Bicycling", 2, 2
'''    gt.Nodes.Add "TitleKey", gtRelationshipChild, "GymnasticsKey", "Gymnastics", 3, 3
'''    gt.Nodes.Add "GymnasticsKey", gtRelationshipChild, "FloorExerciseKey", "Floor Exercise", 0, 0
'''    gt.Nodes.Add "GymnasticsKey", gtRelationshipChild, "VaultKey", "Vault", 0, 0
'''    gt.Nodes.Add "GymnasticsKey", gtRelationshipChild, "BalanceBeamKey", "Balance Beam", 0, 0
'''    gt.Nodes.Add "GymnasticsKey", gtRelationshipChild, "UnevenParallelBarsKey", "Uneven Parallel Bars", 0, 0
'''    gt.Nodes.Add "TitleKey", gtRelationshipChild, "RowingKey", "Rowing", 4, 4
'''    gt.Nodes.Add "BalanceBeamKey", gtRelationshipChild, "FirstPlaceKey", "1", 7, 7
'''    gt.Nodes.Add "BalanceBeamKey", gtRelationshipChild, "SecondPlaceKey", "2", 8, 8
'''    gt.Nodes.Add "BalanceBeamKey", gtRelationshipChild, "ThirdPlaceKey", "3", 9, 9
'''    gt.Nodes.Add "BalanceBeamKey", gtRelationshipChild, "FourthPlaceKey", "4"
'''    gt.Nodes.Add "BalanceBeamKey", gtRelationshipChild, "TiedForFourthKey", "4"
'''    gt.Nodes.Add "TitleKey", gtRelationshipChild, "SwimmingKey", "Swimming", 10, 10
'''    gt.Nodes.Add "TitleKey", gtRelationshipChild, "WaterPoloKey", "Water Polo", 11, 11
'''
'''    'Nodes added in LevelDef 3 automatically have four subitem
'''    'objects associated with them.  We now populate the data
'''    'in those objects by setting the text and image properties
'''    gt.Nodes("FirstPlaceKey").SubItems(1).Text = " E. Shore"
'''    gt.Nodes("FirstPlaceKey").SubItems(2).Text = " U.S.A."
'''    gt.Nodes("FirstPlaceKey").SubItems(3).Text = "9.985 "
'''    gt.Nodes("FirstPlaceKey").SubItems(2).Image = 12
'''
'''    gt.Nodes("SecondPlaceKey").SubItems(1).Text = " A. Bruchhauser"
'''    gt.Nodes("SecondPlaceKey").SubItems(2).Text = " Germany"
'''    gt.Nodes("SecondPlaceKey").SubItems(3).Text = "9.980 "
'''    gt.Nodes("SecondPlaceKey").SubItems(2).Image = 13
'''
'''    gt.Nodes("ThirdPlaceKey").SubItems(1).Text = " B. Walton"
'''    gt.Nodes("ThirdPlaceKey").SubItems(2).Text = " Canada"
'''    gt.Nodes("ThirdPlaceKey").SubItems(3).Text = "9.975 "
'''    gt.Nodes("ThirdPlaceKey").SubItems(2).Image = 16
'''
'''    gt.Nodes("FourthPlaceKey").SubItems(1).Text = " A. Bakker"
'''    gt.Nodes("FourthPlaceKey").SubItems(2).Text = " U.K."
'''    gt.Nodes("FourthPlaceKey").SubItems(3).Text = "9.950 "
'''    gt.Nodes("FourthPlaceKey").SubItems(2).Image = 14
'''
'''    gt.Nodes("TiedForFourthKey").SubItems(1).Text = " C. Kasparov"
'''    gt.Nodes("TiedForFourthKey").SubItems(2).Text = " Russia"
'''    gt.Nodes("TiedForFourthKey").SubItems(3).Text = "9.950 "
'''    gt.Nodes("TiedForFourthKey").SubItems(2).Image = 15
'''
'''    'To show the user more of the tree by default, we expand the Title Node
'''    gt.Nodes("TitleKey").Expanded = True
'''
'''End Sub
