VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Begin VB.Form frmPreDeliveryAdvice 
   Caption         =   "Action supplier status change"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lbCOLAction 
      Appearance      =   0  'Flat
      Height          =   3345
      Left            =   4815
      TabIndex        =   18
      Top             =   4845
      Width           =   2865
   End
   Begin VB.CheckBox chkETA 
      Caption         =   "set new ETA"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   9330
      TabIndex        =   16
      Top             =   6375
      Width           =   1680
   End
   Begin MSComCtl2.DTPicker dtpETA 
      Height          =   315
      Left            =   7785
      TabIndex        =   15
      Top             =   6330
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      _Version        =   393216
      Format          =   244973569
      CurrentDate     =   40023
   End
   Begin VB.ListBox lbStatus 
      Appearance      =   0  'Flat
      Height          =   3345
      Left            =   135
      TabIndex        =   12
      Top             =   4845
      Width           =   4530
   End
   Begin VB.TextBox txtSuppMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   360
      Left            =   7785
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   5010
      Width           =   3270
   End
   Begin VB.TextBox txtDiarize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   8250
      TabIndex        =   8
      Top             =   7110
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox txtCustMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   360
      Left            =   7785
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   5685
      Visible         =   0   'False
      Width           =   3285
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
      Left            =   7905
      Picture         =   "frmPreDeliveryAdvice.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7560
      Width           =   1545
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   9420
      Picture         =   "frmPreDeliveryAdvice.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7545
      Width           =   1515
   End
   Begin TrueOleDBGrid60.TDBGrid POGrid 
      Height          =   1125
      Left            =   150
      OleObjectBlob   =   "frmPreDeliveryAdvice.frx":0714
      TabIndex        =   0
      Top             =   1680
      Width           =   10905
   End
   Begin TrueOleDBGrid60.TDBGrid COGrid 
      Height          =   1335
      Left            =   150
      OleObjectBlob   =   "frmPreDeliveryAdvice.frx":55C7
      TabIndex        =   1
      Top             =   3210
      Width           =   10890
   End
   Begin TrueOleDBGrid60.TDBGrid GP 
      Height          =   1110
      Left            =   165
      OleObjectBlob   =   "frmPreDeliveryAdvice.frx":B2BE
      TabIndex        =   13
      Top             =   150
      Width           =   10890
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Availability status"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   225
      TabIndex        =   17
      Top             =   4590
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Revised ETA"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   7815
      TabIndex        =   14
      Top             =   6105
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Message from supplier"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   7800
      TabIndex        =   11
      Top             =   4755
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter date or period from today. e.g. 23/4/2008 or 2w or 3m etc."
      ForeColor       =   &H8000000D&
      Height          =   480
      Left            =   10140
      TabIndex        =   9
      Top             =   7065
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Message to customers"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   7785
      TabIndex        =   7
      Top             =   5445
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier orders outstanding"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   210
      TabIndex        =   3
      Top             =   1425
      Width           =   2460
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer orders outstanding"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   195
      TabIndex        =   2
      Top             =   2955
      Width           =   2460
   End
End
Attribute VB_Name = "frmPreDeliveryAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oProd As a_Product
Dim XC As XArrayDB  'OSPOs
Dim XD As XArrayDB  'OSCOs
Dim XPROS As XArrayDB  'OSCOs
'Dim XIN As XArrayDB
Dim rs As ADODB.Recordset
Dim POLS As ADODB.Recordset
Dim COLS As ADODB.Recordset
Dim PROS As ADODB.Recordset
Dim OpenResult As Integer
Dim dteRediarize As Date
Dim strSignature As String
Dim oSM As New z_StockManager
Dim strProductStatus As String
Dim strCOLAction As String
Dim Mode As String
Dim sOldStatus As String

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = False
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuCancelINactive.Enabled = False
    Forms(0).mnuFulfil.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Forms(0).mnuEmail.Enabled = False
    Forms(0).mnuOutlook.Enabled = False

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.SetMenu"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.SetMenu"
End Sub

Public Sub component(x As String, PID As String, XMLType As String, OldStatus As String)
    On Error GoTo errHandler
    Dim i As Integer
    Dim iFound As Integer
    
    Mode = XMLType
    Set XC = New XArrayDB
    Set XD = New XArrayDB
    Set XPROS = New XArrayDB
    sOldStatus = OldStatus
    oSM.LoadForPreDelAdvice x, POLS, COLS, PROS, PID, XMLType
    LoadPOs
    LoadCOs
    LoadPROS
    For i = 1 To oPC.Configuration.ProductStatus.Count
        If Left(oPC.Configuration.ProductStatus.ItemByOrdinalIndex(i), 4) = "(" & FNS(XPROS.Value(1, 3)) & ")" Then
            iFound = i
            Exit For
        End If
    Next
    If iFound > 0 Then
        Me.lbStatus.ListIndex = iFound - 1
    End If
  '  oPC.Configuration.ProductStatus
 '   me.lbStatus.ListIndex =
    SetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.Component(x)", x
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.component(x,PID,XMLType,OldStatus)", Array(x, PID, XMLType, _
         OldStatus)
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
        strProductStatus = ""

        Me.Hide
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.cmdClose_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub clearform()
    On Error GoTo errHandler
    Me.chkETA = 0
    Me.txtDiarize = ""
    XC.Clear
    XD.Clear
    POGrid.ReBind
    COGrid.ReBind

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.clearform"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.clearform"
End Sub
Public Function GetNewStatus() As String
    On Error GoTo errHandler
    GetNewStatus = strProductStatus
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.GetNewStatus"
End Function
Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim i As Integer
Dim oSM As New z_StockManager
Dim ar As New arCustReport2
Dim bOK As Boolean
Dim lngPaid As Long
Dim f As New frmTrackingActions
Dim rs As ADODB.Recordset
Dim oSQL As New z_SQL
Dim strPOLS As String
Dim strCOLS As String
Dim strPROS As String
Dim xMLDoc As ujXML
Dim XMLArgs As String
Dim strPos As String
Dim lngPSCID As Long
    If lbStatus > "" Then
        If MsgBox("You are setting the status of " & IIf(Mode = "R" Or Mode = "", CStr(PROS.RecordCount), CStr(POLS.RecordCount)) & " book" & IIf(PROS.RecordCount = 1, "", "s") & " to """ & lbStatus & """" & vbCrLf _
                    & IIf(chkETA = 1, "and are setting the ETA to " & Format(dtpETA.Value, "dd/mm/yyyy"), ""), vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
            Exit Sub
        End If
    Else
        If chkETA = 1 Then
            If MsgBox("You are setting the ETA to " & Format(dtpETA.Value, "dd/mm/yyyy"), vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    If oPC.Configuration.SignTransactions = True Then
    strPos = "Pos 03.1"
        If SecurityControl(enSECURITY_PO_SIGN, , "Sign this supplier report", DOCAPPROVAL, , , , strSignature) = False Then
                strPos = "Pos 3.2"
               Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
'
    If Not POLS Is Nothing Then
        If Not POLS.BOF Then POLS.MoveFirst
        If Not POLS.eof Then
         '   POLS.MoveFirst
            Set xMLDoc = New ujXML
            With xMLDoc
                .docProgID = "MSXML2.DOMDocument"
                .docInit "doc_TrackingActions"
                    .chCreate "MessageType"
                        .elText = "doc_TrackingActions"
                    .elCreateSibling "DetailLines", True
                    For i = 1 To POLS.RecordCount
                            .chCreate "ITEM"
                            .chCreate "POLID"
                                .elText = FNS(POLS.fields("POL_ID"))
                            .navUP
                            .navUP
                            POLS.MoveNext
                    Next i
                 strPOLS = .docXML
            End With
        End If
    End If
    Set xMLDoc = Nothing
    If Not COLS Is Nothing Then
        If Not COLS.BOF Then COLS.MoveFirst
        If Not COLS.eof Then
          '  COLS.MoveFirst
            Set xMLDoc = New ujXML
            With xMLDoc
                .docProgID = "MSXML2.DOMDocument"
                .docInit "doc_TrackingActions"
                    .chCreate "MessageType"
                        .elText = "doc_TrackingActions"
                    .elCreateSibling "DetailLines", True
                    For i = 1 To COLS.RecordCount
                            .chCreate "ITEM"
                            .chCreate "COLID"
                                .elText = FNS(COLS.fields("COL_ID"))
                            .navUP
                            .navUP
                            COLS.MoveNext
                    Next i
                 strCOLS = .docXML
            End With
        End If
    End If
'    oSM.Action_POLSet strPOLS, strCOLS, oPC.WorkstationName, dteRediarize, Me.txtSuppMsg, txtCustMsg, GetStatus, lngPaid, strSignature
    Set xMLDoc = Nothing
    If Not PROS Is Nothing Then
        If Not PROS.BOF Then PROS.MoveFirst
        If Not PROS.eof Then
            Set xMLDoc = New ujXML
            With xMLDoc
                .docProgID = "MSXML2.DOMDocument"
                .docInit "doc_TrackingActions"
                    .chCreate "MessageType"
                        .elText = "doc_TrackingActions"
                    .elCreateSibling "DetailLines", True
                    For i = 1 To PROS.RecordCount
                            .chCreate "ITEM"
                            .chCreate "PID"
                            'changed from 3 to 2
                                .elText = FNS(PROS.fields(3))
                            .navUP
                            .navUP
                            PROS.MoveNext
                    Next i
                 strPROS = .docXML
            End With
        End If
    End If
    For i = 0 To lbStatus.ListCount - 1
        If lbStatus.Selected(i) Then
            strProductStatus = lbStatus.List(i)
            Exit For
        End If
    Next
    For i = 0 To lbCOLAction.ListCount - 1
        If lbCOLAction.Selected(i) Then
            strCOLAction = lbCOLAction.List(i)
            Exit For
        End If
    Next
    
    oSM.ProductStatusChange strPROS, strPOLS, strCOLS, oPC.WorkstationName, txtSuppMsg, oPC.Configuration.ProductStatus.Key(strProductStatus), _
                    IIf(chkETA = 1, Me.dtpETA, CDate(0)), strSignature, lngPSCID, oPC.Configuration.COActions.Key(strCOLAction)
    Set xMLDoc = Nothing
'
    clearform
    Set oProd = Nothing
    cmdOK.Enabled = Not (oProd Is Nothing)
    Screen.MousePointer = vbDefault
    
    Me.Hide
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.cmdOK_Click", , , , "strPOS", Array(strPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub
Public Function SetETA(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim dteTemp As Date
    bOK = True
    If IsDate(val) Then
       dteRediarize = val
    Else
        bOK = SetField_DIARYPERIODS(dteRediarize, val, "ETA", 1)
    End If
    SetETA = bOK
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.SetETA(val)", val
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.SetETA(val)", val
End Function
Private Sub LoadPOs()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    If POLS.RecordCount = 0 Then Exit Sub
    XC.Clear
    XC.ReDim 1, POLS.RecordCount, 1, 10
'    For i = 1 To POGrid.Columns.Count
'        POGrid.Columns(i - 1).Width = GetSetting("PBKS", "frmPreDeliveryAdviceA", CStr(i), POGrid.Columns(i - 1).Width)
'    Next
    POLS.MoveFirst
    lngIndex = 1
    Do While Not POLS.eof
'        XC.Value(lngIndex, 1) = FNS(POLS.Fields("CodeF"))
'        XC.Value(lngIndex, 2) = FNS(POLS.Fields("P_Title"))
        XC.Value(lngIndex, 1) = FNS(POLS.fields("POL_Ref"))
        XC.Value(lngIndex, 2) = FNS(POLS.fields("QtyFIRM"))
        XC.Value(lngIndex, 3) = FNS(POLS.fields("QtySS"))
        XC.Value(lngIndex, 4) = FNS(POLS.fields("QtyREC"))
        XC.Value(lngIndex, 5) = FNS(POLS.fields("POL_ETA"))
        POLS.MoveNext
        lngIndex = lngIndex + 1
    Loop
    XC.QuickSort 1, POLS.RecordCount, 2, XORDER_DESCEND, XTYPE_STRING
    POGrid.Array = XC
    POGrid.ReBind

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.LoadPOs"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.LoadPOs"
End Sub
Private Sub LoadPROS()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    If PROS.RecordCount = 0 Then Exit Sub
    XPROS.Clear
    XPROS.ReDim 1, PROS.RecordCount, 1, 11
'    For i = 1 To GP.Columns.Count
'        GP.Columns(i - 1).Width = GetSetting("PBKS", "frmPreDeliveryAdviceB", CStr(i), GP.Columns(i - 1).Width)
'    Next
    PROS.MoveFirst
    lngIndex = 1
    Do While Not PROS.eof
        XPROS.Value(lngIndex, 1) = FNS(PROS.fields(0))
        XPROS.Value(lngIndex, 2) = FNS(PROS.fields(1))
        XPROS.Value(lngIndex, 3) = FNS(PROS.fields(2))
       ' XPROS.Value(lngIndex, 4) = FNS(PROS.Fields(3))
        PROS.MoveNext
        lngIndex = lngIndex + 1
   Loop
    XPROS.QuickSort 1, PROS.RecordCount, 1, XORDER_DESCEND, XTYPE_STRING
    GP.Array = XPROS
    GP.ReBind

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.LoadPROS"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.LoadPROS"
End Sub

Private Sub LoadCOs()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    If COLS.RecordCount = 0 Then Exit Sub
    XD.Clear
    XD.ReDim 1, COLS.RecordCount, 1, 11
'    For i = 1 To COGrid.Columns.Count
'        COGrid.Columns(i - 1).Width = GetSetting("PBKS", "frmPreDeliveryAdviceB", CStr(i), COGrid.Columns(i - 1).Width)
'    Next
    COLS.MoveFirst
    lngIndex = 1
    Do While Not COLS.eof
        XD.Value(lngIndex, 1) = FNS(COLS.fields("Customer"))
        XD.Value(lngIndex, 2) = FNS(COLS.fields("DocCode"))
        XD.Value(lngIndex, 3) = FNS(COLS.fields("DocDate"))
        XD.Value(lngIndex, 4) = FNS(COLS.fields("QTY"))
        XD.Value(lngIndex, 5) = FNS(COLS.fields("QtyDisp"))
        XD.Value(lngIndex, 6) = FNS(COLS.fields("COL_ETA"))
        XD.Value(lngIndex, 7) = FNS(COLS.fields("CodeF"))
        XD.Value(lngIndex, 8) = FNS(COLS.fields("P_Title"))
        COLS.MoveNext
        lngIndex = lngIndex + 1
   Loop
    XD.QuickSort 1, COLS.RecordCount, 1, XORDER_DESCEND, XTYPE_STRING
    COGrid.Array = XD
    COGrid.ReBind

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.LoadCOs"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.LoadCOs"
End Sub
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.POGrid, "frmPreDeliveryAdviceA"
    SaveLayout Me.COGrid, "frmPreDeliveryAdviceB"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.mnuSaveLayout"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.mnuSaveLayout"
End Sub


Private Sub COGrid_LostFocus()
    On Error GoTo errHandler
    COGrid.Update
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.COGrid_LostFocus"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.COGrid_LostFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub dtpETA_Change()
    On Error GoTo errHandler
    Me.chkETA = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.dtpETA_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub dtpETA_Click()
    On Error GoTo errHandler
 '   Me.chkETA = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.dtpETA_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.Form_Activate"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.Form_Deactivate"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim i As Integer

    SetGridLayout Me.GP, Me.Name
    SetGridLayout Me.POGrid, "frmPreDeliveryAdviceA"
    SetGridLayout Me.COGrid, "frmPreDeliveryAdviceB"
    SetFormSize Me


    LoadListbox lbStatus, oPC.Configuration.ProductStatus
    lbStatus.ListIndex = -1
    For i = 0 To lbStatus.ListCount - 1
        If lbStatus.List(i) = sOldStatus Then
            lbStatus.Selected(i) = True
        End If
    Next
    LoadListbox lbCOLAction, oPC.Configuration.COActions
    lbCOLAction.ListIndex = -1
    For i = 0 To lbStatus.ListCount - 1
        If lbStatus.List(i) = sOldStatus Then
            lbStatus.Selected(i) = True
        End If
    Next
    dtpETA.Value = DateAdd("m", 1, Date)
    Me.cmdOK.Enabled = True
    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.Form_Load"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtMsg_Change()
    On Error GoTo errHandler
Dim strArg As String
Dim iStart As Integer
Dim iEnd As Integer
Dim oU As New z_UTIL
Dim strResult As String
Dim f As frmFindTextBite

    iStart = 0
    iEnd = 0
    iStart = InStr(1, txtCustMsg, "?") + 1
    If iStart = 0 Then Exit Sub
    strResult = ""
    iEnd = InStr(iStart, txtCustMsg, "?")
    If iStart > 0 And iEnd > iStart Then
        strArg = Trim(Mid(txtCustMsg, iStart, iEnd - iStart))
        strResult = oU.GetTextBite(strArg)
        If strResult > "" Then
                txtCustMsg = Replace(txtCustMsg, "?" & strArg & "?", strResult)
        End If
    Else
    End If

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.txtMsg_Change"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.txtMsg_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    SaveLayout Me.GP, Me.Name, Me.Height, Me.Width
    SaveLayout Me.POGrid, "frmPreDeliveryAdviceA"
    SaveLayout Me.COGrid, "frmPreDeliveryAdviceB"
    XC.Clear
    XD.Clear
    XPROS.Clear
    'XIN.Clear
    Set XC = Nothing
    Set XD = Nothing
    Set XPROS = Nothing
   ' Set XIN = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtCustMsg_Change()
    On Error GoTo errHandler
    txtCustMsg = HandleTextWithBites(txtCustMsg)
    Me.cmdOK.Enabled = (txtCustMsg > "" And COLS.RecordCount > 0) Or (Me.chkETA = 1) Or (Me.txtSuppMsg > "" And POLS.RecordCount > 0)
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.txtCustMsg_Change"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.txtCustMsg_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtSuppMsg_Change()
    On Error GoTo errHandler
    txtSuppMsg = HandleTextWithBites(txtSuppMsg)
  '  Me.cmdOK.Enabled = (txtCustMsg > "" And COLS.RecordCount > 0) Or (Me.chkETA = 1) Or (Me.txtSuppMsg > "" And POLS.RecordCount > 0)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPreDeliveryAdvice.txtSuppMsg_Change"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPreDeliveryAdvice.txtSuppMsg_Change", , EA_NORERAISE
    HandleError
End Sub
