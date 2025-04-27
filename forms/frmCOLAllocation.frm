VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCOLAllocation 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Reserve items for customers"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12840
   ControlBox      =   0   'False
   FillColor       =   &H00FFC0FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   12840
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   60
      TabIndex        =   22
      Top             =   5760
      Width           =   405
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   60
      TabIndex        =   18
      Top             =   4980
      Width           =   405
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   60
      TabIndex        =   17
      Top             =   5235
      Width           =   405
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   60
      TabIndex        =   16
      Top             =   5490
      Width           =   405
   End
   Begin VB.CommandButton cmdExcelExport 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Exp"
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
      Left            =   5730
      Picture         =   "frmCOLAllocation.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Sets all the allocations to those shown when the form was opened"
      Top             =   5415
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.Frame frmGenerate 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Create documents"
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
      Height          =   2115
      Left            =   9975
      TabIndex        =   10
      Top             =   4230
      Width           =   2115
      Begin VB.CommandButton cmdGenGDNs 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Create documents"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   480
         Picture         =   "frmCOLAllocation.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Creates invoices for all products/customer orderlines where qty > 0"
         Top             =   1185
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton cmdGenAppros 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Create appros"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   465
         Picture         =   "frmCOLAllocation.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Creates invoices for all products/customer orderlines where qty > 0"
         Top             =   1170
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton cmdGenInv 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Create invoices"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   465
         Picture         =   "frmCOLAllocation.frx":0A9E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Creates invoices for all products/customer orderlines where qty > 0"
         Top             =   1170
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Use customer disc && list price"
         ForeColor       =   &H8000000D&
         Height          =   510
         Left            =   75
         TabIndex        =   12
         Top             =   600
         Width           =   1830
      End
      Begin VB.OptionButton optCO 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Use C.O. disc && price"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   75
         TabIndex        =   11
         Top             =   270
         Value           =   -1  'True
         Width           =   1830
      End
   End
   Begin VB.CommandButton cmdCheckAll 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Check all"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Sets all the allocations to those shown when the form was opened"
      Top             =   4770
      Width           =   1140
   End
   Begin VB.CommandButton cmdUncheck 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Uncheck all"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7470
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Sets all the allocations to those shown when the form was opened"
      Top             =   4770
      Width           =   1140
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
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
      Left            =   6750
      Picture         =   "frmCOLAllocation.frx":0E28
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Sets all the allocations to those shown when the form was opened"
      Top             =   5415
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox txtResult 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8E8E8&
      ForeColor       =   &H000000C0&
      Height          =   3750
      Left            =   9885
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      ToolTipText     =   "These products have too many allocations, i.e. there is not sufficient stock to meet the allocations. Reduce the allocations."
      Top             =   390
      Width           =   2295
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   270
      Left            =   3060
      TabIndex        =   3
      Top             =   5730
      Visible         =   0   'False
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdQuit 
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
      Left            =   7770
      Picture         =   "frmCOLAllocation.frx":11B2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Sets all the allocations to those shown when the form was opened (removes locks)"
      Top             =   5415
      Width           =   1000
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Save"
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
      Left            =   8790
      Picture         =   "frmCOLAllocation.frx":153C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close this form without creating invoices (keeps allocations and locks)"
      Top             =   5415
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   4590
      Left            =   165
      OleObjectBlob   =   "frmCOLAllocation.frx":18C6
      TabIndex        =   0
      Top             =   90
      Width           =   9615
   End
   Begin VB.Label lblGrey 
      BackColor       =   &H00BEE0C7&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer blocked"
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
      Left            =   525
      TabIndex        =   23
      Top             =   5715
      Width           =   2580
   End
   Begin VB.Label lblpink 
      BackColor       =   &H00BEE0C7&
      BackStyle       =   0  'Transparent
      Caption         =   "Allocation > Qty on hand"
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
      Height          =   300
      Left            =   540
      TabIndex        =   21
      Top             =   4935
      Width           =   2580
   End
   Begin VB.Label lblGreen 
      BackColor       =   &H00BEE0C7&
      BackStyle       =   0  'Transparent
      Caption         =   "Allocation < ordered"
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
      Height          =   300
      Left            =   525
      TabIndex        =   20
      Top             =   5190
      Width           =   2580
   End
   Begin VB.Label lblRed 
      BackColor       =   &H00BEE0C7&
      BackStyle       =   0  'Transparent
      Caption         =   "Allocation > ordered"
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
      Left            =   525
      TabIndex        =   19
      Top             =   5445
      Width           =   2580
   End
   Begin VB.Label lblItalic 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Italic text means other orders exist"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   210
      TabIndex        =   15
      Top             =   4680
      Width           =   3270
   End
   Begin VB.Label lblItemsCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   3450
      TabIndex        =   7
      Top             =   4710
      Width           =   3720
   End
   Begin VB.Label lblOverComm 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Over-committed products"
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
      Height          =   390
      Left            =   9765
      TabIndex        =   4
      Top             =   120
      Width           =   2265
   End
End
Attribute VB_Name = "frmCOLAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oRS As ADODB.Recordset
Dim WithEvents cALLOC As chex_COLAllocation
Attribute cALLOC.VB_VarHelpID = -1
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim iRecs As Integer
Dim lngArrayRows As Long
Dim lngDELID As Long
Dim rs As New ADODB.Recordset
Dim strType As String
Dim bOKtoClose As Boolean
Dim roProduct As a_Product
Dim bSkipResize As Boolean
Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayout Me.Grid1, Me.Name, Me.Height, Me.Width
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.mnuSaveLayout"
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
    'Forms(0).mnuInvAdd.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.SetMenu"
End Sub

Private Sub PrepareNewSlate()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    Set cALLOC = Nothing
    Set cALLOC = New chex_COLAllocation
    cALLOC.GenerateCOLALLOCationset lngDELID 'INSERT INTO tCOLAlloc
    cALLOC.Load lngDELID
    LoadGrid
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.PrepareNewSlate"
End Sub

Private Sub cALLOC_Valid(pResult As String)
    On Error GoTo errHandler
'MsgBox pOK
    EnableOK (pResult = "")
    txtResult = pResult
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.cALLOC_Valid(pResult)", pResult, EA_NORERAISE
    HandleError
End Sub
Public Sub component(pcALLOC As chex_COLAllocation, pType As String, pViewOnly As Boolean, GenerateDocumentType As String)
    On Error GoTo errHandler
    strType = pType
    If pType = "DELIVERY" Then
        If pViewOnly Then
            Me.Caption = "Allocate delivered items to outstanding customer orders"
            Grid1.FetchRowStyle = False
            Me.cmdGenInv.Visible = False
            Me.cmdSave.Visible = False
            Me.cmdQuit.Visible = True
            Me.Grid1.Columns.Item(8).Visible = False
            Me.cmdPrint.Visible = True
            Me.cmdExcelExport.Visible = Me.cmdPrint.Visible
        Else
            Me.cmdGenInv.Visible = False
            Me.cmdSave.Visible = True
            Me.cmdQuit.Visible = False
            Me.cmdPrint.Visible = False
        End If
    Else
        Me.cmdGenInv.Visible = True
        Me.cmdSave.Visible = True
        Me.cmdQuit.Visible = True
        Me.cmdPrint.Visible = False
        Me.cmdExcelExport.Visible = True
        Me.Caption = "Fulfil customer orders from stock"
    End If
    Select Case GenerateDocumentType
        Case "I"
            cmdGenAppros.Visible = False
            cmdGenGDNs.Visible = False
            cmdGenInv.Visible = True
        Case "A"
            cmdGenAppros.Visible = True
            cmdGenGDNs.Visible = False
            cmdGenInv.Visible = False
        Case "G"
            cmdGenAppros.Visible = False
            cmdGenGDNs.Visible = True
            cmdGenInv.Visible = False
    End Select
    Set cALLOC = pcALLOC
    LoadGrid
    lblItemsCount.Caption = cALLOC.Count & " items"
    cALLOC.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.component(pcALLOC,pType,pViewOnly)", Array(pcALLOC, pType, pViewOnly)
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim tmp As String
Dim qtyRecs As Long
Dim lngAwaiting As Long
Dim lngAllocation As Long
Dim lngAvailableToAllocate As Long
Dim i As Integer
Dim oALLOC As a_COLAllocation
Dim lngFirm As Long
Dim lngSS As Long
Dim AllocFirm As Long
Dim AllocSS As Long
Dim strWSName As String

    strWSName = oPC.NameOfPC

    i = 0
    Set XA = New XArrayDB
    XA.Clear
    iRecs = i
    lngIndex = 1
    lngArrayRows = cALLOC.Count
    XA.ReDim 1, lngArrayRows, 1, 25
    Set rs = New ADODB.Recordset
    rs.Fields.Append "PID", adVarChar, 40, adFldKeyColumn
    rs.Fields.Append "Bal", adInteger
    rs.Open
    For Each oALLOC In cALLOC
        lngAvailableToAllocate = oALLOC.QtyOnHand
        
        XA.Value(lngIndex, 17) = oALLOC.PID
        rs.Find ("PID = " & oALLOC.PID)
        
        If rs.eof Then
            rs.AddNew
                rs.Fields("PID") = oALLOC.PID
                rs.Fields("BAL") = oALLOC.QtyOnHand
            rs.Update
        End If
        rs.Update
        XA.Value(lngIndex, 1) = oALLOC.ProductOH & "/" & oALLOC.ProductRES & " - " & oALLOC.TitleShort(40)
        XA.Value(lngIndex, 2) = oALLOC.code
        XA.Value(lngIndex, 3) = oALLOC.CustomerAcnoName
        XA.Value(lngIndex, 4) = oALLOC.OrderCode
        XA.Value(lngIndex, 5) = oALLOC.OrderDateF
        AllocFirm = oALLOC.AllocatedQty
        AllocSS = oALLOC.AllocatedQtySS
        AllocateQtys FNN(rs.Fields("BAL")), oALLOC.OrderedQty, IIf(oPC.AllowsSSInvoicing, oALLOC.DeliveredSoFarFirm, oALLOC.DeliveredSoFar), oALLOC.OrderedSSQty, oALLOC.DeliveredSoFarSS, AllocFirm, AllocSS, oALLOC.CustomerBlocked
        oALLOC.AllocatedQty = AllocFirm
        oALLOC.AllocatedQtySS = AllocSS
        If oPC.AllowsSSInvoicing Then
            XA.Value(lngIndex, 6) = oALLOC.OrderedQty & "/" & oALLOC.OrderedSSQty
            XA.Value(lngIndex, 7) = oALLOC.DeliveredSoFarFirm & "/" & oALLOC.DeliveredSoFarSS
            XA.Value(lngIndex, 9) = oALLOC.AllocatedQty & "/" & oALLOC.AllocatedQtySS
        Else
            XA.Value(lngIndex, 6) = oALLOC.OrderedQty
            XA.Value(lngIndex, 7) = oALLOC.DeliveredSoFar
            XA.Value(lngIndex, 9) = oALLOC.AllocatedQty
        End If
        XA.Value(lngIndex, 8) = IIf(strWSName <> oALLOC.WSLock, "**Locked by other workstation:" & oALLOC.WSLock, oALLOC.OrderDetails)
        'If strWSName <> oALLOC.WSLock Then MsgBox "TEST"
        rs.Fields("BAL") = rs.Fields("BAL") - (oALLOC.AllocatedQty + oALLOC.AllocatedQtySS)
        rs.Update
        
        oALLOC.BeginEdit
        'oALLOC.SetACtionYN True
        oALLOC.SetACtionYN (FNN(oALLOC.AllocatedQty) + FNN(oALLOC.AllocatedQtySS)) > 0
        
        oALLOC.ApplyEdit
        If strWSName <> oALLOC.WSLock Then
            XA.Value(lngIndex, 10) = 0
        Else
            XA.Value(lngIndex, 10) = IIf((FNN(oALLOC.AllocatedQty) + FNN(oALLOC.AllocatedQtySS)) > 0, -1, 0)
        End If
        XA.Value(lngIndex, 11) = IIf(oALLOC.UsesSubstitutions = "Y", "Y", "")
        XA.Value(lngIndex, 12) = oALLOC.COLID
        XA.Value(lngIndex, 13) = ""
        XA.Value(lngIndex, 14) = oALLOC.QtyOnHand
        XA.Value(lngIndex, 15) = oALLOC.key
     '   XA.Value(lngIndex, 17) = ""
        XA.Value(lngIndex, 18) = IIf(oALLOC.QtyonCO > XA.Value(lngIndex, 6), 1, 0)
        XA.Value(lngIndex, 19) = oALLOC.WSLock
        XA.Value(lngIndex, 23) = oALLOC.CustomerName
        XA.Value(lngIndex, 24) = oALLOC.CustomerBlocked
        XA.Value(lngIndex, 25) = oALLOC.Status
        
        
        If oALLOC.AllocatedQty > val(oALLOC.OrderedQty) - val(oALLOC.DeliveredSoFar) + val(oALLOC.DeliveredSoFarSS) Then
            MarkRowsValid 4, cALLOC(val(oALLOC.key)).key
        ElseIf oALLOC.AllocatedQty > val(oALLOC.QtyOnHand) Then
            MarkRowsValid 3, cALLOC(val(oALLOC.key)).key
        ElseIf oALLOC.AllocatedQty < val(oALLOC.OrderedQty) - val(oALLOC.DeliveredSoFar) + val(oALLOC.DeliveredSoFarSS) Then
            MarkRowsValid 2, cALLOC(val(oALLOC.key)).key
        ElseIf oALLOC.AllocatedQty = val(oALLOC.OrderedQty) - val(oALLOC.DeliveredSoFar) + val(oALLOC.DeliveredSoFarSS) Then
            MarkRowsValid 1, cALLOC(val(oALLOC.key)).key
        End If
        
        lngIndex = lngIndex + 1
    Next
    Set Grid1.Array = XA
    Grid1.ReBind
   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.LoadGrid"
End Sub
Private Sub AllocateQtys(Avail As Long, OrdFirm As Long, DelFirm As Long, OrdSS As Long, DelSS As Long, AllocFirm As Long, AllocSS As Long, CustomerBlocked As Boolean)
    On Error GoTo errHandler
Dim bal As Long
        If CustomerBlocked = False Then
        If oPC.AllowsSSInvoicing Then
            AllocFirm = GetMin(OrdFirm - DelFirm, Avail, False)
            bal = Avail - AllocFirm
            AllocSS = GetMin(OrdSS - DelSS, bal, False)
        Else
            AllocFirm = GetMin(OrdFirm - DelFirm, Avail, False)
        End If
        End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.AllocateQtys(Avail,OrdFirm,DelFirm,OrdSS,DelSS,AllocFirm,AllocSS)", _
         Array(Avail, OrdFirm, DelFirm, OrdSS, DelSS, AllocFirm, AllocSS)
End Sub
Private Sub cmdReset_Click()
    On Error GoTo errHandler
    If cALLOC.IsEditing Then cALLOC.CancelEdit
    Set cALLOC = Nothing
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.cmdReset_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdCheckAll_Click()
    On Error GoTo errHandler
Dim oALLOC As a_COLAllocation
Dim lngIndex As Long
Dim strWSName As String

    strWSName = oPC.NameOfPC

    lngIndex = 1
    For Each oALLOC In cALLOC
        If strWSName = oALLOC.WSLock Then
            oALLOC.BeginEdit
            oALLOC.SetACtionYN True
            oALLOC.ApplyEdit
            XA.Value(lngIndex, 10) = -1
        End If
        lngIndex = lngIndex + 1
    Next
    Grid1.Refresh

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.cmdCheckAll_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdGenAppros_Click()
    On Error GoTo errHandler
Dim oG As New z_OrderFulfilmentDocGen
Dim strError As String

    If MsgBox("You are about to generate appros. Continue?", vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    WaitMsg "Creating appros  . . .", True, Me
    
    cALLOC.Save strError, "NORMAL", "I"
    oG.GenerateApprosFromCOLALLOCs IIf(optCO = True, "CO", "C")
    
    WaitMsg "", False, Me
    bOKtoClose = True
    MsgBox "Appros have been generated. They are not yet issued." & vbCrLf & "Please check and issue each appro that has been generated.", vbInformation, "Please note"
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocationAppros.cmdGenAppros_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdGenGDNs_Click()
Dim oG As New z_OrderFulfilmentDocGen
Dim strError As String
Dim oIG As Z_InvoiceGeneration

    If MsgBox("You are about to generate delivery documents. Continue?", vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    WaitMsg "Creating documents  . . .", True, Me
    
    cALLOC.Save strError, "NORMAL", "G"
    oG.GenerateGDNsFromCOLALLOCs IIf(optCO = True, "CO", "C")
    If oPC.getProperty("GenerateInvoicesForGDNsAuto") = "TRUE" Then
        Set oIG = New Z_InvoiceGeneration
        oIG.GenerateInvoicesForGDNsIssuedFortCustomerWithoutCompleteRequiremen
    End If
    WaitMsg "", False, Me
    bOKtoClose = True
    MsgBox "Goods delivery notes have been generated. They are not yet issued." & vbCrLf & "Please check and issue each document that has been generated.", vbInformation, "Please note"
    Unload Me

End Sub

Private Sub cmdUncheck_Click()
    On Error GoTo errHandler
Dim oALLOC As a_COLAllocation
Dim lngIndex As Long

    lngIndex = 1
    For Each oALLOC In cALLOC
        oALLOC.BeginEdit
        oALLOC.SetACtionYN False
        oALLOC.ApplyEdit
        XA.Value(lngIndex, 10) = 0
        lngIndex = lngIndex + 1
    Next
    Grid1.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.cmdUncheck_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdGenInv_Click()
    On Error GoTo errHandler
Dim oG As New z_OrderFulfilmentDocGen
Dim strError As String

    If MsgBox("You are about to generate invoices. Continue?", vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    WaitMsg "Creating invoices  . . .", True, Me
    
    cALLOC.Save strError, "NORMAL", "I"
    oG.GenerateInvoicesFromCOLALLOCs IIf(optCO = True, "CO", "C")
    
    WaitMsg "", False, Me
    bOKtoClose = True
    MsgBox "Invoices have been generated. They are not yet issued." & vbCrLf & "Please check and issue each invoice that has been generated.", vbInformation, "Please note"
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.cmdGenInv_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdLoad_Click()
    On Error GoTo errHandler
Screen.MousePointer = vbHourglass
    Set cALLOC = New chex_COLAllocation
    cALLOC.Load 'oDEL.TRID
    LoadGrid
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.cmdLoad_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdclose_Click()
    On Error GoTo errHandler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    
Dim ar As New arCOLSFulfilled
    ar.component cALLOC
    ar.Show vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdExcelExport_Click()
    On Error GoTo errHandler
Dim sFile As String
Dim fs As New FileSystemObject
Dim strExecutable As String

    Screen.MousePointer = vbHourglass
    
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
        fs.CreateFolder oPC.SharedFolderRoot & "\TEMP"
    End If
    sFile = oPC.SharedFolderRoot & "\TEMP\OrderFulfilment.csv"
    If fs.FileExists(sFile) Then
        fs.DeleteFile sFile, True
    End If
    Me.Grid1.ExportToDelimitedFile sFile, , ","
    
    Screen.MousePointer = vbDefault
    If MsgBox("Spreadsheet file saved in: " & sFile & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
            strExecutable = GetPDFExecutable(sFile)
            F_7_AB_1_ShellAndWaitSimple strExecutable & " " & sFile, vbNormalFocus, 10000
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.cmdExcelExport_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdQuit_Click()
    On Error GoTo errHandler
    Dim strWSName As String
    If MsgBox("This will delete all details in the slate, confirm you want to close the form.", vbQuestion + vbOKCancel, "Warning") = vbCancel Then
        Exit Sub
    End If
    strWSName = oPC.NameOfPC
    
    cALLOC.ClearAllocations strWSName
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.cmdQuit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errHandler
Dim strError As String
Dim arReserve() As Reserve
Dim oCA As a_COLAllocation
Dim cALLOCOCS As c_COLsPerDEL
Dim rpt As arCOLSFulfilled

    If strType <> "DELIVERY" Then
        If MsgBox("You want to save the allocations without issuing invoices yet. These will be available until you next prepare an order allocation.", vbInformation + vbYesNo, "Confirm") = vbNo Then
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    Grid1.Update
    If Not cALLOC Is Nothing Then    'And strType = "DELIVERY"
        cALLOC.Save strError, strType, "I"
        lngDELID = cALLOC.DELID
        Set cALLOC = Nothing
        Set cALLOC = New chex_COLAllocation
        cALLOC.Load lngDELID, True
    End If
    If strType = "DELIVERY" Then
        Set rpt = New arCOLSFulfilled
        rpt.component cALLOC
        rpt.Printer.Orientation = ddOPortrait
        rpt.PrintReport False
    End If

    Screen.MousePointer = vbDefault

    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
    strType = "NORMAL"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    SetMenu
    bSkipResize = True
    If Me.WindowState <> 2 Then
        Top = 100
        Left = 100
        Height = 6700
        Width = 12400
    End If
    If GetSetting("PBKS", "COLALLOC", "DiscountFrom", "C") = "C" Then
        Me.optC = True
    Else
        Me.optCO = True
    End If

    bOKtoClose = False
    Grid1.Columns(8).BackColor = vbWhite
    bSkipResize = False
    SetGridLayout Me.Grid1, Me.Name
    SetFormSize Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
Dim lngDiffH As Long
    If bSkipResize Then Exit Sub
  '  lngDiffH = Grid1.Width
    Grid1.Width = NonNegative_Lng(Me.Width - 2900)
'    lngDiffH = Grid1.Width - lngDiffH
    txtResult.Left = NonNegative_Lng(Me.Width - 2600)
    frmGenerate.Left = NonNegative_Lng(Me.Width - 2600)
    lblOverComm.Left = NonNegative_Lng(Me.Width - 2600)
    lngDiff = Grid1.Height
    Grid1.Height = NonNegative_Lng(Me.Height - (Grid1.Top + 2400))
    lngDiff = (Grid1.Height - lngDiff)
    lblItalic.Top = lblItalic.Top + lngDiff
    Frame1.Top = Frame1.Top + lngDiff
    Frame2.Top = Frame2.Top + lngDiff
    Frame3.Top = Frame3.Top + lngDiff
    lblpink.Top = lblpink.Top + lngDiff
    lblGreen.Top = lblGreen.Top + lngDiff
    lblRed.Top = lblRed.Top + lngDiff
    lblGrey.Top = lblGrey.Top + lngDiff
    Frame5.Top = Frame5.Top + lngDiff
    ProgressBar1.Top = ProgressBar1.Top + lngDiff
    cmdUncheck.Top = cmdUncheck.Top + lngDiff
    cmdCheckAll.Top = cmdCheckAll.Top + lngDiff
    cmdExcelExport.Top = cmdExcelExport.Top + lngDiff
    cmdPrint.Top = cmdPrint.Top + lngDiff
    cmdQuit.Top = cmdQuit.Top + lngDiff
    cmdSave.Top = cmdSave.Top + lngDiff
    frmGenerate.Top = frmGenerate.Top + lngDiff
    txtResult.Height = txtResult.Height + lngDiff

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub optC_Click()
    On Error GoTo errHandler
    SaveSetting "PBKS", "COLALLOC", "DiscountFrom", IIf(optC = True, "C", "CO")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.optC_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optCO_Click()
    On Error GoTo errHandler
    SaveSetting "PBKS", "COLALLOC", "DiscountFrom", IIf(optC = True, "C", "CO")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.optCO_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo errHandler
Dim strError As String
'    If strType <> "DELIVERY" And bOKtoClose = False Then
'        If MsgBox("You want to close this form without generating invoices? ", vbQuestion + vbYesNo, "Confirm") = vbNo Then
'            Cancel = True
'            Exit Sub
'        End If
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.Form_QueryUnload(Cancel,UnloadMode)", Array(Cancel, UnloadMode), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    rs.Close
    Set rs = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_DblClick()
    On Error GoTo errHandler
Dim frm As frmProductPrev
Dim frmNB As frmProductNBPrev   'non book form
Dim lngprod As Long
Dim str As String
Dim strErrPos As String

    If XA.UpperBound(1) = 0 Then Exit Sub
    If IsNull(Grid1.Bookmark) Then Exit Sub
    If Err Then Exit Sub

    str = FNS(XA.Value(Grid1.Bookmark, 17))
    If str = "" Then Exit Sub
    Set roProduct = New a_Product
    WaitMsg "Loading . . .", True, Me
    roProduct.Load str, 0, ""
    If roProduct.PID = "" Then Exit Sub
    If roProduct.ProductType = "B" Then
            Set frm = New frmProductPrev
            frm.component roProduct
            frm.Show
    Else
        Set frmNB = New frmProductNBPrev
        frmNB.component roProduct
        frmNB.Show
    End If
    Set roProduct = Nothing
    WaitMsg "", False, Me

    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmCOLAllocation: Grid1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmCOLAllocation: Grid1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.Grid1_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
Dim strStatus As String
    If Left(XA(Bookmark, 8), 3) = "**" Then
        RowStyle.BackColor = COLOR_CANCELLED
        Exit Sub
    End If
    If InStr(1, XA(Bookmark, 8), "*") > 0 Then
        RowStyle.ForeColor = vbRed
    End If
    strStatus = XA(Bookmark, 25)
    If strStatus = "INVALID" Then
        RowStyle.BackColor = COLOUR_FULFIL_MORETHANONHAND
    ElseIf strStatus = "OK" Then
'        RowStyle.BackColor = COLOUR_FULFIL_OK
        Grid1.Columns(1).BackColor = COLOUR_FULFIL_OK
        Grid1.Columns(3).BackColor = COLOUR_FULFIL_OK
        Grid1.Columns(6).BackColor = COLOUR_FULFIL_OK
        Grid1.Columns(7).BackColor = COLOUR_FULFIL_OK
        Grid1.Columns(9).BackColor = COLOUR_FULFIL_OK
        Grid1.Columns(8).BackColor = vbWhite
    ElseIf strStatus = "MORE" Then
        RowStyle.BackColor = COLOUR_FULFIL_LESSTHANORDERED
    ElseIf strStatus = "TOOMUCH" Then
        RowStyle.BackColor = COLOUR_FULFIL_MORETHANORDERED
    End If
    If XA(Bookmark, 18) <> 0 Then
        RowStyle.Font.Italic = True
        RowStyle.Font.Underline = True
    End If
    If XA(Bookmark, 24) = True Then  'CUstomer is blocked
        RowStyle.BackColor = COLOR_CANCELLED
    End If
    Grid1.Columns(8).BackColor = vbWhite
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
Dim i As Integer
Dim oC As a_Copy
Dim lngResult As Long
Dim iOK As Integer
Dim oALLOC As a_COLAllocation
Dim lngFirm As Long
Dim lngSS As Long
Dim bOK As Boolean
Dim tmp As Long

    If Left(Grid1.Columns(7).Value, 3) = "**" Then
        Cancel = True
        Exit Sub
    End If
    i = ColIndex + 1
    Select Case i
    Case 9
        If oPC.AllowsSSInvoicing Then
            If Grid1.Text > "" Then
                bOK = SSFIRM_OK(Grid1.Text)

                If bOK = True Then
                    Set oALLOC = cALLOC(XA(Grid1.Bookmark, 15))
                    oALLOC.BeginEdit
                    TotalSSFIRM Grid1.Text, lngFirm, lngSS
                    oALLOC.SetAllocatedQty lngFirm
                    oALLOC.SetAllocatedQtySS lngSS
                    Grid1.Columns(ColIndex).Value = CStr(lngFirm) & "/" & CStr(lngSS)
                    oALLOC.ApplyEdit
                    cALLOC.GetStatus
                    
                    rs.Find "PID = " & oALLOC.PID, 0, adSearchForward, 1
                    If Not rs.eof Then
                        If InStr(1, OldValue, "/") > 0 Then
                            If SSFIRM_OK(CStr(OldValue)) Then
                                tmp = TotalSSFIRM(CStr(OldValue), lngFirm, lngSS)
                            Else
                                tmp = 0
                            End If
                        Else
                            tmp = FNN(OldValue)
                        End If
                            rs.Fields("BAL") = FNN(rs.Fields("BAL")) - (FNN(oALLOC.AllocatedQty) + FNN(oALLOC.AllocatedQtySS)) + FNN(tmp)
                        rs.Update
                    End If
                    If (FNN(oALLOC.AllocatedQty) + FNN(oALLOC.AllocatedQtySS)) > TotalSSFIRM(XA(Grid1.Bookmark, 6), lngFirm, lngSS) - val(XA(Grid1.Bookmark, 7)) Then
                        MarkRowsValid 4, cALLOC(val(XA(Grid1.Bookmark, 15))).key    'toomuch
                    ElseIf (FNN(oALLOC.AllocatedQty) + FNN(oALLOC.AllocatedQtySS)) > val(XA(Grid1.Bookmark, 14)) Then
                        MarkRowsValid 3, cALLOC(val(XA(Grid1.Bookmark, 15))).key    'invalid
                    ElseIf (FNN(oALLOC.AllocatedQty) + FNN(oALLOC.AllocatedQtySS)) < TotalSSFIRM(XA(Grid1.Bookmark, 6), lngFirm, lngSS) - val(XA(Grid1.Bookmark, 7)) Then
                        MarkRowsValid 2, cALLOC(val(XA(Grid1.Bookmark, 15))).key    'more
                    ElseIf (FNN(oALLOC.AllocatedQty) + FNN(oALLOC.AllocatedQtySS)) = TotalSSFIRM(XA(Grid1.Bookmark, 6), lngFirm, lngSS) - val(XA(Grid1.Bookmark, 7)) Then
                        MarkRowsValid 1, cALLOC(val(XA(Grid1.Bookmark, 15))).key    'OK
                    End If
                    oALLOC.SetACtionYN (FNN(oALLOC.AllocatedQty) + FNN(oALLOC.AllocatedQtySS)) > 0
                    XA.Value(Grid1.Bookmark, 10) = IIf((FNN(oALLOC.AllocatedQty) + FNN(oALLOC.AllocatedQtySS)) > 0, -1, 0)

                    Grid1.Update
                    cmdSave.Enabled = True
                    rs.MoveFirst
                    For i = 1 To rs.RecordCount - 1
                        If rs.Fields("BAL") < 0 Then
                            Me.cmdSave.Enabled = False
                            Exit For
                        End If
                        rs.MoveNext
                    Next
                Else
                    Cancel = True
                End If
            Else
                Cancel = True
            End If
        Else
            If ConvertToLng(Grid1.Text, lngResult) Then
                Grid1.Text = CStr(lngResult)
                Set oALLOC = cALLOC(XA(Grid1.Bookmark, 15))
                oALLOC.BeginEdit
                oALLOC.SetAllocatedQty lngResult
                oALLOC.ApplyEdit
                cALLOC.GetStatus
                
                rs.Find "PID = " & oALLOC.PID, 0, adSearchForward, 1
                If Not rs.eof Then
                        rs.Fields("BAL") = FNN(rs.Fields("BAL")) - FNN(oALLOC.AllocatedQty) + FNN(OldValue)
                    rs.Update
                End If
                If lngResult > val(XA(Grid1.Bookmark, 6)) - val(XA(Grid1.Bookmark, 7)) Then
                    MarkRowsValid 4, cALLOC(val(XA(Grid1.Bookmark, 15))).key
                ElseIf lngResult > val(XA(Grid1.Bookmark, 14)) Then
                    MarkRowsValid 3, cALLOC(val(XA(Grid1.Bookmark, 15))).key
                ElseIf lngResult < val(XA(Grid1.Bookmark, 6)) - val(XA(Grid1.Bookmark, 7)) Then
                    MarkRowsValid 2, cALLOC(val(XA(Grid1.Bookmark, 15))).key
                ElseIf lngResult = val(XA(Grid1.Bookmark, 6)) - val(XA(Grid1.Bookmark, 7)) Then
                    MarkRowsValid 1, cALLOC(val(XA(Grid1.Bookmark, 15))).key
                End If
                Grid1.Update
                cmdSave.Enabled = True
                For i = 1 To rs.RecordCount
                    If rs.Fields("BAL") < 0 Then
                        Me.cmdSave.Enabled = False
                        Exit For
                    End If
                Next
            Else
                Cancel = True
            End If
        End If
    Case 10
       ' cALLOC.GetStatus
        
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
         OldValue, Cancel), EA_NORERAISE
    HandleError
End Sub
Private Function TotalSSFIRM(p As String, Firm As Long, SS As Long) As Long
    On Error GoTo errHandler
Dim tmp As Long

    If InStr(1, p, "/") > 0 Then
        If ConvertToLng(Left(p, InStr(1, p, "/") - 1), Firm) And ConvertToLng(Mid(p, InStr(1, p, "/") + 1, 99), SS) Then
            TotalSSFIRM = SS + Firm
        Else
            TotalSSFIRM = 0
        End If
    Else
        ConvertToLng p, Firm
        TotalSSFIRM = Firm
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.TotalSSFIRM(p,Firm,SS)", Array(p, Firm, SS)
End Function
Private Function SSFIRM_OK(p As String) As Boolean
    On Error GoTo errHandler
Dim SS As Long
Dim Firm As Long
    If InStr(1, p, "/") > 0 Then
        SSFIRM_OK = ConvertToLng(Left(p, InStr(1, p, "/") - 1), Firm) And ConvertToLng(Mid(p, InStr(1, p, "/") + 1, 99), SS)
    Else
        SSFIRM_OK = ConvertToLng(p, Firm)
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.SSFIRM_OK(p)", p, EA_NORERAISE
    HandleError
End Function
Private Sub Grid1_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Dim oALLOC As a_COLAllocation
Dim i As Integer
    i = ColIndex + 1
    If i = 10 Then   'checkbox
        Set oALLOC = cALLOC(XA(Grid1.Bookmark, 15))
        oALLOC.BeginEdit
        oALLOC.SetACtionYN (Grid1.Text = -1)
        oALLOC.ApplyEdit
    End If
    Grid1.Columns(8).BackColor = vbWhite
        cALLOC.GetStatus

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.Grid1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant
    Screen.MousePointer = vbHourglass

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) '', 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    Grid1.Refresh
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.Grid1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 3, 4, 8
            GetRowType = XTYPE_STRING
        Case 6, 7, 9
            GetRowType = XTYPE_NUMBER
        Case 5
            GetRowType = XTYPE_DATE
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.GetRowType(ColIndex)", ColIndex
End Function
Private Sub EnableOK(pOK As Boolean)
    On Error GoTo errHandler
    Me.cmdSave.Enabled = pOK
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.EnableOK(pOK)", pOK
End Sub
Private Sub MarkRowsValid(pOK As Integer, pKey As String)
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To lngArrayRows
        If XA(i, 15) = pKey Then
            Select Case pOK
                Case 1
                    XA(i, 25) = "OK"
                Case 2
                    XA(i, 25) = "MORE"
                Case 3
                    XA(i, 25) = "INVALID"
                Case 4
                    XA(i, 25) = "TOOMUCH"
            End Select
        End If
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.MarkRowsValid(pOK,pKey)", Array(pOK, pKey)
End Sub
Private Sub Grid1_LostFocus()
    On Error GoTo errHandler
    Grid1.Update
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation.Grid1_LostFocus", , EA_NORERAISE
    HandleError
End Sub


