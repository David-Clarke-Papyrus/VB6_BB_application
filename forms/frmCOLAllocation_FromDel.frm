VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmCOLAllocation_FromDel 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Reserve items for customers"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14790
   ControlBox      =   0   'False
   FillColor       =   &H00FFC0FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   14790
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
      Height          =   630
      Left            =   7875
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Sets all the allocations to those shown when the form was opened"
      Top             =   5085
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   60
      TabIndex        =   11
      Top             =   5745
      Width           =   405
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   60
      TabIndex        =   9
      Top             =   5430
      Width           =   405
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   60
      TabIndex        =   7
      Top             =   5100
      Width           =   405
   End
   Begin VB.TextBox txtResult 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8E8E8&
      ForeColor       =   &H000000C0&
      Height          =   4500
      Left            =   12390
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      ToolTipText     =   "These products have too many allocations, i.e. there is not sufficient stock to meet the allocations. Reduce the allocations."
      Top             =   420
      Width           =   2295
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   270
      Left            =   3765
      TabIndex        =   4
      Top             =   5715
      Visible         =   0   'False
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdGenInv 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Create invoices"
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
      Height          =   630
      Left            =   10935
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Creates invoices for all products/customer orderlines where qty > 0"
      Top             =   5085
      Visible         =   0   'False
      Width           =   1200
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
      Height          =   630
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Sets all the allocations to those shown when the form was opened"
      Top             =   5085
      Width           =   990
   End
   Begin VB.CommandButton cmdSave 
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
      Left            =   9885
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print the allocation slip"
      Top             =   5085
      Width           =   990
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   4830
      Left            =   180
      OleObjectBlob   =   "frmCOLAllocation_FromDel.frx":0000
      TabIndex        =   0
      Top             =   90
      Width           =   12135
   End
   Begin VB.Label lblItemsCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   3045
      TabIndex        =   14
      Top             =   5025
      Width           =   4410
   End
   Begin VB.Label Label4 
      BackColor       =   &H00D3D3CB&
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
      TabIndex        =   12
      Top             =   5700
      Width           =   2580
   End
   Begin VB.Label Label3 
      BackColor       =   &H00D3D3CB&
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
      TabIndex        =   10
      Top             =   5385
      Width           =   2580
   End
   Begin VB.Label Label2 
      BackColor       =   &H00D3D3CB&
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
      Left            =   525
      TabIndex        =   8
      Top             =   5055
      Width           =   2580
   End
   Begin VB.Label Label1 
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
      Left            =   12270
      TabIndex        =   5
      Top             =   135
      Width           =   2265
   End
End
Attribute VB_Name = "frmCOLAllocation_FromDel"
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


Private Sub PrepareNewSlate()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    Set cALLOC = Nothing
    Set cALLOC = New chex_COLAllocation
    cALLOC.GenerateCOLAllocationset lngDELID 'INSERT INTO tCOLAlloc
    cALLOC.Load lngDELID
    LoadGrid
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.PrepareNewSlate"
End Sub

Private Sub cALLOC_Valid(pResult As String)
    On Error GoTo errHandler
'MsgBox pOK
    EnableOK (pResult = "")
    txtResult = pResult
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.cALLOC_Valid(pResult)", pResult, EA_NORERAISE
    HandleError
End Sub
Public Sub component(pcALLOC As chex_COLAllocation, pType As String, pViewOnly As Boolean)
    On Error GoTo errHandler
    strType = pType
    If strType = "DELIVERY" Or strType = "TRANSFER" Then
        If pViewOnly Then
            Me.Caption = "Allocate delivered items to outstanding customer orders"
            Grid1.FetchRowStyle = False
            Me.cmdGenInv.Visible = False
            Me.cmdSave.Visible = False
            Me.cmdQuit.Visible = True
            Me.Grid1.Columns.Item(5).Visible = False
            Me.cmdPrint.Visible = True
        Else
            Me.cmdGenInv.Visible = False
            Me.cmdSave.Visible = True
            Me.cmdQuit.Visible = False
            Me.cmdPrint.Visible = False
        End If
    Else
        Me.cmdGenInv.Visible = True
        Me.cmdSave.Visible = False
        Me.cmdQuit.Visible = True
        Me.cmdPrint.Visible = False
        Me.Caption = "Fulfil customer orders from stock"
    End If
    Set cALLOC = pcALLOC
    LoadGrid
    lblItemsCount.Caption = cALLOC.Count & " items"
    cALLOC.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.component(pcALLOC,pType,pViewOnly)", Array(pcALLOC, pType, _
         pViewOnly)
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

    i = 0
    Set XA = New XArrayDB
    XA.Clear
    iRecs = i
    lngIndex = 1
    lngArrayRows = cALLOC.Count
    XA.ReDim 1, lngArrayRows, 1, 14
    Set rs = New ADODB.Recordset
    rs.Fields.Append "PID", adVarChar, 40, adFldKeyColumn
    rs.Fields.Append "Bal", adInteger
    rs.Open
    For Each oALLOC In cALLOC
        lngAvailableToAllocate = oALLOC.QtyOnHand
        lngAvailableToAllocate = 1
        XA.Value(lngIndex, 13) = oALLOC.PID   ' was ""
        rs.Find ("PID = " & oALLOC.PID)
        
        If rs.eof Then
            rs.AddNew
                rs.Fields("PID") = oALLOC.PID
                rs.Fields("BAL") = oALLOC.QtyOnHand
            rs.Update
        End If
        
        rs.Fields("BAL") = rs.Fields("BAL") - oALLOC.AllocatedQty
        rs.Update
        
        XA.Value(lngIndex, 1) = oALLOC.CodeTitleShort(40) & "  (" & oALLOC.ProductOH & " / " & oALLOC.ProductRES & ")"
        XA.Value(lngIndex, 2) = IIf(oALLOC.DepositValue > 0, "(Dep." & oALLOC.DepositValueF & ") ", "") & oALLOC.CustomerName
        XA.Value(lngIndex, 3) = oALLOC.OrderedQty
        XA.Value(lngIndex, 4) = oALLOC.DeliveredSoFar
        XA.Value(lngIndex, 5) = oALLOC.OrderDetails
        
        oALLOC.AllocatedQty = GetMin(oALLOC.OrderedQty - oALLOC.DeliveredSoFar, rs.Fields("BAL"), False)
            rs.Fields("BAL") = rs.Fields("BAL") - oALLOC.AllocatedQty
            rs.Update
        XA.Value(lngIndex, 7) = NonNegative_Lng(oALLOC.AllocatedQty)
        
        oALLOC.BeginEdit
        oALLOC.SetACtionYN True
        oALLOC.ApplyEdit
        XA.Value(lngIndex, 6) = oALLOC.Note
        XA.Value(lngIndex, 8) = -1
        
        XA.Value(lngIndex, 9) = oALLOC.COLID
        XA.Value(lngIndex, 10) = ""
        XA.Value(lngIndex, 11) = oALLOC.QtyOnHand
        XA.Value(lngIndex, 12) = oALLOC.Key
        XA.Value(lngIndex, 14) = ""
                    
        If oALLOC.AllocatedQty > val(oALLOC.OrderedQty) - val(oALLOC.DeliveredSoFar) Then
            MarkRowsValid 4, cALLOC(val(oALLOC.Key)).Key
        ElseIf oALLOC.AllocatedQty > val(oALLOC.QtyOnHand) Then
            MarkRowsValid 3, cALLOC(val(oALLOC.Key)).Key
        ElseIf oALLOC.AllocatedQty < val(oALLOC.OrderedQty) - val(oALLOC.DeliveredSoFar) Then
            MarkRowsValid 2, cALLOC(val(oALLOC.Key)).Key
        ElseIf oALLOC.AllocatedQty = val(oALLOC.OrderedQty) - val(oALLOC.DeliveredSoFar) Then
            MarkRowsValid 1, cALLOC(val(oALLOC.Key)).Key
        End If
        
        lngIndex = lngIndex + 1
    Next
    Set Grid1.Array = XA
    Grid1.ReBind
   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.LoadGrid"
End Sub

Private Sub cmdReset_Click()
    On Error GoTo errHandler
    If cALLOC.IsEditing Then cALLOC.CancelEdit
    Set cALLOC = Nothing
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.cmdReset_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdGenInv_Click()
    On Error GoTo errHandler
Dim oG As New z_OrderFulfilmentDocGen
Dim strError As String
    WaitMsg "Creating invoices  . . .", True, Me
    
    cALLOC.Save strError, "NORMAL", "I"
    oG.GenerateInvoicesFromCOLALLOCs "CO"
    
    WaitMsg "", False, Me
    bOKtoClose = True
    MsgBox "Invoices have been generated. They are not yet issued." & vbCrLf & "Please check and issue each invoice that has been generated.", vbInformation, "Please note"
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.cmdGenInv_Click", , EA_NORERAISE
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
    ErrorIn "frmCOLAllocation_FromDel.cmdLoad_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    
Dim oC As chex_COLAllocation
Dim ar As New arCOLSFulfilled
'    Set oC = New chex_COLAllocation
'    oC.Load oDel.TRID, True
    ar.component cALLOC
    ar.Show vbModal
    
    
    
    
   ' Grid1.PrintInfo.SettingsDeviceName ='  oPC.Configuration.DocumentControls.FindDCByTypeName("CUSTOMER ORDER").GetPrinter(oPC.NameOfPC)
'    Grid1.PrintInfo.SettingsPaperSize = 9  'vbPRPSA4
'    Grid1.PrintInfo.PageHeader = "\t" & "Hold these products in reserve"
'    Grid1.PrintInfo.PageFooter = "\tPage:  \p of page \P"
'    Grid1.PrintInfo.SettingsOrientation = 2
'        Grid1.PrintInfo.PageSetup
'        Grid1.PrintInfo.PrintPreview 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdQuit_Click()
10        On Error GoTo errHandler
20        Unload Me
30        Exit Sub
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "frmCOLAllocation_FromDel.cmdQuit_Click", , EA_NORERAISE
60        HandleError
End Sub

Private Sub cmdSave_Click()
10        On Error GoTo errHandler
      Dim strError As String
      Dim arReserve() As Reserve
      Dim oCA As a_COLAllocation
      Dim cALLOCOCS As c_COLsperDEL
      Dim rpt As arCOLSFulfilled
      Dim x As a_COLAllocation
      Dim i As Integer
      Dim bMustPrint As Boolean
      Dim oSQL As z_SQL
      Dim rs As ADODB.Recordset


20        If strType <> "DELIVERY" And strType <> "TRANSFER" Then
30            If MsgBox("You want to save the allocations without issuing invoices yet. These will be available until you next prepare an order allocation.", vbInformation + vbYesNo, "Confirm") = vbNo Then
40                Exit Sub
50            End If
60        End If
70        Screen.MousePointer = vbHourglass
        LogSaveToFile "Pos 1 (click Enter to continue)"
80        Grid1.Update
90        If Not cALLOC Is Nothing Then
100           cALLOC.Save strError, strType, "I"
110           lngDELID = cALLOC.DELID
120           Set cALLOC = Nothing
130           Set cALLOC = New chex_COLAllocation
140           cALLOC.Load lngDELID, True
150       End If
160       bMustPrint = False
170       For i = 1 To cALLOC.Count
180           If cALLOC(i).AllocatedQty > 0 Or cALLOC(i).AllocatedQtySS > 0 Then
190               bMustPrint = True
200           End If
210       Next
        LogSaveToFile "Pos 4 (click Enter to continue)"
220       If bMustPrint Then
230           Set oSQL = New z_SQL
240           Set rs = New ADODB.Recordset
              
         LogSaveToFile "Pos 5 (click Enter to continue)"
         
250           oSQL.RunGetRecordset "SELECT * FROM " & IIf(strType = "TRANSFER", "vTransferInAllocations", "vDeliveryAllocations") & "  WHERE TR_ID = " & CStr(lngDELID), enText, Array(), "", rs
         LogSaveToFile "Pos 6 (click Enter to continue)"
260           Set rpt = New arCOLSFulfilled
270           rpt.component rs
         LogSaveToFile "Pos 6 (click Enter to continue)"
280           rpt.Printer.Orientation = ddOPortrait
290            Screen.MousePointer = vbDefault
300           rpt.Show vbModal
         LogSaveToFile "Pos 7 (click Enter to continue)"
            '  rpt.PrintReport False
310           If Err > 0 Then
320               MsgBox "Problem printing customer delivery allocation", vbInformation + vbOKOnly, "Status"
330           End If
340       End If

         LogSaveToFile "Pos 8 (click Enter to continue)"

350       Unload Me
360       Exit Sub
errHandler:
370       If ErrMustStop Then Debug.Assert False: Resume
380       ErrorIn "frmCOLAllocation_FromDel.cmdSave_Click", , EA_NORERAISE
390       HandleError
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
    strType = "NORMAL"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        TOP = 100
        Left = 100
        Height = 6400
        Width = 14910
    End If
    bOKtoClose = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo errHandler
Dim strError As String
    If strType <> "DELIVERY" And strType <> "TRANSFER" And bOKtoClose = False Then
        If MsgBox("You want to close this form without generating invoices? ", vbQuestion + vbYesNo, "Confirm") = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.Form_QueryUnload(Cancel,UnloadMode)", Array(Cancel, UnloadMode), _
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
    ErrorIn "frmCOLAllocation_FromDel.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_DblClick()
Dim frm As frmProductPrev
Dim frmNB As frmProductNBPrev   'non book form
Dim lngprod As Long
Dim str As String
Dim strErrPos As String

On Error Resume Next
    If XA.UpperBound(1) = 0 Then Exit Sub
    If IsNull(Grid1.Bookmark) Then Exit Sub
    If Err Then Exit Sub

On Error GoTo errHandler
    str = FNS(XA.Value(Grid1.Bookmark, 13))
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
        LogSaveToFile "Access violation in frmCOLAllocation_FromDel: Grid1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmCOLAllocation_FromDel: Grid1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.Grid1_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
Dim strStatus As String
    strStatus = XA(Bookmark, 14)
    If strStatus = "MORETHANONHAND" Then
        RowStyle.BackColor = COLOUR_FULFIL_MORETHANONHAND
    ElseIf strStatus = "OK" Then
        RowStyle.BackColor = COLOUR_FULFIL_OK
    ElseIf strStatus = "LESSTHANORDERED" Then
        RowStyle.BackColor = COLOUR_FULFIL_LESSTHANORDERED
    ElseIf strStatus = "MORETHANORDERED" Then
        RowStyle.BackColor = COLOUR_FULFIL_MORETHANORDERED
    End If
    Grid1.Columns(5).BackColor = vbWhite
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, _
         Bookmark, RowStyle), EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
Dim i As Integer
Dim oC As a_Copy
Dim lngResult As Long
Dim iOK As Integer
Dim oALLOC As a_COLAllocation

    i = ColIndex + 1
    Select Case i
    Case 7
        If ConvertToLng(Grid1.text, lngResult) Then
            Grid1.text = CStr(lngResult)
            Set oALLOC = cALLOC(XA(Grid1.Bookmark, 12))
            oALLOC.BeginEdit
            oALLOC.SetAllocatedQty lngResult
            oALLOC.ApplyEdit
            cALLOC.GetStatus
            
            rs.Find "PID = " & oALLOC.PID, 0, adSearchForward, 1
            If Not rs.eof Then
                    rs.Fields("BAL") = FNN(rs.Fields("BAL")) - FNN(oALLOC.AllocatedQty) + FNN(OldValue)
                rs.Update
            End If
            If lngResult > val(XA(Grid1.Bookmark, 3)) - val(XA(Grid1.Bookmark, 4)) Then
                MarkRowsValid 4, cALLOC(val(XA(Grid1.Bookmark, 12))).Key
            ElseIf lngResult > val(XA(Grid1.Bookmark, 10)) Then
                MarkRowsValid 3, cALLOC(val(XA(Grid1.Bookmark, 12))).Key
            ElseIf lngResult < val(XA(Grid1.Bookmark, 3)) - val(XA(Grid1.Bookmark, 4)) Then
                MarkRowsValid 2, cALLOC(val(XA(Grid1.Bookmark, 12))).Key
            ElseIf lngResult = val(XA(Grid1.Bookmark, 3)) - val(XA(Grid1.Bookmark, 4)) Then
                MarkRowsValid 1, cALLOC(val(XA(Grid1.Bookmark, 12))).Key
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
    Case 7
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", _
         Array(ColIndex, OldValue, Cancel), EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Dim oALLOC As a_COLAllocation
Dim i As Integer
    i = ColIndex + 1
    If i = 8 Then   'checkbox
        Set oALLOC = cALLOC(XA(Grid1.Bookmark, 12))
        oALLOC.BeginEdit
        oALLOC.SetACtionYN (Grid1.text = -1)
        oALLOC.ApplyEdit
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.Grid1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_AfterColEdit(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Dim oALLOC As a_COLAllocation
Dim i As Integer
    i = ColIndex + 1
    If i = 7 Then   'checkbox
        Set oALLOC = cALLOC(XA(Grid1.Bookmark, 12))
        oALLOC.BeginEdit
        oALLOC.SetACtionYN (XA(Grid1.Bookmark, 8) = -1)
        oALLOC.ApplyEdit
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.Grid1_AfterColEdit(ColIndex)", ColIndex, EA_NORERAISE
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
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    Grid1.Refresh
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.Grid1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 5, 6
            GetRowType = XTYPE_STRING
        Case 3, 4, 7
            GetRowType = XTYPE_NUMBER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.GetRowType(ColIndex)", ColIndex
End Function
Private Sub EnableOK(pOK As Boolean)
    On Error GoTo errHandler
    Me.cmdSave.Enabled = pOK
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.EnableOK(pOK)", pOK
End Sub
Private Sub MarkRowsValid(pOK As Integer, pKey As String)
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To lngArrayRows
        If XA(i, 12) = pKey Then
            Select Case pOK
                Case 1
                    XA(i, 14) = "OK"
                Case 2
                    XA(i, 14) = "LESSTHANORDERED"
                Case 3
                    XA(i, 14) = "MORETHANONHAND"
                Case 4
                    XA(i, 14) = "MORETHANORDERED"
            End Select
        End If
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.MarkRowsValid(pOK,pKey)", Array(pOK, pKey)
End Sub
Private Sub Grid1_LostFocus()
    On Error GoTo errHandler
    Grid1.Update
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOLAllocation_FromDel.Grid1_LostFocus", , EA_NORERAISE
    HandleError
End Sub






