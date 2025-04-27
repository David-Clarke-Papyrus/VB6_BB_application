VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmCOLDELLAllocation 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Customer order line reconciliation"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12270
   ControlBox      =   0   'False
   FillColor       =   &H00FFC0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   12270
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   270
      Left            =   3930
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdGenInv 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Generate invoices"
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
      Height          =   600
      Left            =   10140
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5100
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrepareNewSlate 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Prepare new allocation slate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5085
      Width           =   1695
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Use existing allocation slate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5085
      Width           =   1695
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5055
      Width           =   1200
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
      Height          =   630
      Left            =   8940
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5070
      Width           =   1200
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   4830
      Left            =   180
      OleObjectBlob   =   "frmCOLDELLAllocation.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   11970
   End
   Begin VB.Label lblResult 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   750
      Left            =   3570
      TabIndex        =   3
      Top             =   5130
      Width           =   4050
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00706034&
      Height          =   4875
      Left            =   135
      Top             =   120
      Width           =   11130
   End
End
Attribute VB_Name = "frmCOLDELLAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oRS As ADODB.Recordset
Dim WithEvents cCOLDELL As chex_COLDELLAllocation
Attribute cCOLDELL.VB_VarHelpID = -1
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim iRecs As Integer
Dim lngArrayRows As Long
Dim lngDELID As Long
Dim rs As New ADODB.Recordset
Dim strType As String

Private Sub cCOLDELL_Valid(pResult As String)
'MsgBox pOK
    EnableOK (pResult = "")
    lblResult.Caption = pResult
End Sub
Public Sub Component(pcCOLDELL As chex_COLDELLAllocation, pType As String)
    strType = pType
    If pType = "DELIVERY" Then
        Me.cmdGenInv.Visible = False
        Me.cmdLoad.Visible = False
        Me.cmdPrepareNewSlate.Visible = False
        Me.cmdReset.Visible = False
    End If
    Set cCOLDELL = pcCOLDELL
    Me.cmdPrepareNewSlate.Enabled = False
    Me.cmdLoad.Enabled = False
    LoadGrid
End Sub

Private Sub LoadGrid()
Dim objItm As ListItem
Dim lngIndex As Long
Dim tmp As String
Dim qtyRecs As Long
Dim lngAwaiting As Long
Dim lngAllocation As Long
Dim lngAvailableToAllocate As Long
Dim i As Integer
Dim oCOLDELL As a_COLDELLAllocation

    i = 0
    Set XA = New XArrayDB
    XA.Clear
    iRecs = i
    lngIndex = 1
    lngArrayRows = cCOLDELL.Count
    XA.ReDim 1, lngArrayRows, 1, 12
    Set rs = New ADODB.Recordset
    rs.Fields.Append "PID", adVarChar, 40, adFldKeyColumn
    rs.Fields.Append "Bal", adInteger
    rs.Open
    For Each oCOLDELL In cCOLDELL
        lngAvailableToAllocate = oCOLDELL.QtyOnHand
            XA.Value(lngIndex, 1) = oCOLDELL.Titleshort(15) & "  (" & oCOLDELL.QtyOnHand & " / " & oCOLDELL.QtyReserved & ")"
        '    XA.Value(lngIndex, 2) = oCOLDELL.QtyReserved
            XA.Value(lngIndex, 2) = oCOLDELL.CustomerName
            XA.Value(lngIndex, 3) = oCOLDELL.OrderedQty
            XA.Value(lngIndex, 4) = oCOLDELL.DeliveredSoFar
            XA.Value(lngIndex, 5) = oCOLDELL.OrderDetails
            XA.Value(lngIndex, 6) = oCOLDELL.AllocatedQty
            XA.Value(lngIndex, 7) = oCOLDELL.COLID
            XA.Value(lngIndex, 8) = ""
            XA.Value(lngIndex, 9) = ""
            XA.Value(lngIndex, 10) = oCOLDELL.Key
            XA.Value(lngIndex, 11) = ""
            XA.Value(lngIndex, 12) = ""
            rs.Find ("PID = " & oCOLDELL.pID)
            If rs.EOF Then
                rs.AddNew
                    rs.Fields("PID") = oCOLDELL.pID
                    rs.Fields("BAL") = oCOLDELL.QtyOnHand - oCOLDELL.AllocatedQty
                rs.Update
            Else
                rs.Fields("BAL") = rs.Fields("BAL") - oCOLDELL.AllocatedQty
                rs.Update
            End If
            If rs.Fields("BAL") < 0 Then
                XA.Value(lngIndex, 12) = "INVALID"
            End If
            lngIndex = lngIndex + 1
    Next
    XA.QuickSort 1, lngArrayRows, 1, XORDER_ASCEND, XTYPE_STRING, 2, XORDER_ASCEND, XTYPE_INTEGER, 3, XORDER_ASCEND, XTYPE_INTEGER, 4, XORDER_ASCEND, XTYPE_INTEGER
    Set Grid1.Array = XA
    Grid1.ReBind
   
  '  Grid1.SetFocus
End Sub

Private Sub cmdReset_Click()
    If cCOLDELL.IsEditing Then cCOLDELL.CancelEdit
    Set cCOLDELL = Nothing
    Unload Me
End Sub

Private Sub cmdGenInv_Click()
Dim oG As New z_InvoiceGenerator
Dim strError As String
    cCOLDELL.Save strError, "NORMAL"
    oG.GenerateInvoicesFromCOLDELLs
    Unload Me
End Sub

Private Sub cmdLoad_Click()
Screen.MousePointer = vbHourglass
    Set cCOLDELL = New chex_COLDELLAllocation
    cCOLDELL.Load 'oDEL.TRID
    LoadGrid
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdClose_Click()
Dim strError As String
    If Not cCOLDELL Is Nothing Then
        cCOLDELL.Save strError, strType
    End If
    If strType = "DELIVERY" Then
        Grid1.PrintInfo.PageHeader = "\t" & "Hold these products in reserve"
        Grid1.PrintInfo.PageFooter = "\tPage:  \p of page \P"
        Grid1.PrintInfo.SettingsOrientation = 2
        Grid1.PrintInfo.PrintData
    End If
    Unload Me
End Sub

Private Sub cmdPrepareNewSlate_Click()
'Dim oG As New z_InvoiceGenerator
    Screen.MousePointer = vbHourglass
    Set cCOLDELL = Nothing
    Set cCOLDELL = New chex_COLDELLAllocation
    cCOLDELL.GenerateCOLDELLAllocationset lngDELID
    cCOLDELL.Load lngDELID
    LoadGrid
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Initialize()
    strType = "NORMAL"
End Sub

Private Sub Form_Unload(cancel As Integer)
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Grid1_AfterUpdate()
'    Grid1.ReBind
End Sub

Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
Dim strStatus As String
    strStatus = XA(Bookmark, 12)
    If strStatus = "INVALID" Then
        RowStyle.BackColor = &HFFC0FF
    ElseIf strStatus = "OK" Then
        RowStyle.BackColor = &HDBFAFB
    ElseIf strStatus = "MORE" Then
        RowStyle.BackColor = &HC0FFC0
    ElseIf strStatus = "TOOMUCH" Then
        RowStyle.BackColor = &H8080FF
    End If
End Sub
Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, cancel As Integer)
Dim i As Integer
Dim oC As a_Copy
Dim lngResult As Long
Dim iOK As Integer
Dim oCDALLOC As a_COLDELLAllocation
On Error GoTo Errh

    i = ColIndex + 1
'    On Error Resume Next
    Select Case i
    Case 6
        If ConvertToLng(Grid1.Text, lngResult) Then
            Grid1.Text = CStr(lngResult)
            Set oCDALLOC = cCOLDELL(val(XA(Grid1.Bookmark, 10)))
            oCDALLOC.BeginEdit
            oCDALLOC.SetAllocatedQty lngResult
            oCDALLOC.ApplyEdit
            
            rs.Find "PID = " & oCDALLOC.pID, 0, adSearchForward, 1
            If Not rs.EOF Then
                    rs.Fields("BAL") = rs.Fields("BAL") - FNN(oCDALLOC.AllocatedQty) + FNN(OldValue)
                rs.Update
            End If
            If rs.Fields("BAL") < 0 Then
                MarkRowsValid 3, cCOLDELL(val(XA(Grid1.Bookmark, 10))).Key
            ElseIf lngResult > val(XA(Grid1.Bookmark, 3)) - val(XA(Grid1.Bookmark, 4)) Then
                MarkRowsValid 4, cCOLDELL(val(XA(Grid1.Bookmark, 10))).Key
            ElseIf lngResult < val(XA(Grid1.Bookmark, 3)) - val(XA(Grid1.Bookmark, 4)) Then
                MarkRowsValid 2, cCOLDELL(val(XA(Grid1.Bookmark, 10))).Key
            ElseIf lngResult = val(XA(Grid1.Bookmark, 3)) - val(XA(Grid1.Bookmark, 4)) Then
                MarkRowsValid 1, cCOLDELL(val(XA(Grid1.Bookmark, 10))).Key
            End If
            cmdClose.Enabled = True
            For i = 1 To rs.RecordCount
                If rs.Fields(1) < 0 Then
                    Me.cmdClose.Enabled = False
                    Exit For
                End If
            Next
            cCOLDELL.GetStatus
         '   cCOLDELL(val(XA(Grid1.Bookmark, 10))).status = "A"
        '    EnableOK iOK = 1
       '     MarkRowsValid iOK, oInv.InvoiceLines(val(XA(Grid1.Bookmark, 10))).Key
        Else
            cancel = True
        End If
    End Select
   ' On Error GoTo 0
    Exit Sub
Errh:
    MsgBox Error
    Exit Sub
    Resume
End Sub
Private Sub EnableOK(pOK As Boolean)
    Me.cmdGenInv.Enabled = pOK
End Sub
Private Sub MarkRowsValid(pOK As Integer, pKey As String)
Dim i As Integer
    For i = 1 To lngArrayRows
        If XA(i, 10) = pKey Then
            Select Case pOK
                Case 1
                    XA(i, 12) = "OK"
                Case 2
                    XA(i, 12) = "MORE"
                Case 3
                    XA(i, 12) = "INVALID"
                Case 4
                    XA(i, 12) = "TOOMUCH"
            End Select
        End If
    Next i
End Sub
Private Sub Grid1_LostFocus()
  '  Grid1.ReBind
End Sub

Private Sub Grid1_SelChange(cancel As Integer)
  ' Grid1.SelStart = 1
  '  Grid1.SelLength = 99
End Sub
