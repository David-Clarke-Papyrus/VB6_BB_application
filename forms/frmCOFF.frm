VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmCOFF 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Customer order line reconciliation"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10320
   FillColor       =   &H00FFC0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2670
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   8925
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2685
      Width           =   1200
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   2400
      Left            =   150
      OleObjectBlob   =   "frmCOFF.frx":0000
      TabIndex        =   0
      Top             =   135
      Width           =   9915
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00706034&
      Height          =   2430
      Left            =   135
      Top             =   120
      Width           =   9990
   End
End
Attribute VB_Name = "frmCOFF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oRS As ADODB.Recordset
Dim oInv As a_Invoice
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim iCoffs As Integer
Dim iRecs As Integer
Dim lngArrayRows As Long

Public Sub Component(pinv As a_Invoice)
Dim strSQL As String
Dim lngINVID As Long
    Set oInv = pinv
    oInv.BeginEdit
End Sub

Private Sub LoadCOLs()
Dim objItm As ListItem
Dim lngIndex As Long
Dim tmp As String
Dim qtyRecs As Long
Dim lngAwaiting As Long
Dim lngAllocation As Long
Dim lngAvailableToAllocate As Long
Dim i As Integer
Dim oIL As a_InvoiceLine
Dim oCoff As a_Coff

    i = 0
    Set XA = New XArrayDB
    XA.Clear
    iRecs = i
    lngIndex = 1
    lngArrayRows = 0
    For Each oIL In oInv.InvoiceLines
        lngArrayRows = oIL.COFFs.Count
    Next
    XA.ReDim 1, lngArrayRows, 1, 12
    
    For Each oIL In oInv.InvoiceLines
        lngAvailableToAllocate = oIL.Qty
        For Each oCoff In oIL.COFFs
            lngAvailableToAllocate = lngAvailableToAllocate - oCoff.COFFQTY
        Next
        lngArrayRows = lngArrayRows + oIL.COFFs.Count
        For Each oCoff In oIL.COFFs
            XA.Value(lngIndex, 1) = oIL.CodeF & ": " & Left(oIL.Title, 21) & "    (" & oIL.QtyF & ")"
            XA.Value(lngIndex, 2) = oCoff.CODateF
            XA.Value(lngIndex, 3) = oCoff.COCode
            XA.Value(lngIndex, 4) = oCoff.COLQty
           ' oCoff.COFFQTY
            lngAwaiting = oCoff.COLQty - oCoff.COLQtyDispatched
            XA.Value(lngIndex, 6) = lngAwaiting
            lngAllocation = GetMin(lngAwaiting, lngAvailableToAllocate)
            oCoff.COFFCOLID = oCoff.COLID
            oCoff.COFFILID = oIL.InvoiceLineID
            oCoff.SETCOFFQTY lngAllocation
            lngAvailableToAllocate = lngAvailableToAllocate - lngAllocation
            XA.Value(lngIndex, 5) = oCoff.COLQtyDispatched
            XA.Value(lngIndex, 7) = lngAllocation
            XA.Value(lngIndex, 8) = oCoff.COLID
            XA.Value(lngIndex, 9) = oIL.InvoiceLineID
            XA.Value(lngIndex, 10) = oIL.Key
            XA.Value(lngIndex, 11) = oCoff.Key
            lngIndex = lngIndex + 1
        Next
    Next
    XA.QuickSort 1, qtyRecs, 2, XORDER_ASCEND, XTYPE_DATE
    Grid1.Array = XA
End Sub

Private Sub cmdCancel_Click()
    If oInv.IsEditing Then oInv.CancelEdit
    Set oInv = Nothing
    Unload Me
End Sub

Private Sub cmdOK_Click()
    oInv.ApplyEdit
    Unload Me
End Sub

Private Sub Grid1_AfterUpdate()
    Grid1.ReBind
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
Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim i As Integer
Dim oC As a_Copy
Dim lngResult As Long
Dim iOK As Integer
    i = ColIndex + 1
    On Error Resume Next
    Select Case i
    Case 7
        If ConvertToLng(Grid1.Text, lngResult) Then
            Grid1.Text = CStr(lngResult)
            oInv.InvoiceLines(val(XA(Grid1.Bookmark, 10))).COFFs(val(XA(Grid1.Row + 1, 11))).SETCOFFQTY lngResult
            oInv.InvoiceLines(val(XA(Grid1.Bookmark, 10))).CoffsValid iOK
            EnableOK iOK = 1
            MarkRowsValid iOK, oInv.InvoiceLines(val(XA(Grid1.Bookmark, 10))).Key
        Else
            Cancel = True
        End If
    End Select
    On Error GoTo 0

End Sub
Private Sub EnableOK(pOK As Boolean)
    cmdOK.Enabled = pOK
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
Private Sub Form_Load()
    LoadCOLs
End Sub

Private Sub Grid1_LostFocus()
    Grid1.ReBind
End Sub

Private Sub Grid1_SelChange(Cancel As Integer)
  ' Grid1.SelStart = 1
  '  Grid1.SelLength = 99
End Sub
