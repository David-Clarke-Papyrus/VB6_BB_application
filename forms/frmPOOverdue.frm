VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmPOOverdue 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Purchase orders overdue"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   2490
      Left            =   0
      OleObjectBlob   =   "frmPOOverdue.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   11100
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00706034&
      Height          =   2580
      Left            =   0
      Top             =   0
      Width           =   11130
   End
End
Attribute VB_Name = "frmPOOverdue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oRS As ADODB.Recordset
Dim cPOLOS As c_POLsOS
Attribute cPOLOS.VB_VarHelpID = -1
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim iRecs As Integer
Dim lngArrayRows As Long
'Private Sub cCOLDELL_Valid(pResult As String)
''MsgBox pOK
'    EnableOK (pResult = "")
'    lblResult.Caption = pResult
'End Sub
Public Sub Component(pcPOLOS As c_POLsOS)
Dim strSQL As String
    Set cPOLOS = pcPOLOS
    cPOLOS.beginedit
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
    lngArrayRows = 0
'    For Each oCOLDELL In cCOLDELL
        lngArrayRows = cCOLDELL.Count
'    Next
    XA.ReDim 1, lngArrayRows, 1, 12
    
    For Each oCOLDELL In cCOLDELL
        lngAvailableToAllocate = oCOLDELL.QtyOnHand
'        For Each oCoff In oIL.COFFs
'            lngAvailableToAllocate = lngAvailableToAllocate - oCoff.COFFQTY
'        Next
'        lngArrayRows = lngArrayRows + oIL.COFFs.Count
            XA.Value(lngIndex, 1) = oCOLDELL.Titleshort(15) & "  (" & oCOLDELL.QtyOnHand & ")" '& "(" & oCOLDELL.QtyJustReceived & ")"
            XA.Value(lngIndex, 2) = oCOLDELL.CustomerName
            XA.Value(lngIndex, 3) = oCOLDELL.OrderDetails
            XA.Value(lngIndex, 4) = oCOLDELL.OrderedQty
            XA.Value(lngIndex, 5) = oCOLDELL.DeliveredSoFar
            XA.Value(lngIndex, 6) = oCOLDELL.QtyOS
            XA.Value(lngIndex, 7) = 0
            XA.Value(lngIndex, 8) = oCOLDELL.COLID
         '   XA.Value(lngIndex, 9) = oCOLDELL.DELLID
            XA.Value(lngIndex, 10) = oCOLDELL.Key
            lngIndex = lngIndex + 1
    Next
    XA.QuickSort 1, qtyRecs, 2, XORDER_ASCEND, XTYPE_DATE
    Grid1.Array = XA
End Sub

Private Sub cmdCancel_Click()
    If cCOLDELL.IsEditing Then cCOLDELL.CancelEdit
    Set cCOLDELL = Nothing
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim strError As String
 '   cCOLDELL.ApplyEdit strError
    cCOLDELL.Save strError
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
            cCOLDELL(val(XA(Grid1.Bookmark, 10))).SetAllocatedQty lngResult
            If lngResult > val(XA(Grid1.Bookmark, 4)) Then
                MarkRowsValid 4, cCOLDELL(val(XA(Grid1.Bookmark, 10))).Key
            ElseIf lngResult < val(XA(Grid1.Bookmark, 4)) Then
                MarkRowsValid 2, cCOLDELL(val(XA(Grid1.Bookmark, 10))).Key
            ElseIf lngResult = val(XA(Grid1.Bookmark, 4)) Then
                MarkRowsValid 1, cCOLDELL(val(XA(Grid1.Bookmark, 10))).Key
            End If
         '   cCOLDELL(val(XA(Grid1.Bookmark, 10))).status = "A"
        '    EnableOK iOK = 1
       '     MarkRowsValid iOK, oInv.InvoiceLines(val(XA(Grid1.Bookmark, 10))).Key
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
    LoadGrid
End Sub

Private Sub Grid1_LostFocus()
    Grid1.ReBind
End Sub

Private Sub Grid1_SelChange(Cancel As Integer)
  ' Grid1.SelStart = 1
  '  Grid1.SelLength = 99
End Sub

