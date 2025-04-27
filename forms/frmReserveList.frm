VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmReserveList 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Reserve List"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Height          =   570
      Left            =   2085
      TabIndex        =   3
      Top             =   5010
      Width           =   4920
      Begin VB.OptionButton optReturned 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Returned to stock recently"
         Height          =   315
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   2115
      End
      Begin VB.OptionButton optCollected 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Collected recently"
         Height          =   315
         Left            =   1170
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   1440
      End
      Begin VB.OptionButton optReserved 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Reserved"
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdPrintList 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print list"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5010
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   570
      Left            =   9630
      Picture         =   "frmReserveList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5025
      Width           =   1035
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   4920
      Left            =   0
      OleObjectBlob   =   "frmReserveList.frx":00AB
      TabIndex        =   0
      Top             =   15
      Width           =   10680
   End
End
Attribute VB_Name = "frmReserveList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XA As XArrayDB
Dim rs As ADODB.Recordset
Private Enum enMode
    eReservations = 1
    eCollections = 2
    eReturnsToStock = 3
End Enum
Dim iMode As enMode

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    G1.Update
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub G1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
    If iMode = eReservations Then
       If Button = 2 Then   ' Check if right mouse button
                            ' was clicked.
          G1.Update
          PopupMenu Forms(0).mnuReserveList   ' Display the File menu as a
                            ' pop-up menu.
       End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.G1_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrintList_Click()
    On Error GoTo errHandler
Dim ar As New arReservations

    ar.component XA
    ar.Show vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.cmdPrintList_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    If Me.WindowState <> 2 Then
        Me.TOP = 1060
        Left = 120
    End If
    FetchReserved eReservations
    LoadGrid
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub FetchReserved(pMode As enMode)
    On Error GoTo errHandler
Dim OpenResult As Integer

    Set rs = Nothing
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If pMode = eReservations Then
        rs.open "Select * FROM vReservedStock WHERE STATUS = 1", oPC.COShort, adOpenForwardOnly, adLockOptimistic
    ElseIf pMode = eCollections Then
        rs.open "Select * FROM vReservedStock WHERE STATUS = 2 AND STATUSCHANGEDATE > '" & ReverseDate(DateAdd("m", -1, Date)) & "'", oPC.COShort, adOpenForwardOnly, adLockOptimistic
    ElseIf pMode = eReturnsToStock Then
        rs.open "Select * FROM vReservedStock WHERE STATUS = 3 AND STATUSCHANGEDATE > '" & ReverseDate(DateAdd("m", -1, Date)) & "'", oPC.COShort, adOpenForwardOnly, adLockOptimistic
    End If
    iMode = pMode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.FetchReserved(pMode)", pMode
End Sub

Private Sub RefreshList(pMode As enMode)
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    FetchReserved pMode
    LoadGrid
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.RefreshList(pMode)", pMode
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim lngArrayRows As Long
    If rs.eof Then
        If Not XA Is Nothing Then
            XA.Clear
            XA.ReDim 1, 0, 1, 16
            G1.ReBind
            G1.Refresh
        End If
        Exit Sub
    End If
    rs.MoveFirst
    lngArrayRows = rs.RecordCount
    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, lngArrayRows, 1, 19
    lngIndex = 1
    Do While lngIndex <= lngArrayRows
        XA.Value(lngIndex, 1) = FNS(rs.fields("TITLE"))
        XA.Value(lngIndex, 2) = FNN(rs.fields("QTYOH"))
        XA.Value(lngIndex, 3) = FNN(rs.fields("QTYRES"))
        XA.Value(lngIndex, 4) = FNS(rs.fields("CUSTTITLE")) & " " & FNS(rs.fields("CUSTINITIALS")) & " " & FNS(rs.fields("CUSTNAME")) & " (ph:" & FNS(rs.fields("PHONE")) & ")"
        XA.Value(lngIndex, 5) = FNN(rs.fields("QTYORDERED"))
      '  XA.Value(lngIndex, 6) = FNN(rs.Fields("COL_QtyDispatched"))
        XA.Value(lngIndex, 7) = FNN(rs.fields("QTYALLOC"))
        XA.Value(lngIndex, 8) = Format(rs.fields("ORDERDATE"), "dd/mm/yyyy")
        XA.Value(lngIndex, 9) = FNS(rs.fields("ORDERDOCNO"))  ' & " " & FNS(rs.Fields("TR_Date")) & " " & FNS(rs.Fields("COL_REF"))
        XA.Value(lngIndex, 10) = FNS(rs.fields("NOTE"))
     '   XA.Value(lngIndex, 11) = FND(rs.Fields("COL_LastActionDate"))
        XA.Value(lngIndex, 12) = FND(rs.fields("STATUSCHANGEDATE"))
        XA.Value(lngIndex, 13) = FNS(rs.fields("P_ID"))
        XA.Value(lngIndex, 14) = FNN(rs.fields("COL_ID"))
        XA.Value(lngIndex, 15) = FNS(rs.fields("ITEMCODE"))
        XA.Value(lngIndex, 16) = FNN(rs.fields("COLALLOC_ID"))
        XA.Value(lngIndex, 17) = Format(FND(rs.fields("DELIVERYDATE")), "Short date")
        XA.Value(lngIndex, 18) = FNS(rs.fields("PHONE"))
        XA.Value(lngIndex, 19) = FNS(rs.fields("ACNO"))
        lngIndex = lngIndex + 1
        rs.MoveNext
    Loop
    XA.QuickSort 1, lngArrayRows, 1, XORDER_ASCEND, XTYPE_STRING
    G1.Array = XA
    G1.ReBind
    G1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.LoadGrid"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
Dim i As Long
Dim OpenResult As Integer

    If XA Is Nothing Then Exit Sub
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    For i = 1 To XA.UpperBound(1)
        If XA.Value(i, 10) > "" Then
            oPC.COShort.execute "UPDATE tCOLALLOC SET COLALLOC_NOTE =  '" & XA.Value(i, 10) & "' WHERE COLALLOC_ID = " & XA.Value(i, 16)
        End If
    Next
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Set XA = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub G1_Click()
    On Error GoTo errHandler
    If IsNull(G1.Bookmark) Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    
    Clipboard.SetText Left(FNS(XA(G1.Bookmark, 15)), ISBNLENGTH)
 '   txtNote = XA.Value(G1.Bookmark, 4)
 '   txtAction = XA.Value(G1.Bookmark, 10)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.G1_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub CustomerCollects()
    On Error GoTo errHandler
Dim frm As New frmOffReserve
Dim strNote As String
Dim lngQty As Long
Dim OpenResult As Integer

    frm.component XA.Value(G1.Bookmark, 1) & " for " & XA.Value(G1.Bookmark, 4), CLng(XA.Value(G1.Bookmark, 7))
    frm.Show vbModal
    If frm.GetChoice <> "Yes" Then
        Exit Sub
    Else
        strNote = frm.GetNote
    End If
    lngQty = frm.GetQty
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If XA.Value(G1.Bookmark, 14) > 0 Then
        oPC.COShort.execute "UPDATE tCOL SET COL_NOTE = '" & strNote & "' WHERE COL_ID = " & XA.Value(G1.Bookmark, 14)
    End If
    oPC.COShort.execute "UPDATE tProduct SET P_QtyReserved = dbo.NonNegative(P_QtyReserved - " & lngQty & ") WHERE P_ID = '" & XA.Value(G1.Bookmark, 13) & "'"
    oPC.COShort.execute "UPDATE tStoreP SET STP_QTYRESERVED = dbo.NonNegative(STP_QTYRESERVED - " & lngQty & ") WHERE STP_P_ID = '" & XA.Value(G1.Bookmark, 13) & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
    oPC.COShort.execute "UPDATE tCOLALLOC  SET COLALLOC_NOTE = '" & "Collected: " & Format(Date, "dd/mm/yyyy") & "' + ISNULL(COLALLOC_NOTE,'') + '" & " " & strNote & " " & FNS(XA.Value(G1.Bookmark, 10)) & "', COLALLOC_STATUS = 2,COLALLOC_STATUSCHANGE_DATE = GETDATE() WHERE COLALLOC_ID = " & CLng(XA.Value(G1.Bookmark, 16))
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Unload frm
    G1.Delete
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.CustomerCollects"
End Sub
Public Sub ReturnToStock()
    On Error GoTo errHandler
Dim frm As New frmOffReserveNoCollection
Dim strNote As String
Dim lngQty As Long
Dim OpenResult As Integer

    frm.component XA.Value(G1.Bookmark, 1) & " for " & XA.Value(G1.Bookmark, 4), CLng(XA.Value(G1.Bookmark, 7))
    frm.Show vbModal
    If frm.GetChoice <> "Yes" Then
        Exit Sub
    End If
    lngQty = frm.GetQty
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "UPDATE tProduct SET P_QtyReserved = dbo.NonNegative(P_QtyReserved - " & lngQty & ") WHERE P_ID = '" & XA.Value(G1.Bookmark, 13) & "'"
    oPC.COShort.execute "UPDATE tStoreP SET STP_QTYRESERVED = dbo.NonNegative(STP_QTYRESERVED - " & lngQty & ") WHERE STP_P_ID = '" & XA.Value(G1.Bookmark, 13) & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
    oPC.COShort.execute "UPDATE tCOLALLOC SET COLALLOC_NOTE = '" & "Returned to stock: " & Format(Date, "dd/mm/yyyy") & "' + ISNULL(COLALLOC_NOTE,'') + '" & " " & FNS(XA.Value(G1.Bookmark, 10)) & "',COLALLOC_STATUS = 3 ,COLALLOC_STATUSCHANGE_DATE = GETDATE() WHERE COLALLOC_ID = " & CLng(XA.Value(G1.Bookmark, 16))
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Unload frm
    G1.Delete
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.ReturnToStock"
End Sub
'Private Sub G1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'   ' XA.Value(LastRow, 4) = txtNote
'    XA.Value(LastRow, 10) = txtAction
'    txtNote = XA.Value(G1.Bookmark, 4)
'    txtAction = XA.Value(G1.Bookmark, 10)
' '   If Not IsNumeric(LastRow) Then Exit Sub
' '   txtNote = FNS(XA.Value(LastRow, 9))
' '   txtAction = Format(FND(XA.Value(LastRow, 11)), "dd/mm/yyyy") & " " & FNS(XA.Value(LastRow, 10))
'End Sub

Private Sub G1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) 'XTYPE_INTEGER
    G1.Refresh

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1
            GetRowType = 4
        Case 2, 3
            GetRowType = 9
        Case Else
            GetRowType = 11
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.GetRowType(ColIndex)", ColIndex
End Function




Private Sub optCollected_Click()
    On Error GoTo errHandler
    If optCollected = True Then
        RefreshList eCollections
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.optCollected_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optReserved_Click()
    On Error GoTo errHandler
    If optReserved = True Then
        RefreshList eReservations
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.optReserved_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optReturned_Click()
    On Error GoTo errHandler
    If optReturned = True Then
        RefreshList eReturnsToStock
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReserveList.optReturned_Click", , EA_NORERAISE
    HandleError
End Sub
