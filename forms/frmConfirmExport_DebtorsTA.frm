VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmConfirmExport_DebtorsTA 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Confirm export"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   10260
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTotal 
      Height          =   315
      Left            =   6810
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "txtTotal"
      Top             =   3210
      Width           =   1335
   End
   Begin VB.CommandButton cmdTick 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Tick all"
      Height          =   315
      Left            =   4515
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3960
      Width           =   1305
   End
   Begin VB.CommandButton cmdUnTickSelected 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Un-tick selected"
      Height          =   315
      Left            =   2745
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   1305
   End
   Begin VB.CommandButton cmdTickSelected 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Tick selected"
      Height          =   315
      Left            =   1425
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3960
      Width           =   1305
   End
   Begin VB.Frame frDebtors 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Filter by document type"
      ForeColor       =   &H8000000D&
      Height          =   675
      Left            =   4005
      TabIndex        =   5
      Top             =   75
      Width           =   3525
      Begin VB.OptionButton optALL 
         BackColor       =   &H00D3D3CB&
         Caption         =   "All"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   2715
         TabIndex        =   8
         Top             =   300
         Width           =   555
      End
      Begin VB.OptionButton optCN 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Credit notes"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   1305
         TabIndex        =   7
         Top             =   300
         Width           =   1140
      End
      Begin VB.OptionButton optINV 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Invoices"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdUntick 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Un-tick all"
      Height          =   315
      Left            =   5835
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Default         =   -1  'True
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
      Left            =   330
      Picture         =   "frmConfirmExport_DebtorsTA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3675
      Width           =   1000
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Continue"
      CausesValidation=   0   'False
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
      Left            =   8355
      Picture         =   "frmConfirmExport_DebtorsTA.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3675
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Height          =   2160
      Left            =   255
      OleObjectBlob   =   "frmConfirmExport_DebtorsTA.frx":0714
      TabIndex        =   0
      Top             =   930
      Width           =   9135
   End
   Begin VB.Label lblSelectedStatus 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   270
      TabIndex        =   14
      Top             =   3180
      Width           =   5775
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   5820
      TabIndex        =   13
      Top             =   3270
      Width           =   870
   End
   Begin VB.Label lblLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Un-tick any documents you do not wish to export"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   300
      TabIndex        =   2
      Top             =   690
      Width           =   5475
   End
End
Attribute VB_Name = "frmConfirmExport_DebtorsTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As adodb.Recordset
Dim bCancelled As Boolean
Dim bIE As String
Dim bCustomerSupplier As String
Dim x As XArrayDB
Dim QtyInGrid As Long
Dim QtySelected As Long
Dim TotalValue As Double
Dim FilterType As String

Public Sub Component(pType As String)
    On Error GoTo errHandler
10        Screen.MousePointer = vbHourglass
            FilterType = pType
20        If FilterType = "DR" Then
30            frDebtors.Visible = True
40        Else
50            frDebtors.Visible = False
60        End If
70        lblLabel.Caption = "Un-tick any documents you do not wish to export"
80        Set rs = New adodb.Recordset
90        rs.CursorLocation = adUseClient
100       rs.Open "Select Dte,ProcessingDate,Acno,Reference,Descr,Amt,Action,RowID FROM tPASTEL ORDER BY ProcessingDate", oPC.CO, adOpenDynamic, adLockOptimistic
110       Set x = New XArrayDB
          
120       LoadRecordset
130       Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrorIn "frmConfirmExport_DebtorsTA.Component(pType)", pType
End Sub
Private Sub LoadRecordset()
    On Error GoTo errHandler
10        G.Update
20        If FilterType = "DR" Then
30            If optINV = True Then
40                    rs.Filter = "DESCR = 'TAX_INVOICE'"
50                    If rs.State = 1 Then rs.Close
60                    rs.Open "Select Dte,ProcessingDate,Acno,Reference,Descr,Amt,Action,RowID FROM tPASTEL WHERE DESCR = 'TAX_INVOICE' ORDER BY ProcessingDate", oPC.CO, adOpenDynamic
70            Else
80                If Me.optCN = True Then
90                    rs.Filter = "DESCR = 'TAX_CREDITNOTE'"
100                   If rs.State = 1 Then rs.Close
110                   rs.Open "Select Dte,ProcessingDate,Acno,Reference,Descr,Amt,Action,RowID FROM tPASTEL WHERE DESCR = 'TAX_CREDITNOTE' ORDER BY ProcessingDate", oPC.CO, adOpenDynamic
120               Else
130                   rs.Filter = ""
140                   If rs.State = 1 Then rs.Close
150                   rs.Open "Select Dte,ProcessingDate,Acno,Reference,Descr,Amt,Action,RowID FROM tPASTEL ORDER BY ProcessingDate", oPC.CO, adOpenDynamic
160               End If
170           End If
180       Else
190           rs.Filter = ""
200           If rs.State = 1 Then rs.Close
210           rs.Open "Select Dte,ProcessingDate,Acno,Reference,Descr,Amt,Action,RowID FROM tPASTEL ORDER BY ProcessingDate", oPC.CO, adOpenDynamic
220       End If
230       Set x = Nothing
240       Set x = New XArrayDB
250       LoadArray
260       G.Array = x
270       G.ReBind
280       G.Refresh
290       RedisplayTotals
    Exit Sub
errHandler:
    ErrorIn "frmConfirmExport_DebtorsTA.LoadRecordset"
End Sub
Private Sub LoadArray()
    On Error GoTo errHandler
      Dim i As Integer
      Dim T As String

10        For i = 1 To rs.RecordCount
20            x.ReDim 1, i, 1, 10
30            x(i, 1) = Format(FND(rs.Fields("Dte")), "dd/mm/yyyy")
40            x(i, 2) = Format(FND(rs.Fields("ProcessingDate")), "dd/mm/yyyy")
50            x(i, 3) = FNS(rs.Fields("Reference"))
60            x(i, 4) = FNS(rs.Fields("Acno"))
70            x(i, 5) = FNS(rs.Fields("Descr"))
80            x(i, 6) = Format(FNDBL(rs.Fields("Amt")), "###,##0.00")
90            x(i, 7) = FNB(rs.Fields("Action"))
100           x(i, 8) = StripToNumerics(FNS(rs.Fields("Reference")))
110           x(i, 9) = FNN(rs.Fields("RowID"))
120           rs.MoveNext
130       Next i
          
          
    Exit Sub
errHandler:
    ErrorIn "frmConfirmExport_DebtorsTA.LoadArray"
End Sub
Private Sub RecalculateGrid()
    On Error GoTo errHandler
Dim i As Integer
    If x.Count(1) = 0 Then Exit Sub
    QtyInGrid = 0
    QtySelected = 0
    TotalValue = 0
    QtyInGrid = x.UpperBound(1)
    For i = 1 To x.UpperBound(1)
        If x(i, 7) = True Then
            QtySelected = QtySelected + 1
            TotalValue = TotalValue + CDbl(x(i, 6))
        End If
    Next
    Exit Sub
errHandler:
    ErrorIn "frmConfirmExport_DebtorsTA.RecalculateGrid"
End Sub
Private Sub cmdTickSelected_Click()
Dim i As Integer
    G.Update
    For i = 1 To G.SelBookmarks.Count
        x(G.SelBookmarks(i - 1), 7) = 1
        rs.Find "RowID = " & x(G.SelBookmarks(i - 1), 9), , adSearchForward, 1
        rs.Fields("Action") = x(G.SelBookmarks(i - 1), 7)
        rs.Update
    Next
    RedisplayTotals
    G.Refresh
End Sub

Private Sub cmdUntick_Click()
Dim i As Integer
    G.Update
    For i = 1 To x.UpperBound(1)
        x(i, 7) = 0
        rs.Find "RowID = " & x(i, 9), , adSearchForward, 1
        rs.Fields("Action") = x(i, 7)
        rs.Update
    Next
    RedisplayTotals
    G.Refresh

End Sub

Private Sub cmdUnTickSelected_Click()
Dim i As Integer
    G.Update
    For i = 1 To G.SelBookmarks.Count
        x(G.SelBookmarks(i - 1), 7) = 0
        rs.Find "RowID = " & x(G.SelBookmarks(i - 1), 9), , adSearchForward, 1
        rs.Fields("Action") = x(G.SelBookmarks(i - 1), 7)
        rs.Update
    Next
    RedisplayTotals
    G.Refresh
End Sub

Private Sub cmdTick_Click()
Dim i As Integer
    G.Update
    For i = 1 To x.UpperBound(1)
        x(i, 7) = True
        rs.Find "RowID = " & x(i, 9), , adSearchForward, 1
        rs.Fields("Action") = x(i, 7)
        rs.Update
    Next
    RedisplayTotals
    G.Refresh
End Sub

Private Sub Form_Load()
    SetGridLayout G, Me.Name
    SetFormSize Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    G.Update
    SaveLayout G, Me.Name, Me.Height, Me.Width
    
End Sub

Private Sub G_AfterUpdate()
        rs.Find "RowID = " & x(G.Bookmark, 9), , adSearchForward, 1
        rs.Fields("Action") = x(G.Bookmark, 7)
        rs.Update
    RedisplayTotals
End Sub
Private Sub RedisplayTotals()
    RecalculateGrid
    lblSelectedStatus.Caption = "Rows in grid:" & CStr(QtyInGrid) & "    Qty selected: " & CStr(QtySelected)
    txtTotal = CStr(TotalValue)
End Sub
Private Sub G_HeadClick(ByVal ColIndex As Integer)
Static Direction As Variant
    Screen.MousePointer = vbHourglass

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    If ColIndex = 1 Then ColIndex = 7
 '   If ColIndex = 2 Then
 '       X.QuickSort X.LowerBound(1), X.UpperBound(1), 4, Direction, GetRowType(5) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
 '   Else
        x.QuickSort x.LowerBound(1), x.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1)
 '   End If
    
    G.Refresh
    Screen.MousePointer = vbDefault
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    Select Case ColIndex
        Case 3, 4, 5
            GetRowType = XTYPE_STRING
        Case 6, 8
            GetRowType = XTYPE_DOUBLE
        Case 1, 2
            GetRowType = XTYPE_DATE
    End Select
End Function


Private Sub G_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property

Private Sub OKButton_Click()
Dim i As Integer
Dim oB As z_Batch

    G.Update
    If x Is Nothing Then Exit Sub
    
    
    bCancelled = False
    Me.Hide
End Sub
Private Sub cmdCancel_Click()
    bCancelled = True
    Me.Hide
End Sub

Private Sub optall_Click()
    LoadRecordset
    RecalculateGrid
End Sub

Private Sub optCN_Click()
    LoadRecordset
    RecalculateGrid
End Sub

Private Sub optINV_Click()
    LoadRecordset
    RecalculateGrid

End Sub
Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    G.Width = Me.Width - (G.Left + 400)
    lngDiff = G.Height
    G.Height = Me.Height - (G.Top + 1720)
    lngDiff = G.Height - lngDiff
    cmdCancel.Top = cmdCancel.Top + lngDiff
    cmdTickSelected.Top = cmdCancel.Top
    cmdUnTickSelected.Top = cmdCancel.Top
    cmdUntick.Top = cmdCancel.Top
    cmdTick.Top = cmdCancel.Top
    OKButton.Top = cmdCancel.Top
    Me.lblSelectedStatus.Top = OKButton.Top - 500
    Me.txtTotal.Top = OKButton.Top - 550
    lblTotal.Top = OKButton.Top - 500
End Sub

