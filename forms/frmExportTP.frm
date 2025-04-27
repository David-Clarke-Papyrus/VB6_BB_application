VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmExportTP 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Export data from trading partners"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select"
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
      Height          =   735
      Left            =   4215
      TabIndex        =   3
      Top             =   15
      Width           =   3255
      Begin VB.OptionButton optSupplier 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Supplier"
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
         Height          =   270
         Left            =   1965
         TabIndex        =   5
         Top             =   315
         Width           =   1020
      End
      Begin VB.OptionButton optCustomer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Customer"
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
         Height          =   270
         Left            =   495
         TabIndex        =   4
         Top             =   315
         Width           =   1380
      End
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Get data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5100
      Width           =   1260
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   4275
      Left            =   105
      OleObjectBlob   =   "frmExportTP.frx":0000
      TabIndex        =   2
      Top             =   795
      Width           =   7395
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fields"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   105
      TabIndex        =   0
      Top             =   420
      Width           =   1710
   End
End
Attribute VB_Name = "frmExportTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim rs As New ADODB.Recordset
Dim rsD As ADODB.Recordset
Dim frmTP As New frmTPTemplate
Dim lngFldCount As Long
Dim lngIndex As Long

 
Private Sub cmdGo_Click()
    G1.Update
    PrepareGrid
    If FetchData(IIf(optCustomer = True, "CUSTOMER", "SUPPLIER")) Then
        LoadGrid
        frmTP.component XA, XB, IIf(optCustomer = True, "CUSTOMER", "SUPPLIER")
        frmTP.Show vbModal
    End If
End Sub
Private Sub PrepareGrid()
Dim Col As TrueOleDBGrid60.Column
Dim i As Long
   ' For i = XA.UpperBound(1) To 1 Step -1
   For i = 1 To frmTP.GT.Splits(0).Columns.Count
        frmTP.GT.Splits(0).Columns.Remove (0)
   Next
    For i = XA.UpperBound(1) To 1 Step -1
        If XA(i, 3) = True Then
            Set Col = frmTP.GT.Splits(0).Columns.Add(0)
            Col.Caption = XA(i, 2)
            Col.Visible = True
        End If
    Next i
End Sub
Private Function FetchData(pType As String) As Boolean
Dim strSQL As String
Dim strFields As String
Dim strWhere As String
    FetchData = True
    strFields = ""
    strWhere = ""
    lngFldCount = 0
    For lngIndex = 1 To XA.UpperBound(1)
        If XA(lngIndex, 3) = True Then
            strFields = strFields & "," & XA(lngIndex, 5)   '5 is the field column
            lngFldCount = lngFldCount + 1
        End If
        If XA(lngIndex, 4) > "" Then
            strWhere = strWhere & XA(lngIndex, 5) & " " & ParseExpression(XA(lngIndex, 5), XA(lngIndex, 4)) '4 is the Where column
        End If
    Next
    
    
    strFields = right(strFields, Len(strFields) - 1)
    
    'Finalize WHERE clause
    If UCase(pType) = "CUSTOMER" Then
        If strWhere > "" Then strWhere = " WHERE TP_ROLE = 3 AND " & strWhere
    ElseIf UCase(pType) = "SUPPLIER" Then
        If strWhere > "" Then strWhere = " WHERE TP_ROLE = 2 AND " & strWhere
    End If
    
    strSQL = "SELECT " & strFields & " FROM tTP JOIN tADD on ADD_TP_ID = TP_ID LEFT JOIN vSalesSumm1 on SCY_TP_ID = TP_ID" & strWhere
    Set rsD = New ADODB.Recordset
    rsD.CursorLocation = adUseClient
On Error GoTo ERRH2
    rsD.Open strSQL, oPC.CO, adOpenForwardOnly
    Exit Function
ERRH2:
    MsgBox "One or more filters are incorrectly expressed"
    FetchData = False
    
End Function
Private Function ParseExpression(pFieldName As String, pExp As String) As String
    pExp = Replace(pExp, " AND ", " AND " & pFieldName & " ")
    ParseExpression = Replace(pExp, " OR ", " OR " & pFieldName & " ")
End Function
Private Sub LoadGrid()
Dim j As Long
    Set XB = New XArrayDB
    XB.Clear
    XB.ReDim 1, rsD.RecordCount, 1, lngFldCount
    lngIndex = 1
    If rsD.EOF Then Exit Sub
    rsD.MoveFirst
    For lngIndex = 1 To rsD.RecordCount
        For j = 1 To lngFldCount
            XB.Value(lngIndex, j) = FNS(rsD.Fields(j - 1))
        Next j
        If Not rsD.EOF Then rsD.MoveNext
    Next
    XB.QuickSort 1, rsD.RecordCount, 1, XORDER_ASCEND, XTYPE_STRING  ', 4, XORDER_ASCEND, XTYPE_DATE
   
End Sub
Private Sub Form_Load()
    LoadG1
End Sub
Private Sub LoadG1()
On Error GoTo Errh

Dim lngIndex As Long
    rs.CursorLocation = adUseClient
    rs.Open "Select * from tDmpTP", oPC.CO, adOpenKeyset, adLockOptimistic
    
    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, rs.RecordCount, 1, 6
    lngIndex = 1
    
    For lngIndex = 1 To rs.RecordCount
        XA.Value(lngIndex, 1) = lngIndex
        XA.Value(lngIndex, 2) = FNS(rs.Fields("Name"))
        XA.Value(lngIndex, 3) = FNB(rs.Fields("Selected"))
        XA.Value(lngIndex, 4) = ""
        XA.Value(lngIndex, 5) = FNS(rs.Fields("Field"))
        XA.Value(lngIndex, 6) = FNS(rs.Fields("Type"))
        rs.MoveNext
    Next
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_LONG  ', 4, XORDER_ASCEND, XTYPE_DATE
    G1.Array = XA
    G1.ReBind
   ' G1.SetFocus
    Exit Sub
Errh:
    MsgBox Error
    Resume
End Sub

Private Sub Form_Unload(cancel As Integer)
    rs.Close
    Set rs = Nothing
    Set XA = Nothing
    Unload frmTP
    Set frmTP = Nothing
End Sub

Private Sub G1_AfterColUpdate(ByVal colINdex As Integer)
    If colINdex + 1 <> 1 Then Exit Sub
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_LONG  ', 4, XORDER_ASCEND, XTYPE_DATE
    G1.Refresh
End Sub


Private Sub G1_BeforeColUpdate(ByVal colINdex As Integer, OldValue As Variant, cancel As Integer)
    Select Case colINdex
    Case 0
        If (Not IsNumeric(G1.Text)) Then
            cancel = True
            Exit Sub
        End If
        XA(G1.Bookmark, colINdex + 1) = CLng(G1.Text)
    Case Else
        XA(G1.Bookmark, colINdex + 1) = G1.Text
    End Select
End Sub

