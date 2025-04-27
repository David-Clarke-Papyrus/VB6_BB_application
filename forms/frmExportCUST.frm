VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmExportCUST 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Export data from customers"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   2070
      TabIndex        =   3
      Text            =   "500"
      Top             =   5115
      Width           =   840
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
      Left            =   6255
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5100
      Width           =   1260
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   4680
      Left            =   90
      OleObjectBlob   =   "frmExportCUST.frx":0000
      TabIndex        =   2
      Top             =   390
      Width           =   7395
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Max records to return"
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
      Height          =   330
      Left            =   90
      TabIndex        =   4
      Top             =   5130
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Top             =   90
      Width           =   1710
   End
End
Attribute VB_Name = "frmExportCUST"
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
Dim bMore As Boolean
    Screen.MousePointer = vbHourglass
    G1.Update
    PrepareGrid
    If FetchData(CLng(txtMax), bMore) Then
        LoadGrid
        frmTP.Component XA, XB, "CUSTOMER"
        frmTP.Show vbModal
    End If
    Screen.MousePointer = vbDefault
    If bMore Then MsgBox "There are more records to return, increase the value in the 'maximum records to return' box and fetch again."
End Sub
Private Sub PrepareGrid()
Dim Col As TrueOleDBGrid60.Column
Dim i As Long
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
Private Function FetchData(pMaxRecords As Long, pbMore As Boolean) As Boolean
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
            strWhere = strWhere & " AND " & XA(lngIndex, 5) & " " & ParseExpression(XA(lngIndex, 5), XA(lngIndex, 4)) '4 is the Where column
        End If
    Next
    
    
    strFields = right(strFields, Len(strFields) - 1)
    
    If left(strWhere, 5) = " AND " Then
        strWhere = right(strWhere, Len(strWhere) - 4)
    End If
    
    If strWhere > "" Then strWhere = " WHERE TP_ROLE = 3 AND " & strWhere
    strWhere = Replace(UCase(strWhere), "TRUE", "1")
    strWhere = Replace(UCase(strWhere), "FALSE", "0")
    strWhere = Replace(UCase(strWhere), "*", "%")
    strWhere = Replace(UCase(strWhere), """", "'")
    
    strSQL = "SELECT " & strFields & " FROM tTP JOIN tADD on ADD_TP_ID = TP_ID LEFT JOIN vSalesSumm1 on SCY_TP_ID = TP_ID LEFT JOIN tDICT on TP_CT_ID = DICT_ID " & strWhere
    Set rsD = New ADODB.Recordset
    rsD.CursorLocation = adUseClient
    rsD.MaxRecords = pMaxRecords
On Error GoTo ERRH2
    rsD.Open strSQL, oPC.COSHORT, adOpenForwardOnly
    pbMore = (rsD.RecordCount > pMaxRecords)
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
Dim k As Long
    Set XB = New XArrayDB
    XB.Clear
    XB.ReDim 1, rsD.RecordCount, 1, lngFldCount
    lngIndex = 1
    If rsD.EOF Then Exit Sub
    rsD.MoveFirst
    For lngIndex = 1 To rsD.RecordCount
        k = 0
        For j = 1 To XA.UpperBound(1)
            If XA(j, 3) = True Then
                k = k + 1
                If UCase(XA(j, 6)) = "CURR" Then
                    XB.Value(lngIndex, k) = Format(FNN(rsD.Fields(k - 1)) / 100, "##0.00")
                Else
                    XB.Value(lngIndex, k) = FNS(rsD.Fields(k - 1))
                End If
            End If
        Next
        If Not rsD.EOF Then rsD.MoveNext

    Next lngIndex
    XB.QuickSort 1, rsD.RecordCount, 1, XORDER_ASCEND, XTYPE_STRING  ', 4, XORDER_ASCEND, XTYPE_DATE
   
End Sub
Private Sub Form_Load()
    LoadG1
End Sub
Private Sub LoadG1()
On Error GoTo Errh

Dim lngIndex As Long
    rs.CursorLocation = adUseClient
    rs.Open "Select * from tDmp WHERE Type = 'C'", oPC.COSHORT, adOpenKeyset, adLockOptimistic
    
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
        XA.Value(lngIndex, 6) = FNS(rs.Fields("DataType"))
        rs.MoveNext
    Next
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_LONG  ', 4, XORDER_ASCEND, XTYPE_DATE
    G1.Array = XA
    G1.ReBind
    Exit Sub
Errh:
    MsgBox Error
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rs.Close
    Set rs = Nothing
    Set XA = Nothing
    Unload frmTP
    Set frmTP = Nothing
End Sub

Private Sub G1_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex + 1 <> 1 Then Exit Sub
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_LONG  ', 4, XORDER_ASCEND, XTYPE_DATE
    G1.Refresh
End Sub


Private Sub G1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim lngTmp As Long
    Select Case ColIndex
    Case 0
        If (Not ConvertToLng(G1.Text, lngTmp)) Then
            Cancel = True
            Exit Sub
        End If
        XA(G1.Bookmark, ColIndex + 1) = CLng(G1.Text)
    Case Else
        XA(G1.Bookmark, ColIndex + 1) = G1.Text
    End Select
End Sub

Private Sub txtMax_Validate(Cancel As Boolean)
Dim lngTmp As Long

    If Not ConvertToLng(txtMax, lngTmp) Then Cancel = True
End Sub

