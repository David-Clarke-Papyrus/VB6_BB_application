VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBranchMatchSalesReport 
   Caption         =   "Branch loyalty sales match report"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   7860
   Begin VB.TextBox txtUntil 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   5850
      TabIndex        =   9
      Top             =   525
      Width           =   1050
   End
   Begin VB.TextBox txtSince 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   4380
      TabIndex        =   7
      Top             =   525
      Width           =   1065
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Get exchanges (ignores duplicates)"
      Height          =   360
      Left            =   2595
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   945
      UseMaskColor    =   -1  'True
      Width           =   4305
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   1215
      TabIndex        =   3
      Top             =   570
      Width           =   675
   End
   Begin VB.TextBox txtBranchcode 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   345
      TabIndex        =   2
      Top             =   600
      Width           =   885
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   90
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmBranchMatchSalesReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4950
      UseMaskColor    =   -1  'True
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Bindings        =   "frmBranchMatchSalesReport.frx":038A
      Height          =   1515
      Left            =   135
      OleObjectBlob   =   "frmBranchMatchSalesReport.frx":039F
      TabIndex        =   0
      Top             =   1470
      Width           =   6765
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "and"
      Height          =   285
      Left            =   5415
      TabIndex        =   10
      Top             =   555
      Width           =   330
   End
   Begin VB.Label lblSince 
      Alignment       =   1  'Right Justify
      Caption         =   "Between (e.g. 13/2/2011)"
      Height          =   285
      Left            =   2145
      TabIndex        =   8
      Top             =   555
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Exchanges on branch but missing on Central"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   75
      TabIndex        =   6
      Top             =   150
      Width           =   4440
   End
   Begin VB.Label lblRecs 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1125
      TabIndex        =   4
      Top             =   4935
      Width           =   3690
   End
End
Attribute VB_Name = "frmBranchMatchSalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As New XArrayDB
Dim i As Long
Dim iMax As Long
Dim rs As ADODB.Recordset
Dim xMLDoc As New ujXML

Public Sub Component(rs As ADODB.Recordset, lngQtyRecsFound As Long)
    iMax = 0
    Do While Not rs.EOF
        iMax = iMax + 1
        rs.MoveNext
    Loop
    lngQtyRecsFound = iMax
    x.ReDim 1, iMax, 1, 6
    rs.MoveFirst
    For i = 1 To iMax
        x(i, 1) = FNS(rs.Fields(3))
        x(i, 2) = FNS(rs.Fields(2))
        x(i, 3) = FNS(rs.Fields(5))
        x(i, 4) = FNS(rs.Fields(4))
        x(i, 6) = FNS(rs.Fields(1))
        rs.MoveNext
    Next
    Set G.Array = x
    G.ReBind
    G.Refresh
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdGo_Click()
Dim oSQL As New z_SQL
Dim lngQtyRecsFound As Long

    If oPC.Configuration.Stores.FindStoreByCode(Me.txtBranchcode) Is Nothing Then
        MsgBox "This store does not exist.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    If oPC.Configuration.Stores.FindStoreByCode(Me.txtBranchcode).IsActive = False Then
        MsgBox "This store is marked as inactive.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    Set x = Nothing
    Set x = New XArrayDB
    MsgBox "This may take more than a minute. Please wait", vbOKOnly + vbInformation, "Warning"
    Screen.MousePointer = vbHourglass
    Set rs = oSQL.MatchExchangeRecs(txtBranchcode)
    If rs.EOF Then
        Screen.MousePointer = vbDefault
        G.ReBind
        G.Refresh
        MsgBox "There are no records to display. Check you have the correct branch code and the correct VPN address in the store record", vbOKOnly + vbInformation
        Set x = Nothing
        Exit Sub
    End If
    Me.Component rs, lngQtyRecsFound
    Me.lblRecs.Caption = CStr(lngQtyRecsFound) & " records found"
    Screen.MousePointer = vbDefault
    Me.cmdSend.Enabled = True
End Sub

Private Sub cmdSelectAll_Click()

    For i = 1 To x.UpperBound(1)
        x(i, 5) = "-1"
    Next
    
    G.Refresh
    
End Sub
Private Sub cmdUnselect_Click()

    For i = 1 To x.UpperBound(1)
        x(i, 5) = "0"
    Next
    
    G.Refresh

End Sub
Private Sub cmdSend_Click()
    On Error GoTo errHandler
Dim oSQL As New z_SQL
Dim oSM As New z_StockManager
    If Not (IsDate(txtSince) And IsDate(txtUntil)) Or Me.txtBranchcode = "" Then
        MsgBox "Problem in date formats or no branch selected"
    Else
        oSQL.SendInvocation "SalesSet", txtBranchcode, "", ReverseDate(CDate(txtSince)), ReverseDate(CDate(txtUntil))
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBranchMatchSalesReport.cmdSend_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub Command1_Click()
    MsgBox x(1, 4)
    MsgBox x(1, 1)
End Sub

Private Function CreateXMLListOfAxchangeNumbers() As String
Dim i As Integer
    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "ExchangeNoSelection"
            .chCreate "MessageType"
                .elText = "ExchangeNoSelection"

            For i = 1 To x.UpperBound(1)
                If x(i, 5) = "-1" Then
                    .elCreateSibling "DL", True
                    .chCreate "Exch"
                        .elText = x(i, 1)
                    .navUP
                End If
            Next i
    End With
    CreateXMLListOfAxchangeNumbers = xMLDoc.docXML

End Function


Private Sub Form_Load()
'    Me.Width = 6555
'    Me.Height = 4185
    Top = 2000
    Left = 500
End Sub


Private Sub G_HeadClick(ByVal ColIndex As Integer)
Static Direction As Variant

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    If ColIndex = 2 Then
        x.QuickSort x.LowerBound(1), x.UpperBound(1), 4, Direction, GetRowType(4)
    Else
        x.QuickSort x.LowerBound(1), x.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    End If
    
    G.Refresh
    Screen.MousePointer = vbDefault

End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    Select Case ColIndex
        Case 1, 2, 3
            GetRowType = XTYPE_STRING
        Case 4
            GetRowType = XTYPE_DATE
    End Select
End Function

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    G.Width = Me.Width - 800
    G.Height = Me.Height - 2850
    cmdClose.Top = NonNegative_Lng(Me.Height - 1200)
    lblRecs.Top = NonNegative_Lng(Me.Height - 1200)
End Sub

