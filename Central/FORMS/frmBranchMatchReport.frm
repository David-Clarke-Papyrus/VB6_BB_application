VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBranchMatchReport 
   Caption         =   "Branch loyalty customer match report"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   7710
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   5700
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmBranchMatchReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4260
      UseMaskColor    =   -1  'True
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3570
      Left            =   165
      TabIndex        =   1
      Top             =   555
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   6297
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Missing on branch"
      TabPicture(0)   =   "frmBranchMatchReport.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblRecs"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "G"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSend"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdGo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtBranchcode"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Missing on Central"
      TabPicture(1)   =   "frmBranchMatchReport.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtBranchCode2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdGo2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdSend2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "G2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblRecs2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.TextBox txtBranchCode2 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   -74745
         TabIndex        =   9
         Top             =   510
         Width           =   885
      End
      Begin VB.CommandButton cmdGo2 
         Caption         =   "Go"
         Height          =   375
         Left            =   -73815
         TabIndex        =   8
         Top             =   480
         Width           =   675
      End
      Begin VB.CommandButton cmdSend2 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get selected from branch"
         Enabled         =   0   'False
         Height          =   360
         Left            =   -71400
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   495
         UseMaskColor    =   -1  'True
         Width           =   2625
      End
      Begin VB.TextBox txtBranchcode 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   255
         TabIndex        =   4
         Top             =   510
         Width           =   885
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go"
         Height          =   375
         Left            =   1185
         TabIndex        =   3
         Top             =   480
         Width           =   675
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Send all to Branch (300 in batch)"
         Enabled         =   0   'False
         Height          =   360
         Left            =   3615
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   465
         UseMaskColor    =   -1  'True
         Width           =   2625
      End
      Begin TrueOleDBGrid60.TDBGrid G 
         Bindings        =   "frmBranchMatchReport.frx":03C2
         Height          =   2310
         Left            =   255
         OleObjectBlob   =   "frmBranchMatchReport.frx":03D7
         TabIndex        =   5
         Top             =   915
         Width           =   6045
      End
      Begin TrueOleDBGrid60.TDBGrid G2 
         Bindings        =   "frmBranchMatchReport.frx":3CD9
         Height          =   2445
         Left            =   -74745
         OleObjectBlob   =   "frmBranchMatchReport.frx":3CEE
         TabIndex        =   10
         Top             =   900
         Visible         =   0   'False
         Width           =   6045
      End
      Begin VB.Label lblRecs2 
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   -74670
         TabIndex        =   11
         Top             =   3345
         Width           =   3690
      End
      Begin VB.Label lblRecs 
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   330
         TabIndex        =   6
         Top             =   3345
         Width           =   3690
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Loyalty customers on Central but not on branch"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   5985
   End
End
Attribute VB_Name = "frmBranchMatchReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As New XArrayDB
Dim X2 As New XArrayDB
Dim i As Long
Dim iMax As Long
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim xMLDoc As New ujXML

Public Sub Component(rs As ADODB.Recordset, lngQtyRecsFound As Long)
    iMax = 0
    Do While Not rs.EOF
        iMax = iMax + 1
        rs.MoveNext
    Loop
    lngQtyRecsFound = iMax
    x.ReDim 1, iMax, 1, 4
    rs.MoveFirst
    For i = 1 To iMax
        x(i, 1) = FNS(rs.Fields(0))
        x(i, 2) = FNS(rs.Fields(1))
        x(i, 3) = FNS(rs.Fields(3))
        x(i, 4) = FNS(rs.Fields(2))
        rs.MoveNext
    Next
End Sub
Public Sub Component2(rs As ADODB.Recordset, lngQtyRecsFound As Long)
    iMax = 0
    Do While Not rs.EOF
        iMax = iMax + 1
        rs.MoveNext
    Loop
    lngQtyRecsFound = iMax
    X2.ReDim 1, iMax, 1, 5
    rs.MoveFirst
    For i = 1 To iMax
        X2(i, 1) = FNS(rs.Fields(0))
        X2(i, 2) = FNS(rs.Fields(1))
        X2(i, 3) = FNS(rs.Fields(3))
        X2(i, 4) = FNS(rs.Fields(2))
        rs.MoveNext
    Next
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
    Set rs = oSQL.MatchLoyaltyCustomers_MissingAtBranch(txtBranchcode)
    If rs.EOF Then
        Screen.MousePointer = vbDefault
        Set G.Array = x
        G.ReBind
        G.Refresh
        MsgBox "There are no records to display. Check you have the correct branch code and the correct VPN address in the store record", vbOKOnly + vbInformation
        Set x = Nothing
        Exit Sub
    End If
    Me.Component rs, lngQtyRecsFound
    Set G.Array = x
    G.ReBind
    G.Refresh
    Me.lblRecs.Caption = CStr(lngQtyRecsFound) & " records found"
    Screen.MousePointer = vbDefault
    Me.cmdSend.Enabled = True
End Sub

Private Sub cmdGo2_Click()
Dim oSQL As New z_SQL
Dim lngQtyRecsFound As Long
    If oPC.Configuration.Stores.FindStoreByCode(Me.txtBranchCode2) Is Nothing Then
        MsgBox "This store does not exist.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    If oPC.Configuration.Stores.FindStoreByCode(Me.txtBranchCode2).IsActive = False Then
        MsgBox "This store is marked as inactive.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    Set X2 = Nothing
    Set X2 = New XArrayDB
    MsgBox "This may take more than a minute. Please wait", vbOKOnly + vbInformation, "Warning"
    Screen.MousePointer = vbHourglass
    Set rs2 = oSQL.MatchLoyaltyCustomers_MissingAtCentral(txtBranchCode2)
    If rs2.EOF Then
        Screen.MousePointer = vbDefault
        Set G2.Array = x
        G2.ReBind
        G2.Refresh
        MsgBox "There are no records to display. Check you have the correct branch code and the correct VPN address in the store record", vbOKOnly + vbInformation
        Set x = Nothing
        Exit Sub
    End If
    Me.Component2 rs2, lngQtyRecsFound
    Set G2.Array = X2
    G2.ReBind
    G2.Refresh
    Me.lblRecs2.Caption = CStr(lngQtyRecsFound) & " records found"
    Screen.MousePointer = vbDefault
    Me.cmdSend2.Enabled = True

End Sub

Private Sub cmdSend_Click()
Dim oSQL As New z_SQL
Dim oSM As New z_StockManager

    G.Update
    Me.cmdSend.Enabled = False
    oSQL.SendLoyaltyMissingSetToBranch Trim(Me.txtBranchcode)
    oSM.SendCustomerChanges
    
End Sub

Private Sub cmdSend2_Click()
Dim oSQL As New z_SQL
    
    G2.Update
    oSQL.SendInvocation "CustomerSet", txtBranchCode2, CreateXMLListOfAcnos2, "", ""
    
End Sub

Function CreateXMLListOfAcnos() As String
Dim i As Integer

    Set xMLDoc = New ujXML
    
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "AcnoSelection"
            .chCreate "MessageType"
                .elText = "AcnoSelection"

            For i = 1 To x.UpperBound(1)
                If x(i, 4) = "-1" Then
                    .elCreateSibling "DetailLine", True
                    .chCreate "ACNO"
                        .elText = x(i, 1)
                    .navUP
                End If
            Next i
    End With
    CreateXMLListOfAcnos = xMLDoc.docXML
End Function
Function CreateXMLListOfAcnos2() As String
Dim i As Integer

    Set xMLDoc = New ujXML
    
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "AcnoSelection"
            .chCreate "MessageType"
                .elText = "AcnoSelection"

            For i = 1 To X2.UpperBound(1)
                If X2(i, 5) = "-1" Then
                    .elCreateSibling "DetailLine", True
                    .chCreate "ACNO"
                        .elText = X2(i, 1)
                    .navUP
                End If
            Next i
    End With
    CreateXMLListOfAcnos2 = xMLDoc.docXML
End Function




Private Sub Form_Load()
    Me.Width = 7555
    Me.Height = 4185
    Top = 2000
    Left = 500
    Me.SSTab1.Tab = 0
    G.Visible = True
    G2.Visible = False
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
    G.Left = 200
    G2.Left = 200
    G.Width = Me.Width - (G.Left + 800)
    G2.Width = G.Width
    G.Height = Me.Height - (G.Top + 2120)
    G2.Height = G.Height
    SSTab1.Width = Me.Width - (SSTab1.Left + 400)
    Me.SSTab1.Height = Me.Height - 1900
    
    lngDiff = G.Height - lngDiff
    cmdClose.Top = SSTab1.Height + SSTab1.Top + 100
    cmdClose.Left = Me.Width - 1440
    lblRecs.Top = lblRecs.Top + lngDiff
    lblRecs2.Top = lblRecs2.Top + lngDiff
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        G.Visible = False
        G2.Visible = True
    Else
        G2.Visible = False
        G.Visible = True
    End If
End Sub

