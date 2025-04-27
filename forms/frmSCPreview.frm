VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmSCPreview 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Supplier claim details"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   16080
   Begin TabDlg.SSTab SSTab 
      Height          =   6210
      Left            =   30
      TabIndex        =   0
      Top             =   375
      Width           =   15945
      _ExtentX        =   28125
      _ExtentY        =   10954
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   11558991
      TabCaption(0)   =   "Claim details"
      TabPicture(0)   =   "frmSCPreview.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTotalClaim"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTotalAcceptedClaimValue"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ClaimsGrid"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdIssue"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdclose"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtTotalClaim"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtTotalAcceptedClaimValue"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Document"
      TabPicture(1)   =   "frmSCPreview.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdToPDF"
      Tab(1).Control(1)=   "arv"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdToPDF 
         BackColor       =   &H00D5D5C1&
         Caption         =   "PDF"
         Height          =   360
         Left            =   -74790
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   510
         Width           =   1380
      End
      Begin VB.TextBox txtTotalAcceptedClaimValue 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   6015
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   3600
         Width           =   1560
      End
      Begin VB.TextBox txtTotalClaim 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1365
         TabIndex        =   5
         Text            =   "txtTotalClaim"
         Top             =   3630
         Width           =   1560
      End
      Begin VB.CommandButton cmdclose 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1365
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSCPreview.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Close the purchase order"
         Top             =   4185
         Width           =   885
      End
      Begin VB.CommandButton cmdIssue 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Issu&e claim"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   7215
         Picture         =   "frmSCPreview.frx":03C2
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   4245
         UseMaskColor    =   -1  'True
         Width           =   1620
      End
      Begin TrueOleDBGrid60.TDBGrid ClaimsGrid 
         Height          =   3240
         Left            =   0
         OleObjectBlob   =   "frmSCPreview.frx":074C
         TabIndex        =   1
         Top             =   360
         Width           =   15855
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 arv 
         Height          =   5235
         Left            =   -74820
         TabIndex        =   4
         Top             =   915
         Width           =   13875
         _ExtentX        =   24474
         _ExtentY        =   9234
         SectionData     =   "frmSCPreview.frx":894F
      End
      Begin VB.Label lblTotalAcceptedClaimValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Total accepted claim value"
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   4020
         TabIndex        =   9
         Top             =   3600
         Width           =   1965
      End
      Begin VB.Label lblTotalClaim 
         BackStyle       =   0  'Transparent
         Caption         =   "Total claim value"
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   45
         TabIndex        =   6
         Top             =   3630
         Width           =   1965
      End
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00915A48&
      Height          =   330
      Left            =   180
      TabIndex        =   7
      Top             =   15
      Width           =   2445
   End
End
Attribute VB_Name = "frmSCPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oSQL As z_SQL
Dim rsSCDetails As ADODB.Recordset
Dim XA As New XArrayDB
Dim lngTRID As Long
Dim lngTPID As Long
Dim ar As arSupplierClaim
Dim mSupplierName As String
Dim sTotalClaimValue As String
Dim mAcceptedClaimValue As Double
Dim mApprovalRequired As Boolean

Public Sub component(ID As Long, StatusF As String, TotalClaimValueF As String, pSupplierName As String, ApprovalRequired As Boolean, pTPID As Long)
    lngTRID = ID
    mSupplierName = pSupplierName
    sTotalClaimValue = TotalClaimValueF
    Me.lblStatus.Caption = StatusF
    mApprovalRequired = ApprovalRequired
    lngTPID = pTPID
    cmdToPDF.Enabled = (StatusF = "Closed")
    Me.ClaimsGrid.Columns(14).Visible = mApprovalRequired
    Me.ClaimsGrid.Columns(15).Visible = mApprovalRequired
    If StatusF = "Closed" Then
        Me.cmdIssue.Enabled = False
    End If
        Me.txtTotalAcceptedClaimValue.Visible = mApprovalRequired
        Me.lblTotalAcceptedClaimValue.Visible = mApprovalRequired
End Sub
Private Sub resize()
    SSTab.Width = NonNegative_Lng(Me.Width - 400)
    SSTab.Height = NonNegative_Lng(Me.Height - 1000)
    If SSTab.Tab = 0 Then
        Me.ClaimsGrid.Left = 50
        Me.ClaimsGrid.TOP = 470
        Me.ClaimsGrid.Width = NonNegative_Lng(SSTab.Width - 300)
        Me.ClaimsGrid.Height = NonNegative_Lng(SSTab.Height - 1800)
        
       ' cmdPrint.top = NonNegative_Lng(ClaimsGrid.Height + 650)
        cmdClose.TOP = NonNegative_Lng(ClaimsGrid.Height + 550)
        cmdClose.Left = NonNegative_Lng(ClaimsGrid.Width - 800)
        
        cmdIssue.TOP = cmdClose.TOP
        cmdIssue.Left = 8000
        cmdIssue.Height = cmdClose.Height
        
        txtTotalClaim.TOP = cmdIssue.TOP
        Me.lblTotalClaim.TOP = txtTotalClaim.TOP
        Me.txtTotalAcceptedClaimValue.TOP = txtTotalClaim.TOP
        Me.lblTotalAcceptedClaimValue.TOP = txtTotalClaim.TOP
    Else
        arv.Width = NonNegative_Lng(SSTab.Width - 400)
        arv.Height = NonNegative_Lng(SSTab.Height - 1300)
    End If
End Sub

Private Sub ClaimsGrid_ButtonClick(ByVal ColIndex As Integer)
    If ColIndex <> 15 Then Exit Sub
    Select Case UCase(XA(ClaimsGrid.Bookmark, 16))
    Case "REJECT"
        XA(ClaimsGrid.Bookmark, 15) = "Rejected"
        XA(ClaimsGrid.Bookmark, 16) = "Accept"
    Case "ACCEPT"
        XA(ClaimsGrid.Bookmark, 15) = "Accepted"
        XA(ClaimsGrid.Bookmark, 16) = "Reject"
    End Select
    ClaimsGrid.Refresh
    mAcceptedClaimValue = Recalculate()
    Me.txtTotalAcceptedClaimValue = Format(mAcceptedClaimValue, "###,##0.00")
    
  '  RefreshTotal

End Sub



Private Sub cmdClose_Click()
Unload Me
End Sub

Private Function Recalculate() As Double
Dim i As Integer
Dim dblTot As Double
    dblTot = 0
    For i = 1 To XA.UpperBound(1)
        If UCase(XA.Value(i, 15)) = "ACCEPTED" Then
            dblTot = dblTot + CDbl(XA.Value(i, 20))
        End If
    Next i
    Recalculate = dblTot
    
End Function
Private Sub cmdIssue_Click()
Dim msg As String

    If mApprovalRequired Then
        If CheckAllResponsesIn = True Then
            Issue msg
        Else
            MsgBox "You cannot issue this claim as not all lines have been responded to. Use the Action button to mark the vendor responses to all lines.", vbOKOnly + vbInformation, "Can't do this"
            Exit Sub
        End If
    Else
        Issue msg
    End If
    If msg = "" Then
        cmdIssue.Enabled = False
        Me.cmdToPDF.Enabled = True
    Else
        MsgBox msg, vbInformation + vbOKOnly, "Can't do this"
    End If
End Sub
Private Function CheckAllResponsesIn() As Boolean
Dim bOK As Boolean
Dim i As Integer

    bOK = True
    For i = 1 To XA.UpperBound(1)
        If Not (UCase(XA.Value(i, 15)) = "ACCEPTED" Or UCase(XA.Value(i, 15)) = "REJECTED") Then
            bOK = False
        End If
    Next i
    CheckAllResponsesIn = bOK
End Function
Private Sub Issue(msg As String)
Dim oSQL As New z_SQL
    msg = oSQL.IssueSuppliersClaim(lngTRID, lngTPID)
End Sub
Private Sub Form_Load()
    SetGridLayout Me.ClaimsGrid, Me.Name & ClaimsGrid.Name
    SetFormSize Me
    txtTotalClaim = sTotalClaimValue
    Set oSQL = New z_SQL
    
    Set rsSCDetails = New ADODB.Recordset
    oSQL.LoadSCDetails lngTRID, rsSCDetails
    If Not rsSCDetails.eof Then
        LoadGrid
        Set ar = Nothing
        Set ar = New arSupplierClaim
        ar.Visible = False
        
        arv.ReportSource = ar
        arv.Zoom = 75
        rsSCDetails.MoveFirst
        ar.component rsSCDetails, mSupplierName, FNS(rsSCDetails.Fields("SCCode")), sTotalClaimValue
    Else
    
    End If

 
End Sub

Private Sub Form_Resize()
    resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveLayout Me.ClaimsGrid, Me.Name & ClaimsGrid.Name
    SaveFormSize Me.Name, Me.Height, Me.Width
End Sub

Private Sub LoadGrid()
Dim lngIndex As Long

    XA.Clear
    XA.ReDim 1, rsSCDetails.RecordCount, 1, 23
    For lngIndex = 1 To rsSCDetails.RecordCount
        XA.Value(lngIndex, 1) = FNS(rsSCDetails.Fields("SupplierInvoiceCode"))
        XA.Value(lngIndex, 2) = FND(rsSCDetails.Fields("SupplierInvoiceDate"))
        XA.Value(lngIndex, 3) = FNS(rsSCDetails.Fields("GRNCode"))
        XA.Value(lngIndex, 4) = FND(rsSCDetails.Fields("GRNDate"))
        XA.Value(lngIndex, 5) = FNS(rsSCDetails.Fields("EAN"))
        XA.Value(lngIndex, 6) = FNS(rsSCDetails.Fields("Description"))
        XA.Value(lngIndex, 7) = FNS(rsSCDetails.Fields("PriceF"))
        XA.Value(lngIndex, 8) = FNS(rsSCDetails.Fields("Discount"))
        XA.Value(lngIndex, 9) = FNDBL(rsSCDetails.Fields("QtyInv"))
        XA.Value(lngIndex, 10) = FNDBL(rsSCDetails.Fields("QtyShort"))
        XA.Value(lngIndex, 11) = FNDBL(rsSCDetails.Fields("CorrectedDiscount"))
        XA.Value(lngIndex, 12) = Format(FNS(rsSCDetails.Fields("CorrectedPriceF")), "###,###.00")
        XA.Value(lngIndex, 13) = FNS(rsSCDetails.Fields("Explanations"))
        XA.Value(lngIndex, 14) = Format(FNS(rsSCDetails.Fields("ClaimValue")), "###,###.00")
        XA.Value(lngIndex, 15) = ""
        XA.Value(lngIndex, 16) = "Accept"
        XA.Value(lngIndex, 19) = FNS(rsSCDetails.Fields("SCStatus"))
        XA.Value(lngIndex, 20) = FNS(rsSCDetails.Fields("TotalClaimValueF"))
        rsSCDetails.MoveNext
    Next
    XA.QuickSort 1, rsSCDetails.RecordCount, 1, XORDER_ASCEND, XTYPE_STRING, 2, XORDER_ASCEND, XTYPE_STRING, 5, XORDER_ASCEND, XTYPE_STRING
    ClaimsGrid.Array = XA
    ClaimsGrid.ReBind

End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
    resize
End Sub

Private Sub cmdToPDF_Click()
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    ar.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "TEMP\" & "Claim_" & mSupplierName & "_" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(ar.Pages)
    OpenFileWithApplication fn, enPDF
End Sub

