VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{801C12A5-BE41-41CD-AE48-C666E77F2F02}#2.0#0"; "CCubeX20.ocx"
Object = "{B4B5B73C-172E-47B1-BFC2-C6F740957D01}#1.0#0"; "VB Control Manager.ocx"
Begin VB.Form frmCustomerRemittance 
   BackColor       =   &H00F9F2EE&
   Caption         =   "Remittances"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   17325
   Begin VBControlManager.ControlManager CM1 
      Height          =   9660
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   17325
      _ExtentX        =   30559
      _ExtentY        =   17039
      BackColor       =   9525832
      TitleBar_CloseVisible=   0   'False
      TitleBar_Height =   0
      TitleBar_Visible=   0   'False
      Begin VB.Frame fr4 
         BackColor       =   &H00F7EDE8&
         Height          =   4137
         Left            =   -30
         TabIndex        =   22
         Top             =   5523
         Width           =   11063
         Begin TrueOleDBGrid60.TDBGrid gCreditsAvailable 
            Height          =   2865
            Left            =   105
            OleObjectBlob   =   "frmCustPmt3.frx":0000
            TabIndex        =   6
            Top             =   195
            Width           =   7350
         End
         Begin VB.CommandButton cmdAddtoRemittance_CN 
            BackColor       =   &H00E7E6D8&
            Caption         =   "Add to remittance->"
            Height          =   345
            Left            =   8970
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Click to find all customers matching the retrictions entered."
            Top             =   3720
            UseMaskColor    =   -1  'True
            Width           =   1620
         End
      End
      Begin VB.Frame fr3 
         BackColor       =   &H00F7EDE8&
         Height          =   3887
         Left            =   0
         TabIndex        =   15
         Top             =   1576
         Width           =   11063
         Begin VB.CommandButton cmdAddtoRemittance_Pay 
            BackColor       =   &H00E7E6D8&
            Caption         =   "Add to remittance->"
            Height          =   345
            Left            =   9240
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Click to find all customers matching the retrictions entered."
            Top             =   2310
            UseMaskColor    =   -1  'True
            Width           =   1620
         End
         Begin TrueOleDBGrid60.TDBGrid gDeposits 
            Height          =   3195
            Left            =   120
            OleObjectBlob   =   "frmCustPmt3.frx":4442
            TabIndex        =   11
            Top             =   240
            Width           =   9930
         End
      End
      Begin VB.Frame fr2 
         BackColor       =   &H00F7EDE8&
         Height          =   9660
         Left            =   11123
         TabIndex        =   14
         Top             =   0
         Width           =   6202
         Begin CCubeX2.ContourCubeX CC 
            Height          =   7905
            Left            =   135
            TabIndex        =   16
            Top             =   165
            Width           =   3510
            Active          =   0   'False
            Transposed      =   0   'False
            NULLValueString =   ""
            Descending      =   0   'False
            NoTotals        =   0   'False
            NoGrandTotals   =   0   'False
            Caption         =   ""
            BackColor       =   16380654
            Enabled         =   -1  'True
            Alive           =   0   'False
            BorderStyle     =   1
            AllowInactiveDimArea=   -1  'True
            AllowExpand     =   -1  'True
            AllowPivot      =   -1  'True
            TotalsString    =   ""
            InactiveDimAreaBkColor=   16380654
            AutoSize        =   0   'False
            UnusedDataAreaColor=   16380654
            MousePointer    =   0
            Object.Visible         =   -1  'True
            InfoURL         =   "http://www.contourcomponents.com/contourcube_user_guide.htm"
            ConnectionString=   ""
            DataSourceType  =   0
            VERSION_NO      =   2
            CCubeXMetadata  =   $"frmCustPmt3.frx":9A9C
         End
         Begin VB.CommandButton cmdSaveLayout 
            BackColor       =   &H00E7E6D8&
            Caption         =   "&Save layout"
            Height          =   390
            Left            =   1140
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Click to find all customers matching the retrictions entered."
            Top             =   8160
            UseMaskColor    =   -1  'True
            Width           =   1650
         End
         Begin VB.CommandButton cmdClose 
            BackColor       =   &H00E7E6D8&
            Cancel          =   -1  'True
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
            Height          =   585
            Left            =   150
            Picture         =   "frmCustPmt3.frx":9F35
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   8115
            Width           =   960
         End
         Begin VB.CommandButton cmdPostBatch 
            BackColor       =   &H00E7E6D8&
            Caption         =   "&Save this remittance"
            Enabled         =   0   'False
            Height          =   390
            Left            =   3900
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Click to find all customers matching the retrictions entered."
            Top             =   8055
            UseMaskColor    =   -1  'True
            Width           =   1965
         End
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H00E7E6D8&
            Caption         =   "&Print this remittance"
            Height          =   390
            Left            =   2175
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Click to find all customers matching the retrictions entered."
            Top             =   8055
            UseMaskColor    =   -1  'True
            Width           =   1650
         End
      End
      Begin VB.Frame Fr1 
         BackColor       =   &H00F7EDE8&
         Height          =   1516
         Left            =   15
         TabIndex        =   13
         Top             =   0
         Width           =   11063
         Begin VB.TextBox txtCustRemittanceCode 
            Height          =   300
            Left            =   135
            MaxLength       =   30
            TabIndex        =   0
            Top             =   360
            Width           =   1950
         End
         Begin VB.TextBox txtArg 
            Height          =   330
            Left            =   4290
            MaxLength       =   30
            TabIndex        =   4
            Top             =   945
            Width           =   1035
         End
         Begin VB.TextBox txtBatchTotal 
            Height          =   285
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   1
            Top             =   360
            Width           =   1439
         End
         Begin VB.TextBox txtDepositDate 
            Height          =   285
            Left            =   4170
            MaxLength       =   30
            TabIndex        =   2
            Top             =   360
            Width           =   1320
         End
         Begin VB.ComboBox cboCC 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   945
            Width           =   4200
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Remittance code"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   135
            TabIndex        =   23
            Top             =   120
            Width           =   1545
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoiced customer"
            ForeColor       =   &H8000000D&
            Height          =   345
            Left            =   105
            TabIndex        =   21
            Top             =   735
            Width           =   2325
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "A/c ref"
            ForeColor       =   &H8000000D&
            Height          =   225
            Left            =   4365
            TabIndex        =   20
            Top             =   705
            Width           =   840
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Total deposited"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   2625
            TabIndex        =   19
            Top             =   165
            Width           =   1545
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Deposit date"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   4515
            TabIndex        =   18
            Top             =   165
            Width           =   1065
         End
         Begin VB.Label lbl1 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer remittance number"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   -2265
            TabIndex        =   17
            Top             =   660
            Width           =   2130
         End
      End
   End
End
Attribute VB_Name = "frmCustomerRemittance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTPID As Long
Dim lngChildCustID As Long
Dim lngAmt As Long
Dim dteDate As Date
Dim strReason As String
Dim bAmt As Boolean
Dim bDate As Boolean
Dim bReason As Boolean
Dim strCustomerName As String
Dim strInvoices As String
Dim lngInvoiceID As Long
Dim XA As New XArrayDB
Dim XCN As New XArrayDB
Dim x As New XArrayDB
Dim rs As New ADODB.Recordset
Dim rsCN As New ADODB.Recordset
Dim tlChildCustomers As z_TextList
Dim flgLoading As Boolean
Dim bDirty As Boolean
Dim hwndDebtorsForm As Long

Dim mStatementLineID As Long

Public Sub Component(pTPID As Long, _
                        pCustomerName As String, _
                        DebtorsFormHandle As Long, _
                        DepositDate As Date, _
                        Reference As String, _
                        StatementLineID As Long)
    On Error GoTo errHandler
Dim i As Integer
    mStatementLineID = StatementLineID
    hwndDebtorsForm = DebtorsFormHandle
    lngTPID = pTPID
    strCustomerName = pCustomerName
    Caption = "Payments from " & pCustomerName
    txtDepositDate = Format(DepositDate, "dd/mm/yyyy")
    txtCustRemittanceCode = Reference
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerRemittance.Component(pTPID,pCustomerName,DebtorsFormHandle,DepositDate," & _
        "Reference,StatementLineID)", Array(pTPID, pCustomerName, DebtorsFormHandle, DepositDate, Reference, _
         StatementLineID)
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
Dim Row As Long, Col As Integer

    flgLoading = True
    bAmt = False
    bDate = False
    Set tlChildCustomers = New z_TextList
    tlChildCustomers.Load ltChildCustomers, CStr(lngTPID)
    LoadComboImmCust
    XA.ReDim 0, 0, 0, 10
    Set gDeposits.Array = XA
    rs.CursorLocation = adUseClient
    rs.Fields.Append "TPID", adInteger
    rs.Fields.Append "TargetTRID", adInteger
    rs.Fields.Append "CustomerName", adVarChar, 100
    rs.Fields.Append "Reference", adVarChar, 100
    rs.Fields.Append "TargetReference", adVarChar, 100
    rs.Fields.Append "Dte", adDate
    rs.Fields.Append "Amount", adDouble
    rs.Fields.Append "SettlementDiscount", adDouble
    rs.Fields.Append "Balance", adDouble
    rs.Fields.Append "DocType", adVarChar, 10
    rs.Open
   
    'LoadFormat
    SetGridLayout gCreditsAvailable, Me.Name & gCreditsAvailable.Name
    SetGridLayout gDeposits, Me.Name & gDeposits.Name
    SetFormSize Me
    SetCM Me, CM1
    
    bDirty = False
    flgLoading = False
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadFormat()
Dim oFSo As New FileSystemObject
'  CommonDialog1.DefaultExt = "cuf"
'  CommonDialog1.DialogTitle = "Load Cube layout"
'  CommonDialog1.InitDir = oPC.SharedFolderRoot & "\CubeFormats"
'  CommonDialog1.CancelError = True
'  On Error Resume Next
'  CommonDialog1.ShowOpen
'  If Err.Number = cdlCancel Then
'    On Error GoTo 0
'    Exit Sub
'  Else
'    On Error GoTo 0
'    LoadContourcubeLayout CommonDialog1.FileName
'  End If
    If oFSo.FileExists(CStr(oPC.LocalFolder & "\Templates\AccountsCC_1.txt")) Then
        LoadContourcubeLayout CStr(oPC.LocalFolder & "\Templates\AccountsCC_1.txt"), Me.CC
    End If

End Sub

Private Sub Preparecube()
    On Error GoTo errHandler
Dim oTLS As New z_TextListSimple
Dim Fact As IViewFact
    
    If rs Is Nothing Then Exit Sub
    If rs.EOF = True And rs.BOF = True Then Exit Sub
    rs.MoveFirst
    If rs.EOF Then
        MsgBox "No records", , "Status"
    End If
    
    If Not rs.EOF Then
        
        CloseCube
        With CC.Cube
            .Dims.Add("CustomerName", "CustomerName", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("TargetReference", "TargetReference", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("Reference", "Reference", , xda_vertical).MoveTo xda_vertical
            .BaseFacts.Add "Amount", "Amount"
            .Facts.Add "Amount", "Amount", xfaa_SUM
            .BaseFacts.Add "SettlementDiscount", "SettlementDiscount"
            .Facts.Add "SettlementDiscount", "SettlementDiscount", xfaa_SUM
            .BaseFacts.Add "Balance", "Balance"
            .Facts.Add "Balance", "Balance", xfaa_SUM
            CC.Facts(0).Appearance.Format = "###,##0.00;(###,##0.00)"
            CC.Facts(0).Caption = "Amount"
            CC.Facts(1).Appearance.Format = "###,##0.00;(###,##0.00)"
            CC.Facts(1).Caption = "Sett.disc."
            CC.Facts(1).Appearance.Format = "###,##0.00;(###,##0.00)"
            CC.Facts(1).Caption = "Balance"
            CC.NoGrandTotals = False
           ' CC.Dims(0).NoTotals = True
           ' CC.Dims(1).NoTotals = True
            CC.TitleSettings.Text = "Payments summary"
            CC.VAxis.DrillDownLevel = 0
            For Each Fact In CC.Facts
              Fact.Visible = True
            Next
            Set rs.ActiveConnection = Nothing
            .Open rs

        End With
        AfterOpen
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformance.Preparecube"
    HandleError
End Sub


Public Sub LoadRemittance(CashbookLineID As Long, TRID As Long)
    If CashbookLineID > 0 Then
    End If

End Sub

Private Sub gCreditsAvailable_ButtonClick(ByVal ColIndex As Integer)

    If ColIndex <> 4 Then Exit Sub
    Select Case UCase(XCN(gCreditsAvailable.Bookmark, 3))
    Case ""
        XCN(gCreditsAvailable.Bookmark, 3) = "Claimed"
        XCN(gCreditsAvailable.Bookmark, 4) = "Un-claim"
    Case "CLAIMED"
        XCN(gCreditsAvailable.Bookmark, 3) = ""
        XCN(gCreditsAvailable.Bookmark, 4) = "CLAIM"
    End Select
    gCreditsAvailable.Refresh

End Sub
Private Sub cmdAddtoRemittance_CN_Click()
Dim ret As Long
' '   keybd_event VK_LSHIFT, 0, 0, 0
'    ret = SendMessage(hwndDebtorsForm, WM_RBUTTONDOWN, 0&, 0&)
'  '  keybd_event VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0
Dim rsClone As ADODB.Recordset
Dim i As Integer

    Me.gCreditsAvailable.Update
    If Not (rs.EOF And rs.BOF) Then rs.MoveFirst
    'Clear destination recordset
   ' rs.MoveFirst
    Do While Not rs.EOF
        If rs.Fields("Customer") = cboCC And rs.Fields("DocType") = "CN" Then
            rs.Delete
            rs.Update
        End If
        rs.MoveNext
    Loop
    'Load destination recordset
    i = 0
    If rs.RecordCount > 0 Then rs.MoveFirst
    Do While i <= XCN.UpperBound(1)
        If UCase(XCN(i, 3)) <> "CLAIMED" Then Exit Do
            rs.AddNew
                rs.Fields("TPID") = Val(tlChildCustomers.key(cboCC))
                rs.Fields("Customer") = cboCC
                rs.Fields("Reference") = XCN(i, 1)
                rs.Fields("TargetReference") = ""
                rs.Fields("Dte") = XCN(i, 0)
                rs.Fields("Amount") = FNDBL(XCN(i, 2)) * -1
                rs.Fields("SettlementDiscount") = 0
                rs.Fields("DocType") = "CN"
                rs.Fields("TargetTRID") = XCN(i, 6)
            rs.Update
     '   End If
        i = i + 1
    Loop

    Preparecube
    bDirty = True
    
    cmdPostBatch.Enabled = TotalPosted = CDbl(txtBatchTotal)

End Sub


Private Sub gDeposits_ButtonClick(ByVal ColIndex As Integer)
Dim frm As frmSelectInvoice
Dim Xcoord As Long
Dim Ycoord As Long
    gDeposits.Update
    If tlChildCustomers.key(cboCC) > 0 Then
        If XA(gDeposits.Bookmark, 1) > "" Then
            PointsToMe Me.hwnd, Xcoord, Ycoord
            
            Set frm = New frmSelectInvoice
            frm.Component tlChildCustomers.key(cboCC), Xcoord, Ycoord
            frm.Show vbModal
            If frm.SelectedDebitID > 0 Then
                XA(gDeposits.Bookmark, 4) = frm.SelectedDebitCode
                XA(gDeposits.Bookmark, 2) = frm.SelectedDebitAmount
                XA(gDeposits.Bookmark, 3) = frm.SelectedDebitOS
                XA(gDeposits.Bookmark, 9) = frm.SelectedDebitID
                XA(gDeposits.Bookmark, 10) = frm.SelectedDebitType
                gDeposits.Refresh
                DoEvents
            End If
            Unload frm
        End If
        XA(gDeposits.Bookmark, 7) = FNDBL(gDeposits.Columns(3)) - FNDBL(gDeposits.Columns(5)) - FNDBL(gDeposits.Columns(6))
        gDeposits.RefetchRow
        
    End If
End Sub

Private Sub gDeposits_SelChange(Cancel As Integer)

    gDeposits.SetFocus

End Sub

Private Sub gDeposits_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
'    If FNDBL(XA(Bookmark, 5)) < 0 Then
'        RowStyle.ForeColor = vbRed
'    End If
End Sub

Private Sub txtArg_KeyPress(KeyAscii As Integer)
Dim bFound As Boolean

    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Trim(txtArg) = "" Or Trim(txtArg) = "*" Then Exit Sub
    
    If KeyAscii = 13 Then  ' The ENTER key.
        bFound = FindCustomer(txtArg, lngTPID, strCustomerName)
        If bFound Then
                Me.cboCC.Enabled = True
                cboCC.Text = strCustomerName
                LoadDepositsGrid
        End If
    End If
    Exit Sub
errHandler:
    ErrPreserve
    If Err = 383 Then
        MsgBox "This customer record is not associated with the parent record: " & strCustomerName
        Err.Clear
        Exit Sub
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub
Private Sub FetchCreditNotes()
    oPC.OpenDBSHort
    Set rsCN = New ADODB.Recordset
    rsCN.CursorLocation = adUseClient
    rsCN.Open "SELECT TPID,DocDate,PayableAmount,DocCode,Balance,DocID FROM vMatchesCNOS WHERE TPID = " & CStr(lngChildCustID) & " ORDER BY DocDate", oPC.COShort, adOpenStatic
End Sub
Private Sub LoadCreditNotesGrid()
    On Error GoTo errHandler
Dim i As Long
    i = 0
    XCN.Clear
  '  oPC.OpenDBSHort
  '  Set rsCN = New ADODB.Recordset
  '  rsCN.CursorLocation = adUseClient
  '  rsCN.Open "SELECT TPID,Dte,DateF,ValueF,DocCode,dblValue,TR_ID FROM vDocsUnpostedPerTP WHERE TPID = " & CStr(lngTPID) & " ORDER BY Dte", oPC.COShort, adOpenStatic
    Do While Not rsCN.EOF
        XCN.ReDim 0, i, 0, 15
        XCN(i, 0) = FNS(rsCN.Fields("DocDate"))
        XCN(i, 1) = FNS(rsCN.Fields("DocCode"))
        XCN(i, 2) = FNS(rsCN.Fields("PayableAmount"))
        XCN(i, 3) = ""
        XCN(i, 5) = FNDBL(rsCN.Fields("Balance"))
        XCN(i, 6) = FNN(rsCN.Fields("DocID"))
        i = i + 1
        rsCN.MoveNext
    Loop
    gCreditsAvailable.Array = XCN
    gCreditsAvailable.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerRemittance.LoadCreditNotesGrid"
End Sub
Private Function FindCustomer(pRef As String, pTPID As Long, pName As String) As Boolean
Dim oSQL As New z_SQL

    oSQL.FindCustomerByOurRefNo pRef, pTPID, pName
    FindCustomer = (pTPID > 0)
    
End Function
Private Sub cboCC_Change()
    If cboCC.DataChanged = True Then
        LoadDepositsGrid
        FetchCreditNotes
        LoadCreditNotesGrid
    End If
End Sub

Private Sub cboCC_Click()
    If flgLoading Then Exit Sub
    Me.txtArg = tlChildCustomers.f4(tlChildCustomers.key(cboCC))
    lngChildCustID = tlChildCustomers.key(cboCC)
    LoadDepositsGrid
    FetchCreditNotes
    LoadCreditNotesGrid
End Sub

Private Sub cboCC_LostFocus()
'LoadDepositsGrid
End Sub

Private Sub cmdClose_Click()
Dim bInProcess As Boolean

    Unload Me
End Sub
Private Sub LoadDepositsGrid()
Dim i As Integer
 Dim vi As ValueItem

    XA.Clear
    If Not (rs.EOF And rs.BOF) Then rs.MoveFirst
    i = 0
    Do While Not rs.EOF
        If rs.Fields("Customer") = cboCC Then
            XA(i, 0) = rs.Fields("Dte").Value
            XA(i, 1) = FNS(rs.Fields("Reference").Value)
            XA(i, 2) = CStr(rs.Fields("Amount").Value)
            XA(i, 3) = CStr(rs.Fields("SettlementDiscount").Value)
            XA(i, 4) = ""
            XA(i, 5) = ""
            XA(i, 6) = FNN(rs.Fields("TargetTRID").Value)
            i = i + 1
        End If
        rs.MoveNext
    Loop
    Me.gDeposits.Refresh
    Me.gDeposits.MoveFirst
End Sub

Private Sub cmdPrint_Click()
    CC.PrintCube True
End Sub
Private Sub cmdSaveLayout_Click()
'Dim fs As New FileSystemObject
'    If Not fs.FolderExists(oPC.SharedFolderRoot & "\CubeFormats") Then
'        fs.CreateFolder (oPC.SharedFolderRoot & "\CubeFormats")
'    End If
'  CommonDialog1.DefaultExt = "cuf"
'  CommonDialog1.DialogTitle = "Save Cube layout"
'  CommonDialog1.InitDir = oPC.SharedFolderRoot & "\CubeFormats"
'  CommonDialog1.CancelError = True
'  On Error Resume Next
'  CommonDialog1.ShowSave
'  If Err.Number = cdlCancel Then
'    On Error GoTo 0
'    Exit Sub
'  Else
'    On Error GoTo 0
'    If Trim(CommonDialog1.FileName) <> "" Then SaveContourCubeLayout CStr(oPC.localfolder & "\Templates\AccountsCC_1.txt")
'  End If


    'If Trim(CommonDialog1.FileName) <> "" Then
    SaveContourCubeLayout CStr(oPC.LocalFolder & "\Templates\AccountsCC_1.txt"), Me.CC
    
End Sub
    


   
    
    
Private Sub cmdPostBatch_Click()
    On Error GoTo errHandler
Dim xMLDoc As ujXML
Dim XMLArgs As String
Dim Strguid As String
Dim i As Integer
Dim oSM As New z_StockManager
Dim lngPaid As Long

    If mStatementLineID > 0 Then
        If MsgBox("There is already a payment or remittance associated with this statement line." & vbCrLf _
            & "Continuing with this remittance will replace it. Continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
                Exit Sub
            
        End If
    End If
    
  '  If oPC.Configuration.SignTransactions = True Then
     '  If SecurityControl(enSECURITY_ACCEPTACPAYMENT, , "Save payment batch", DOCAPPROVAL) = False Then
  '             Exit Sub
     '   End If
  '  End If
    
    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "doc_PaymentBatch"
            .chCreate "MessageType"
                .elText = "PAYMENT_ACTION"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "StaffID"
                .elText = CStr(gSTAFFID)
            .elCreateSibling "RemittanceReference"
                .elText = Trim(txtCustRemittanceCode)
            .elCreateSibling "DepositDate", True
                .elText = Format(txtDepositDate, "YYYYMMDD")
            .elCreateSibling "RemittanceFromTradingPartnerID", True
                .elText = CStr(lngTPID)
            .elCreateSibling "CashbookLineID", True
                .elText = CStr(mStatementLineID)
            .elCreateSibling "DetailLines", True
            i = 0
            rs.MoveFirst
            Do While Not rs.EOF
                .chCreate "I"
                .chCreate "TradingPartnerID"
                    .elText = rs.Fields("TPID").Value
                .elCreateSibling "Amount", True
                    .elText = rs.Fields("Amount").Value
                .elCreateSibling "SettlementDiscount", True
                    .elText = rs.Fields("SettlementDiscount").Value
                .elCreateSibling "TRID", True
                    .elText = CStr(rs.Fields("TargetTRID").Value)
                .navUP
                .navUP
                i = i + 1
                rs.MoveNext
            Loop

         XMLArgs = .docXML
  
    End With
    oSM.InsertScript Strguid, XMLArgs

    If Strguid > "" Then
        oSM.Action_InsertPayments Strguid, lngPaid
    End If
    bDirty = False
    MsgBox "Payments done", vbInformation + vbOKOnly, "Status"
    
    
   
    Unload Me
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.cmdPostBatch_Click", , EA_NORERAISE
    HandleError
End Sub


    
Private Sub cmdAddtoRemittance_Pay_Click()
    On Error GoTo errHandler
Dim rsClone As ADODB.Recordset
Dim i As Integer

    ValidateDeposits
    Me.gDeposits.Update
    If Not (rs.EOF And rs.BOF) Then rs.MoveFirst
    'Clear destination recordset
   ' rs.MoveFirst
    Do While Not rs.EOF
        If rs.Fields("Customer") = cboCC And rs.Fields("DocType") = "PAY" Then
            rs.Delete
            rs.Update
        End If
        rs.MoveNext
    Loop
    'Load destination recordset
    i = 0
    If rs.RecordCount > 0 Then rs.MoveFirst
    Do While i <= XA.UpperBound(1)
        If XA(i, 0) = "" Then Exit Do
            rs.AddNew
                rs.Fields("TPID") = Val(tlChildCustomers.key(cboCC))
                rs.Fields("Customer") = cboCC
                rs.Fields("Reference") = XA(i, 1)
                rs.Fields("TargetReference") = XA(i, 4)
                rs.Fields("Dte") = XA(i, 0)
                rs.Fields("Amount") = FNDBL(XA(i, 5))
                rs.Fields("SettlementDiscount") = FNDBL(XA(i, 6))
                rs.Fields("Balance") = FNDBL(XA(i, 7))
                rs.Fields("TargetTRID") = XA(i, 6)
                rs.Fields("DocType") = "PAY"
            rs.Update
        i = i + 1
    Loop

    Preparecube
     LoadFormat

    bDirty = True
    
    cmdPostBatch.Enabled = TotalPosted = CDbl(txtBatchTotal)
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.cmdTransfer_Click", , EA_NORERAISE
    HandleError
End Sub
Private Function TotalPosted() As Double
    Dim dblTotalPayment As Double
    If rs.RecordCount = 0 Then
        TotalPosted = 0
        Exit Function
    End If
    
    rs.MoveFirst
    dblTotalPayment = 0
    Do While Not rs.EOF
        dblTotalPayment = dblTotalPayment + CDbl(FNDBL(rs.Fields("Amount"))) + CDbl(FNDBL(rs.Fields("SettlementDiscount")))
        rs.MoveNext
    Loop
    TotalPosted = dblTotalPayment
End Function
Private Sub LoadComboImmCust()
Dim oSQL As New z_SQL
Dim rs As ADODB.Recordset
Dim Res As Long

    LoadCombo cboCC, tlChildCustomers
    
End Sub
Private Sub CloseCube()
    On Error GoTo errHandler
 With CC
   .Active = False
   .Cube.Dims.Clear
   .Cube.Facts.Clear
   .Cube.BaseFacts.Clear
 End With
' CheckEnabled
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.CloseCube"
End Sub
Private Sub AfterOpen()
    On Error GoTo errHandler
 CC.Visible = CC.Active
' CheckEnabled
 CheckVisible
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.AfterOpen"
End Sub
Private Sub CheckVisible()
    On Error GoTo errHandler
 CC.Visible = True 'ContourCubeX.Active
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.CheckVisible"
End Sub
Private Sub CM1_SplitterMoveEnd(ByVal IdSplitter As Long, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    resize
End Sub

Private Sub Form_Resize()
    resize
End Sub
Private Sub resize()
Dim lngDiff As Long
    If flgLoading Then Exit Sub
    
    gDeposits.Left = fr3.Left + 50
    gDeposits.Top = NonNegative_Lng(fr3.Top - fr1.Height + 120)
    gDeposits.Width = NonNegative_Lng(fr3.Width - 300)
    gDeposits.Height = NonNegative_Lng(fr3.Height - 700)
    cmdAddtoRemittance_Pay.Top = NonNegative_Lng(fr3.Height - 400)
    cmdAddtoRemittance_Pay.Left = NonNegative_Lng(fr3.Width - 2000)
    
    gCreditsAvailable.Left = fr4.Left + 50
    gCreditsAvailable.Top = NonNegative_Lng(fr4.Top - fr3.Height - fr1.Height + 120)
    gCreditsAvailable.Width = NonNegative_Lng(fr4.Width - 300)
    gCreditsAvailable.Height = NonNegative_Lng(fr4.Height - 1400)
    cmdAddtoRemittance_CN.Top = NonNegative_Lng(fr4.Height - 400)
    cmdAddtoRemittance_CN.Left = NonNegative_Lng(fr4.Width - 2000)
    
  '  CC.Left = Fr1.Width + fr2.Left + 50
  '  CC.Top = fr2.Top + 50
    CC.Width = NonNegative_Lng(fr2.Width - 600)
    CC.Height = NonNegative_Lng(fr2.Height - 1100)
'
    cmdPostBatch.Top = NonNegative_Lng(fr2.Height - 700)
    cmdPostBatch.Left = NonNegative_Lng(fr2.Width - 2500)
    cmdClose.Top = cmdPostBatch.Top
    cmdPrint.Top = cmdPostBatch.Top
    cmdSaveLayout.Top = cmdPostBatch.Top
'
End Sub
Private Sub Form_Unload(Cancel As Integer)

    If bDirty Then
        If MsgBox("It looks like you have unsaved work. Confirm you want to close form.", vbInformation + vbYesNo, "Warning") = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If

    SaveLayout gCreditsAvailable, Me.Name & gCreditsAvailable.Name
    SaveLayout gDeposits, Me.Name & gDeposits.Name
    SaveFormSize Me.Name, Me.Height, Me.Width
    SaveSplits Me.Name, Me.CM1
   
End Sub

Private Sub gDeposits_AfterUpdate()
    bDirty = True
End Sub
Private Sub gDeposits_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 4 Or ColIndex = 5 Or ColIndex = 6 Then
    XA(gDeposits.Bookmark, 7) = FNDBL(gDeposits.Columns(3)) - FNDBL(gDeposits.Columns(5)) - FNDBL(gDeposits.Columns(6))
    gDeposits.RefetchRow
    End If
End Sub

Private Sub txtBatchTotal_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtBatchTotal) And txtBatchTotal > ""
    
    cboCC.Enabled = Not Cancel
    gDeposits.Enabled = Not Cancel
    
    If Not Cancel Then
        cmdPostBatch.Enabled = TotalPosted = FNDBL(txtBatchTotal)
        cmdAddtoRemittance_Pay.Enabled = True
    Else
        cmdAddtoRemittance_Pay.Enabled = False
    End If
    
End Sub

Private Sub txtCustRemittanceCode_Validate(Cancel As Boolean)
    Cancel = Len(txtCustRemittanceCode) < 3 And txtCustRemittanceCode > ""
End Sub

Private Sub txtDepositDate_Validate(Cancel As Boolean)
    Cancel = Not IsDate(txtDepositDate)
End Sub

Private Sub ValidateDeposits()
MsgBox "Validating Deposits"
Dim i As Integer

    For i = 0 To XA.UpperBound(1)
        If XA(i, 10) = "IN" Then
            'If type = IN and SD > 0 and new balance OS > 0 then problem: no fully paid:can't take SD
            'If type = IN and new balance < 0 then problem: overpayment
        End If
    Next

End Sub
