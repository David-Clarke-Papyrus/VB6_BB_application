VERSION 5.00
Object = "{801C12A5-BE41-41CD-AE48-C666E77F2F02}#2.0#0"; "CCubeX20.ocx"
Begin VB.Form frmCRemittancePreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Debtor's transactions"
   ClientHeight    =   8280.001
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8280.001
   ScaleWidth      =   7230
   Begin CCubeX2.ContourCubeX CC 
      Height          =   6735
      Left            =   135
      TabIndex        =   1
      Top             =   405
      Width           =   6900
      Active          =   0   'False
      Transposed      =   0   'False
      NULLValueString =   ""
      Descending      =   0   'False
      NoTotals        =   0   'False
      NoGrandTotals   =   0   'False
      Caption         =   ""
      BackColor       =   13882315
      Enabled         =   -1  'True
      Alive           =   0   'False
      BorderStyle     =   1
      AllowInactiveDimArea=   -1  'True
      AllowExpand     =   -1  'True
      AllowPivot      =   -1  'True
      TotalsString    =   ""
      InactiveDimAreaBkColor=   13882315
      AutoSize        =   0   'False
      UnusedDataAreaColor=   13882315
      MousePointer    =   0
      Object.Visible         =   -1  'True
      InfoURL         =   "http://www.contourcomponents.com/contourcube_user_guide.htm"
      ConnectionString=   ""
      DataSourceType  =   0
      VERSION_NO      =   2
      CCubeXMetadata  =   $"frmCRemittancePreview.frx":0000
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print this remittance"
      Height          =   600
      Left            =   150
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Click to find all customers matching the retrictions entered."
      Top             =   7215
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
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
      Height          =   615
      Left            =   6030
      Picture         =   "frmCRemittancePreview.frx":0499
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   1000
   End
End
Attribute VB_Name = "frmCRemittancePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTRID As Long
Dim lngCashBookLineID As Long
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
Dim x As New XArrayDB
Dim rs As New ADODB.Recordset
Dim tlChildCustomers As z_TextList
Dim flgLoading As Boolean
Dim bDirty As Boolean

Public Sub Component(pTRID As Long, pCustomerName As String, CashbookLineID As Long)
    On Error GoTo errHandler
Dim i As Integer
Dim oSQL As New z_SQL
    lngTRID = pTRID
    lngCashBookLineID = CashbookLineID
    Set rs = New ADODB.Recordset
    If lngTRID > 0 Then
        oSQL.GetDynamicRecordset_Improved "Select * FROM vCRemittances WHERE TR_ID = " & CStr(lngTRID), enText, Array(), "", rs
    Else
        oSQL.GetDynamicRecordset_Improved "Select * FROM vCRemittances WHERE TR_CASHBOOKLINEID = " & CStr(lngCashBookLineID), enText, Array(), "", rs
    End If
    strCustomerName = pCustomerName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCRemittancePreview.component(pTRID,pCustomerName)", Array(pTRID, pCustomerName)
End Sub

Private Sub cmdClose_Click()
Dim bInProcess As Boolean

    Unload Me
End Sub

Private Sub cmdPrint_Click()
    CC.PrintCube True
End Sub

Private Sub Form_Activate()
Dim oFSO As New FileSystemObject

End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    
    bAmt = False
    bDate = False
    If Me.WindowState <> 2 Then
        Left = 70
        Top = 70
        Width = 3990
        Height = 4620
    End If

'    Set tlChildCustomers = New z_TextList
'    tlChildCustomers.Load ltChildCustomers, CStr(lngTRID)
' Allocate space for 300 rows, 4 columns
    XA.ReDim 0, 299, 0, 7

    Dim Row As Long, Col As Integer

' Bind True DBGrid Control to this XArrayDB instance
  '  Set gDeposits.Array = XA
'    rs.CursorLocation = adUseClient
'    rs.Fields.Append "TPID", adInteger
'    rs.Fields.Append "Customer", adVarChar, 100
'    rs.Fields.Append "Reference", adVarChar, 100
'    rs.Fields.Append "Date", adDate
'    rs.Fields.Append "Amount", adDouble
'    rs.Fields.Append "SettlementDiscount", adDouble
'    rs.Open
   
    SetFormSize Me
    Preparecube

    bDirty = False
    flgLoading = False
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.Form_Load", , EA_NORERAISE
    HandleError
End Sub



Private Sub Preparecube()
    On Error GoTo errHandler
Dim oFSO As New FileSystemObject
Dim oTLS As New z_TextListSimple
Dim Fact As IViewFact
    
    If rs Is Nothing Then Exit Sub
    If rs.RecordCount = 0 Then Exit Sub
    rs.MoveFirst
    If rs.EOF Then
        MsgBox "No records", , "Status"
    End If
    
    If Not rs.EOF Then
        
        CloseCube
        With CC.Cube
            .Dims.Add("CustomerName", "CustomerName", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("TargetReference", "TargetReference", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("RemittanceDocCode", "RemittanceDocCode", , xda_vertical).MoveTo xda_vertical
            .BaseFacts.Add "Amount", "Amount"
            .Facts.Add "Amount", "Amount", xfaa_SUM
            .BaseFacts.Add "SettlementDiscount", "SettlementDiscount"
            .Facts.Add "SettlementDiscount", "SettlementDiscount", xfaa_SUM
'            .BaseFacts.Add "Balance", "Balance"
'            .Facts.Add "Balance", "Balance", xfaa_SUM
            CC.Facts(0).Appearance.Format = "###,##0.00;(###,##0.00)"
            CC.Facts(0).Caption = "Amount"
            CC.Facts(1).Appearance.Format = "###,##0.00;(###,##0.00)"
            CC.Facts(1).Caption = "Sett.disc."
'            CC.Facts(1).Appearance.Format = "###,##0.00;(###,##0.00)"
'            CC.Facts(1).Caption = "Balance"
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
        If oFSO.FileExists(CStr(oPC.LocalFolder & "Templates\AccountsCC_2.txt")) Then
            LoadContourcubeLayout CStr(oPC.LocalFolder & "Templates\AccountsCC_2.txt"), Me.CC
        End If
        AfterOpen
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformance.Preparecube"
    HandleError
End Sub

'Private Sub cmdFind1_Click()
'    On Error GoTo errHandler
'    Screen.MousePointer = vbHourglass
'    Find
'    Grid.ReBind
'    Grid.Bookmark = 1
'
'    Screen.MousePointer = vbDefault
'    Exit Sub
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmCRemittancePreview.cmdFind1_Click", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub Find()
'    On Error GoTo errHandler
'Dim bNotFound As Boolean
'Dim frm As frmBrowseCustomers2
'Dim lngTRID As Long
'Dim byear As Boolean
'Dim yr As String
'Dim mth As String
'Dim strDate1 As String
'Dim strDate2 As String
'Dim lngCount As Long
'
'    bNotFound = False
'    If Left(txtArg, 3) = "yr=" Then byear = True
'
'    If txtArg > " " And Not (byear) Then
'        'Search for Reference
'        Set cJNL = Nothing
'        Set cJNL = New c_JNL
'        cJNL.Load bNotFound, 0, "", txtArg, dteDate1, dteDate2
'        If bNotFound Then
'            'Search for customer by AcJNLO
'            Set cJNL = Nothing
'            Set cJNL = New c_JNL
'            SetDateArgs
'            cJNL.Load bNotFound, 0, txtArg, "", dteDate1, dteDate2
'            If bNotFound Then
'               Set frm = New frmBrowseCustomers2
'               frm.component txtArg, lngCount
'               If lngCount > 1 Then
'                    frm.Show vbModal
'                    lngTRID = frm.CustomerID
'                    Unload frm
'                ElseIf lngCount = 1 Then
'                    lngTRID = frm.CustomerID
'                    Unload frm
'                End If
'               If lngTRID > 0 Then
'                    Set cJNL = Nothing
'                    Set cJNL = New c_JNL
'                    SetDateArgs
'                    cJNL.Load bNotFound, lngTRID, "", "", dteDate1, dteDate2
'               End If
'            End If
'        Else
'            enSince = 1
'            cbSince.Caption = TranslateSince(1)
'        End If
'    Else
'        If byear Then
'            yr = Mid(txtArg, 4, 4)
'            mth = Mid(txtArg, 9, 2)
'            If mth > "" Then
'                strDate1 = yr & "-" & mth & "-01"
'                strDate2 = yr & "-" & mth & "-" & LastDayOfMonth(yr & "-" & mth & "-01")
'            Else
'                strDate1 = yr & "-01-01"
'                strDate2 = yr & "-12-31"
'            End If
'            If Not (IsDate(strDate1) And IsDate(strDate2)) Then
'                SetDateArgs
'            Else
'                dteDate1 = CDate(strDate1)
'                dteDate2 = CDate(strDate2)
'            End If
'        Else
'            SetDateArgs
'        End If
'        cJNL.Load bNotFound, 0, "", "", dteDate1, dteDate2
'    End If
'
'EXIT_Handler:
'    mSetfocus Grid
'    MousePointer = vbDefault
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseDBJNLs.Find"
'End Sub

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

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    If flgLoading Then Exit Sub
    
    CC.Width = NonNegative_Lng(Me.Width - 700)
    CC.Height = NonNegative_Lng(Me.Height - 2100)
    cmdClose.Top = NonNegative_Lng(Me.Height - 1300)
    cmdClose.Left = NonNegative_Lng(Me.Width - 1900)
    cmdPrint.Top = cmdClose.Top
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)

    SaveContourCubeLayout CStr(oPC.LocalFolder & "Templates\AccountsCC_2.txt"), Me.CC
    SaveFormSize Me.Name, Me.Height, Me.Width
    
End Sub

Private Sub Label3_Click()
Dim str As String
    str = "Notes" & vbCrLf _
            & "Enter document number, Acc no. or start of customer name followed by '*'. " & vbCrLf _
            & "Hit ENTER to fetch. " & vbCrLf & vbCrLf _
            & "Search for old data like this . . . " & vbCrLf _
            & "yr=2002     fetches all records for 2002" & vbCrLf & vbCrLf _
            & "yr=2002-03     fetches all records for March 2002" & vbCrLf & vbCrLf _
            & "Maximum records returned is settable in PBKS.INI file (ask support person)" & vbCrLf _
            & "This is currently set at " & oPC.MaxBrowseRecs & " records" & vbCrLf
    MsgBox str, vbInformation, "Help"

End Sub
