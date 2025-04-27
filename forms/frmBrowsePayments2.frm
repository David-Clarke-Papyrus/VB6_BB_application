VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{7A5C485E-4ACE-4C72-B64D-46119DEDD852}#4.0#0"; "CCubeX40.ocx"
Begin VB.Form frmBrowsePayments 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse customer payments"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowsePayments2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin CCubeX4.ContourCubeX CC 
      Height          =   2595
      Left            =   90
      TabIndex        =   7
      Top             =   1905
      Width           =   6840
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
      AllowDimOutside =   -1  'True
      AllowExpand     =   -1  'True
      AllowPivot      =   -1  'True
      TotalsString    =   "Totals"
      InactiveDimAreaBkColor=   13882315
      AutoSize        =   0   'False
      UnusedDataAreaColor=   13882315
      MousePointer    =   0
      Object.Visible         =   -1  'True
      InfoURL         =   "http://www.contourcomponents.com/contourcube_user_guide.htm"
      UseThemes       =   0   'False
      WordWrap        =   -1  'True
      FlatStyle       =   0
      FactsVAlignment =   0
      UnusedTreeAreaColor=   16645369
      DimLevelGradient=   14007466
      TreeLineColor   =   14007466
      DimLevelGradientStep=   20
      AllowDimVertical=   -1  'True
      AllowDimHorizontal=   -1  'True
      DrawOptions     =   2
      ConnectionString=   ""
      DataSourceType  =   0
      VERSION_NO      =   2
      CCubeXMetadata  =   $"frmBrowsePayments2.frx":058A
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   6975
      Picture         =   "frmBrowsePayments2.frx":25F8
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   540
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   90
      TabIndex        =   1
      Top             =   105
      Width           =   6810
      Begin VB.CommandButton cbSince 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Since: Last week"
         Height          =   450
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   2310
      End
      Begin VB.CommandButton cmdFind1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Height          =   615
         Left            =   5055
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmBrowsePayments2.frx":2982
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Click to find all customers matching the retrictions entered."
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1000
      End
      Begin VB.TextBox txtArg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   450
         Left            =   2490
         TabIndex        =   0
         Tag             =   "Enter product code,document number, Acc no.,or start of customer name followed by '*'. Hit ENTER to fetch."
         Top             =   240
         Width           =   2500
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   405
         Left            =   6180
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search for . . ."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   2760
         TabIndex        =   2
         Top             =   750
         Width           =   1755
      End
   End
   Begin MSComctlLib.Toolbar GridToolBar 
      Height          =   660
      Left            =   105
      TabIndex        =   8
      Top             =   1200
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   1164
      ButtonWidth     =   820
      ButtonHeight    =   1164
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList"
      HotImageList    =   "HotImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Export to Excel|Export Grid and Chart1 to Excel for printing, additioanal calculation and publishing"
            Object.ToolTipText     =   "Export to Excel|Export Grid and Chart1 to Excel for printing, additioanal calculation and publishing"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Export to Word|Export Grid and Chart1 to Word for printing and publishing"
            Object.ToolTipText     =   "Export to Word|Export Grid and Chart1 to Word for printing and publishing"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Print|Print Grid"
            Object.ToolTipText     =   "Print Grid"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "load"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "save"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   7020
      Top             =   1275
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowsePayments2.frx":2D0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowsePayments2.frx":339E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowsePayments2.frx":3A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowsePayments2.frx":40C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowsePayments2.frx":4754
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowsePayments2.frx":4AEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmBrowsePayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim fs As New FileSystemObject

Dim mcol As c_Payments
Dim tlCustomer As z_TextList
Dim lngTPID As Long
Dim strRef As String
Dim enSince As enumSince
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim strDate1 As String
Dim strDate2 As String
Dim blnNoRecordsReturned As Boolean
Dim flgLoading As Boolean
Dim XA As New XArrayDB
Dim xMLDoc As ujXML
Dim frmCustomer As frmCustomerPreview

Const opTRANSPOSE = 91
Const opCOLLAPSE = 92
Const opEXPAND = 93
Const opPERCENT = 94
Const opSORT_COL = 96
Const opSORT_ROW = 97
Const opEXPORT_HTML = 911
Const opEXPORT_XLS = 1
Const opEXPORT_DOC = 2
Const opPRINT = 4
Const opLOADLAYOUT = 5
Const opSAVELAyoUT = 6

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
      "ShellExecuteA" (ByVal hWnd As Long, ByVal lpszOp As _
      String, ByVal lpszFile As String, ByVal lpszParams As String, _
      ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Public Sub mnuSaveLayout()
    On Error Resume Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.mnuSaveLayout"
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = False
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuCancelINactive.Enabled = False
    Forms(0).mnuFulfil.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.SetMenu"
End Sub



Private Sub cbSince_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    enSince = OptionLoop(enSince, 5)
    cbSince.Caption = TranslateSince(CInt(enSince))
    mSetfocus txtArg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.cbSince_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cbSince_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If KeyAscii = 13 Then
        Find
        LoadCube
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.cbSince_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
 Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFind1_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    Screen.MousePointer = vbHourglass
    Find
    LoadCube
    
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.cmdFind1_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    cmdClose.TOP = Me.TOP + 500
    Me.CC.Left = NonNegative_Lng(Me.Left + 100)
    Me.CC.Height = NonNegative_Lng(Me.Height - 2390)
    CC.Width = NonNegative_Lng(Me.Width - 700)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadCube()
    On Error GoTo errHandler
Dim oTLS As New z_TextListSimple
Dim Fact As IViewFact
    
    If rs Is Nothing Then Exit Sub
    rs.MoveFirst
    If rs.eof Then
        MsgBox "No records", , "Status"
    End If
    
    If Not rs.eof Then
        
        CloseCube
        With CC.Cube
            .Dims.Add("DepositorName", "DepositorName", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("DepositCode", "DepositCode", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("RemittanceReference", "RemittanceReference", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("DepositDate", "DepositDate", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("CustomerName", "CustomerName", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("PaymentCode", "PaymentCode", , xda_vertical).MoveTo xda_vertical
            .BaseFacts.Add "DepositAmount", "DepositAmount"
            .Facts.Add "DepositAmount", "DepositAmount", xfaa_SUM
            .BaseFacts.Add "DepositSettlementAmount", "DepositSettlementAmount"
            .Facts.Add "DepositSettlementAmount", "DepositSettlementAmount", xfaa_SUM
            CC.Facts(0).Appearance.Format = "###,##0.00"
            CC.Facts(0).Caption = "Amount"
            CC.Facts(1).Appearance.Format = "###,##0.00"
            CC.Facts(1).Caption = "Sett.disc."
            CC.NoGrandTotals = False
            CC.TitleSettings.text = "Customer payments summary"
            CC.VAxis.DrillDownLevel = 0
            For Each Fact In CC.Facts
              Fact.Visible = True
            Next
            Set rs.ActiveConnection = Nothing
            .Open rs

        End With
        AfterOpen
        If fs.FileExists(oPC.SharedFolderRoot & "\CubeFormats\CustomerPayments\Default.cuf") Then
            LoadContourcubeLayout oPC.SharedFolderRoot & "\CubeFormats\CustomerPayments\Default.cuf"
        End If
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformance.Preparecube"
    HandleError
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
Private Sub txtArg_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If KeyAscii = 13 Then
        Find
        LoadCube
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

'Private Function ArgIsProductCode() As Boolean
'    On Error GoTo errHandler
'
'   ArgIsProductCode = (IsHashCode(txtArg) Or IsISBN10(txtArg) Or IsISBN13(txtArg))
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowsePayments.ArgIsProductCode"
'End Function
Private Sub SetDateArgs()
    On Error GoTo errHandler
    Select Case enSince
    Case enAny
        dteDate1 = CDate("1995-01-01")
        dteDate2 = DateAdd("d", 1, Date)
    Case enWeek
        dteDate1 = DateAdd("d", -7, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case enMonth
        dteDate1 = DateAdd("m", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case enQuarter
        dteDate1 = DateAdd("q", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case enYear
        dteDate1 = DateAdd("yyyy", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    End Select

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.SetDateArgs"
End Sub

Private Sub Find()
    On Error GoTo errHandler
Dim bNotFound As Boolean
Dim frm As frmBrowseCustomers2
Dim lngTPID As Long
Dim byear As Boolean
Dim yr As String
Dim mth As String
Dim strDate1 As String
Dim strDate2 As String
Dim lngCount As Long

    bNotFound = False
    If Left(txtArg, 3) = "yr=" Then byear = True
    If txtArg > " " And Not (byear) Then
        Set mcol = Nothing
        Set mcol = New c_Payments
            mcol.Load bNotFound, 0, txtArg, "", dteDate1, dteDate2, , rs
            If bNotFound Then
               Set frm = New frmBrowseCustomers2
               frm.component txtArg, lngCount
               If lngCount > 1 Then
                    frm.Show vbModal
                    lngTPID = frm.CustomerID
  '                  Me.txtArg = frm.CustomerName
                    Unload frm
                ElseIf lngCount = 1 Then
                    lngTPID = frm.CustomerID
'                    Me.txtArg = frm.CustomerName
                    Unload frm
                End If
               If lngTPID > 0 Then
                   Set mcol = Nothing
                   Set mcol = New c_Payments
                   SetDateArgs
                   mcol.Load bNotFound, lngTPID, "", "", dteDate1, dteDate2, , rs
               End If
        End If
    Else
        If byear Then
            yr = Mid(txtArg, 4, 4)
            mth = Mid(txtArg, 9, 2)
            If mth > "" Then
                strDate1 = yr & "-" & mth & "-01"
                strDate2 = yr & "-" & mth & "-" & LastDayOfMonth(yr & "-" & mth & "-01")
            Else
                strDate1 = yr & "-01-01"
                strDate2 = yr & "-12-31"
            End If
            If Not (IsDate(strDate1) And IsDate(strDate2)) Then
                SetDateArgs
            Else
                dteDate1 = CDate(strDate1)
                dteDate2 = CDate(strDate2)
            End If
        Else
            SetDateArgs
        End If
        mcol.Load bNotFound, 0, "", "", dteDate1, dteDate2, , rs
    End If

EXIT_Handler:
    mSetfocus CC
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Find"
End Sub


Private Sub cmdFind_LostFocus()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.cmdFind_LostFocus", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    Set tlCustomer = New z_TextList
    Set mcol = New c_Payments
    If Me.WindowState <> 2 Then
        Me.TOP = 50
        Me.Left = 50
        Me.Width = 7300
        Me.Height = 6100
    End If
    SetMenu
    LoadControls
    SetFormSize Me
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set tlCustomer = Nothing
    Set mcol = Nothing
    
    SaveFormSize Me.Name, Me.Height, Me.Width
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    txtArg = ""
    strDate1 = ""
    strDate2 = ""
    lngTPID = 0
    enSince = enWeek
    cbSince.Caption = TranslateSince(CInt(enSince))
    flgLoading = False

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.LoadControls"
End Sub

Private Sub Label3_Click()
    On Error GoTo errHandler
Dim str As String
    If flgLoading Then Exit Sub
    str = "Notes" & vbCrLf _
            & "Enter document number, Acc no.,or start of customer name followed by '*'." & vbCrLf _
            & "Hit ENTER to fetch. " & vbCrLf & vbCrLf _
            & "Search for old data like this . . . " & vbCrLf _
            & "yr=2002     fetches all records for 2002" & vbCrLf & vbCrLf _
            & "yr=2002-03     fetches all records for March 2002" & vbCrLf & vbCrLf _
            & "Maximum records returned is settable  (ask support person)" & vbCrLf _
            & "This is currently set at " & oPC.MaxBrowseRecs & " records" & vbCrLf
    MsgBox str, vbInformation, "Help"
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Label3_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    ExportToXML
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Public Function ExportToXML() As Boolean
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim strPath As String
Dim strBillto As String
Dim strDelto As String
Dim strFOFile As String
Dim strFilename As String
Dim strXML As String
Dim strCommand As String
Dim i As Integer
Dim strHTML As String
Dim objXSL As New MSXML2.DOMDocument60
Dim opXMLDOC As New MSXML2.DOMDocument60
Dim objXMLDOC  As New MSXML2.DOMDocument60
Dim strExecutable As String

    Set xMLDoc = New ujXML
    With xMLDoc
  .docProgID = "MSXML2.DOMDocument"
  .docInit "CO_1"
  .chCreate "CO"
      .elText = "Customer orders at " & Format(Now(), "dd/mm/yyyy HH:NN AM/PM")
  For i = 1 To mcol.Count
      
      .elCreateSibling "DetailLine", True
      .chCreate "Col_1"
          .elText = mcol(i).DepositorName & (IIf(Len(Trim(mcol(i).DepositorAcNo)) <= 1, "", "(" & Trim(mcol(i).DepositorAcNo) & ")"))
      .elCreateSibling "Col_2"
          .elText = mcol(i).DepositCode & mcol(i).StaffNameB
      .elCreateSibling "Col_3"
          .elText = mcol(i).DepositDateF
      .elCreateSibling "Col_4"
          .elText = mcol(i).StatusF
          .navUP
  Next i
    End With
    
'FINALLY PRODUCE THE .XML FILE
    strXML = oPC.SharedFolderRoot & "\TEMP\COs" & ".xml"
    With xMLDoc
  If fs.FileExists(strXML) Then
      fs.DeleteFile strXML
  End If
  .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

''WRITE THE .HTML FILE
    objXSL.async = False
    objXSL.ValidateOnParse = False
    objXSL.resolveExternals = False
    strPath = oPC.SharedFolderRoot & "\Templates\CO_RTF_1.xslt"
    Set fs = New FileSystemObject
    If fs.FileExists(strPath) Then
  objXSL.Load strPath
    End If

    strFilename = oPC.LocalFolder & "\CO.RTF"
    If fs.FileExists(strFilename) Then
  fs.DeleteFile strFilename, True
    End If
    oTF.OpenTextFileToAppend strFilename
    oTF.WriteToTextFile xMLDoc.docObject.transformNode(objXSL)
    oTF.CloseTextFile
    
    strExecutable = GetPDFExecutable(strFilename)
          If strExecutable = "" Then
              MsgBox "There is no application set on this computer to open the file: " & strFilename & ". The document cannot be displayed", vbOKOnly, "Can't do this"
          Else
              Shell strExecutable & " " & strFilename, vbNormalFocus
          End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.ExportToXML"
End Function

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

'Private Sub GridToolBar_ButtonClick(ByVal Button As MSComCtl2.Button)
' Dim DDLevel As Integer
' Dim Checked As Boolean
'
' Checked = (Button.Value = tbrPressed)
'        CC.TitleSettings.Text = "TEST"
'
' With CC
'  Select Case Button.Index
'   Case opTRANSPOSE          'Swap rows and columns
'    .Transposed = Checked
'    .Cube.RootAxis = IIf(.Transposed, _
'     IIf(GridToolBar.Buttons(6).Value = tbrPressed, xda_vertical, xda_horizontal), _
'     IIf(GridToolBar.Buttons(6).Value = tbrPressed, xda_horizontal, xda_vertical))
'   Case opCOLLAPSE           'Expand/Collapse rows and columns
'    If .HAxis.Dims.Count > 0 Then .HAxis.DrillDownLevel = 0
'    If .VAxis.Dims.Count > 0 Then .VAxis.DrillDownLevel = 0
'   Case opEXPAND
'    .HAxis.DrillDownLevel = .HAxis.Width - 1
'    .VAxis.DrillDownLevel = .VAxis.Width - 1
'   Case opPERCENT            'Calculate percents by rows/columns and show it in cells
'    .Active = False
'    Dim Fact As ICubeFact
'    For Each Fact In .Cube.Facts
'      If Left(Fact.Name, 3) <> "_P_" Then
'        If Not .Cube.Facts.Exists("_P_" & Fact.Name) Then
'          .Cube.Facts.AddFormula("_P_" & Fact.Name, Fact.Name & "/%Total(" & Fact.Name & ")").Active = True
'        End If
'        If GridToolBar.Buttons(4).Value = tbrPressed Then
'          .Facts.Item("_P_" & Fact.Name).Visible = True
'          .Facts.Item("_P_" & Fact.Name).Caption = Fact.Caption
'          .Facts.Item("_P_" & Fact.Name).Appearance.Format = "#####0.00%"
'          .Facts.Item(Fact.Name).Enabled = False
'        Else
'          .Facts.Item(Fact.Name).Visible = True
'          .Facts.Item("_P_" & Fact.Name).Enabled = False
'        End If
'      End If
'    Next
'    .Active = True
'
'   Case opSORT_COL, opSORT_ROW        'Sort rows by selected fact values in selected column
'    Dim SortAxis: SortAxis = IIf(Button.Index = 6, xda_vertical, xda_horizontal)
'    Dim col As Long, row As Long: col = .CurrentCell.col: row = .CurrentCell.row
'    If (GridToolBar.Buttons(Button.Index).Value = tbrPressed) Then _
'      .SortGridByFact SortAxis, col, row _
'    Else _
'      .CancelFactSorting (SortAxis)
'   'Export Grid for printing and publishing
'   Case opEXPORT_HTML
'    ExportCube .TitleSettings.Text, xolaprpt_HTML, "html"
'   Case opEXPORT_XLS
'    ExportCube .TitleSettings.Text, xolaprpt_XLS, "xls"
'   Case opEXPORT_DOC
'    ExportCube .TitleSettings.Text, xolaprpt_HTML, "doc"
'   Case opPRINT
'    .PrintCube True, False
'   Case opSAVELAyoUT
'        SaveFormat
'   Case opLOADLAYOUT
'        LoadFormat
'  End Select
' End With
'End Sub

'Private Sub GridToolBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
' Dim ScaleFactor As Double
' Dim SortAxis
' Dim col As Long, row As Long
' ScaleFactor = 1
' With CC
'  Select Case ButtonMenu.key
'   Case "1x1"
'    ScaleFactor = 1
'   Case "1x10"
'    ScaleFactor = 0.1
'   Case "1x100"
'    ScaleFactor = 0.01
'   Case "1x1000"
'    ScaleFactor = 0.001
'   Case "asc"
'       SortAxis = xda_vertical
'       col = .CurrentCell.col: row = .CurrentCell.row
'            CC.Descending = False
'            CC.SortGridByFact SortAxis, col, row
'   Case "desc"
'       SortAxis = xda_vertical
'       col = .CurrentCell.col: row = .CurrentCell.row
'            CC.Descending = True
'            CC.SortGridByFact SortAxis, col, row
'   Case "Nosort"
'       SortAxis = xda_vertical
'       col = .CurrentCell.col: row = .CurrentCell.row
'            CC.CancelFactSorting (SortAxis)
'   Case "hasc"
'       SortAxis = xda_horizontal
'       col = .CurrentCell.col: row = .CurrentCell.row
'            CC.Descending = False
'            CC.SortGridByFact SortAxis, col, row
'   Case "hdesc"
'       SortAxis = xda_horizontal
'       col = .CurrentCell.col: row = .CurrentCell.row
'            CC.Descending = True
'            CC.SortGridByFact SortAxis, col, row
'   Case "hNosort"
'       SortAxis = xda_horizontal
'       col = .CurrentCell.col: row = .CurrentCell.row
'            CC.CancelFactSorting (SortAxis)
'  End Select
'  Dim Fact
'  For Each Fact In .Facts
'   If Fact.Enabled Then Fact.ScaleFactor = ScaleFactor
'  Next
' End With
'End Sub
Private Sub ExportCube(FileName As String, FileFormat As TxOlapReportType, FileType As String)
 'Export OLAP-report to Excel, Word, HTML as file in html format
 FileName = FileName + "." + FileType
 CC.ReportToFile FileName, "", FileFormat
 OpenDocument (FileName)
End Sub
Private Sub OpenDocument(f_name As String)
 Dim Scr_hDC As Long
 Scr_hDC = GetDesktopWindow()
 ShellExecute Scr_hDC, "Open", f_name, "", "", 1
End Sub

Private Sub LoadFormat()
  CommonDialog1.DefaultExt = "cuf"
  CommonDialog1.DialogTitle = "Load Cube layout"
  CommonDialog1.InitDir = oPC.SharedFolderRoot & "\CubeFormats\CustomerPayments"
  CommonDialog1.CancelError = True
  On Error Resume Next
  CommonDialog1.ShowOpen
  If Err.Number = cdlCancel Then
    On Error GoTo 0
    Exit Sub
  Else
    On Error GoTo 0
    LoadContourcubeLayout CommonDialog1.FileName
  End If

End Sub
Private Sub SaveFormat()
Dim fs As New FileSystemObject
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\CubeFormats") Then
        fs.CreateFolder (oPC.SharedFolderRoot & "\CubeFormats")
    End If
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\CubeFormats\CustomerPayments") Then
        fs.CreateFolder (oPC.SharedFolderRoot & "\CubeFormats\CustomerPayments")
    End If
  CommonDialog1.DefaultExt = "cuf"
  CommonDialog1.DialogTitle = "Save Cube layout"
  CommonDialog1.InitDir = oPC.SharedFolderRoot & "\CubeFormats\CustomerPayments"
  CommonDialog1.CancelError = True
  On Error Resume Next
  CommonDialog1.ShowSave
  If Err.Number = cdlCancel Then
    On Error GoTo 0
    Exit Sub
  Else
    On Error GoTo 0
    If Trim(CommonDialog1.FileName) <> "" Then SaveContourCubeLayout CStr(CommonDialog1.FileName)
  End If

End Sub

Public Sub SaveContourCubeLayout(ltFile As String)
'Saving layout procedure
  Dim rsFields, Axis, Object, bInvertFilterSelection, Value, i, j, viewTotalsState, _
      viewGTotalsState, strExpand, fs
  rsFields = Array("Object", "Name", "Property", "Value")
  'Create an ADO recordset with 4 fields:
  Dim rs As New ADODB.Recordset
  rs.Fields.Append rsFields(0), adBSTR, 10
  rs.Fields.Append rsFields(1), adBSTR, 50
  rs.Fields.Append rsFields(2), adVariant, 50
  rs.Fields.Append rsFields(3), adVariant, 255
  rs.Open
  rs.AddNew rsFields, Array("Cube", CC.Name, "RootAxis", CC.Cube.RootAxis)
  With CC
    'Populate recordset with layout properties
    For Each Object In .Facts
      'Fact visibility
      rs.AddNew rsFields, Array("Fact", Object.Name, "Visible", Object.Visible)
    Next
    For i = 0 To 1
        If i = 0 Then Set Axis = .VAxis Else Set Axis = .HAxis
        For Each Object In Axis.Dims
          'Dimension positions and properties
          rs.AddNew rsFields, Array("Dim", Object.Name, "Axis", Object.CubeDim.Axis)
          rs.AddNew rsFields, Array("Dim", Object.Name, "Pos", Object.CubeDim.pos)
        Next
    Next
    For Each Object In .Dims
        rs.AddNew rsFields, Array("Dim", Object.Name, "Totals", Object.NoTotals)
        rs.AddNew rsFields, Array("Dim", Object.Name, "Descending", Object.Descending)
        'Dimension filters:
        'To minimize the file, choose the minimum set between hidden and visible
        'values to save
        bInvertFilterSelection = (Object.CubeDim.GetValues(2).Count > Object.CubeDim.GetValues(1).Count)
        rs.AddNew rsFields, Array("DimsFilter", "InvertFilterSelection", Object.Name, bInvertFilterSelection)
        For Each Value In Object.CubeDim.GetValues(IIf(bInvertFilterSelection, 1, 2))
          rs.AddNew rsFields, Array("DimsFilter", "Filter", Object.Name, Value)
        Next
    Next
    'Save axis expand states
    'Temporarily turn off totals, in order not to save sections that
    'correspond to dimension totals
    viewTotalsState = .NoTotals
    viewGTotalsState = .NoGrandTotals
    .NoTotals = True
    .NoGrandTotals = True
    'Cycle through all sections on both axes and save their state
    If .HAxis.Length > 0 Then
      For i = 0 To .HAxis.Length - 1
        strExpand = ""
        For j = 0 To .HAxis.GetSection(i).CurrentWidth - 1
          strExpand = strExpand & IIf(strExpand = "", "", Chr(10)) & .HAxis.GetSection(i).getValue(j)
        Next j
        rs.AddNew rsFields, Array("Axis", "Horizontal", "Section" & Trim(str(i)), strExpand)
      Next i
    End If
    If .VAxis.Length > 0 Then
      For i = 0 To .VAxis.Length - 1
        strExpand = ""
        For j = 0 To .VAxis.GetSection(i).CurrentWidth - 1
          strExpand = strExpand & IIf(strExpand = "", "", Chr(10)) & .VAxis.GetSection(i).getValue(j)
        Next j
        rs.AddNew rsFields, Array("Axis", "Vertical", "Section" & Trim(str(i)), strExpand)
      Next i
    End If
    'Restore view totals
    .NoTotals = viewTotalsState
    .NoGrandTotals = viewGTotalsState
  End With
  'Verify if the file already exists and eventually delete it before saving
  Set fs = CreateObject("Scripting.FileSystemObject")
  If fs.FileExists(ltFile) Then fs.DeleteFile (ltFile)
  rs.Save ltFile, adPersistXML
  rs.Close
End Sub

Sub LoadContourcubeLayout(ltFile As String)
'Loading layout procedure
  Dim FactSettings, DimSettings, Object, DimFilters, AxisSettings, i, bInvertFilterSelection
  Dim rs As New ADODB.Recordset
  'First open the saved XML layout file
  rs.Open ltFile
  With CC
    'Restore cube properties
    rs.Filter = "Object='Cube'"
    .Cube.RootAxis = CInt(rs.GetRows()(3, 0))
    'Fact visibility
    rs.Filter = "Object='Fact'"
    FactSettings = rs.GetRows()
    For i = 0 To UBound(FactSettings, 2)
      If LCase(CStr(FactSettings(2, i))) = "visible" Then
        If .Facts.Exists(CStr(FactSettings(1, i))) Then _
           .Facts(CStr(FactSettings(1, i))).Visible = CBool(FactSettings(3, i))
      End If
    Next i
    'Set up dimension positions, totalling and sort orders
    rs.Filter = "Object='Dim'"
    DimSettings = rs.GetRows()
    For Each Object In .Dims
        If Object.CubeDim.Axis <> xda_invisible Then Object.CubeDim.MoveTo xda_outside
    Next
    For i = 0 To UBound(DimSettings, 2)
      If .Dims.Exists(CStr(DimSettings(1, i))) Then
        Select Case LCase(CStr(DimSettings(2, i)))
        Case "axis":
          .Dims(CStr(DimSettings(1, i))).CubeDim.MoveTo CInt(DimSettings(3, i))
        Case "pos":
          .Dims(CStr(DimSettings(1, i))).CubeDim.MoveTo .Dims(CStr(DimSettings(1, i))).CubeDim.Axis, CInt(DimSettings(3, i))
        Case "totals":
          .Dims(CStr(DimSettings(1, i))).NoTotals = CBool(DimSettings(3, i))
        Case "descending":
          .Dims(CStr(DimSettings(1, i))).Descending = CBool(DimSettings(3, i))
        End Select
      End If
    Next i
    .Active = True
    'Dimension filter states
    rs.Filter = "Object='DimsFilter'"
    DimFilters = rs.GetRows()
    For i = 0 To UBound(DimFilters, 2)
      If .Dims.Exists(CStr(DimFilters(2, i))) Then
        Select Case LCase(CStr(DimFilters(1, i)))
        Case "invertfilterselection":
          bInvertFilterSelection = CBool(DimFilters(3, i))
          .Dims(CStr(DimFilters(2, i))).CubeDim.Filter IIf(bInvertFilterSelection, xfo_FilterAll, xfo_Reset)
        Case "filter":
          .Dims(CStr(DimFilters(2, i))).CubeDim.FilterValue DimFilters(3, i), Not bInvertFilterSelection
        End Select
      End If
    Next i
    .Cube.DimensionsFilter.Apply
    'Finally, restore expand status of each axis section
    .HAxis.DrillDownLevel = .HAxis.Width - 1
    .VAxis.DrillDownLevel = .VAxis.Width - 1
    rs.Filter = "Object='Axis'"
    AxisSettings = rs.GetRows()
    For i = 0 To UBound(AxisSettings, 2)
      ExpandSection CStr(AxisSettings(1, i)), CStr(AxisSettings(3, i))
    Next i
  End With
  rs.Close
End Sub

Sub ExpandSection(strAxis As String, strExpand As String)
'This procedure restores saved state of an axis section
'It searches for given combination of dim values along the axis,
'and expands the section found
  Dim Axis As IViewAxis, i, j, aExpand
  aExpand = Split(strExpand, Chr(10))
  If LCase(strAxis) = "horizontal" Then Set Axis = CC.HAxis Else Set Axis = CC.VAxis
  On Error Resume Next
  i = 0
  Do While i < Axis.Length
    j = 0
    Do While j <= UBound(aExpand, 1)
      If CStr(Axis.GetSection(i).getValue(j)) <> aExpand(j) Then Exit Do
      j = j + 1
    Loop
    If j > UBound(aExpand, 1) Then Exit Do
    i = i + 1
  Loop
  If i < Axis.Length Then Axis.GetSection(i).Collapse UBound(aExpand, 1), True
  On Error GoTo 0
End Sub


