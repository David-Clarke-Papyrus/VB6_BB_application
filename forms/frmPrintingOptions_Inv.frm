VERSION 5.00
Begin VB.Form frmPrintingOptions_Inv 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Invoice Print"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7965
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture 
      BackColor       =   &H00D3D3CB&
      Height          =   435
      Index           =   1
      Left            =   4125
      ScaleHeight     =   375
      ScaleWidth      =   2850
      TabIndex        =   17
      Top             =   2175
      Width           =   2910
      Begin VB.OptionButton optDESC 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Descending"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   1410
         TabIndex        =   19
         Top             =   45
         Width           =   1320
      End
      Begin VB.OptionButton optASC 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Ascending"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   75
         TabIndex        =   18
         Top             =   45
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture 
      BackColor       =   &H00D3D3CB&
      Height          =   1500
      Index           =   0
      Left            =   4065
      ScaleHeight     =   1440
      ScaleWidth      =   2910
      TabIndex        =   11
      Top             =   300
      Width           =   2970
      Begin VB.OptionButton optTitle 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   195
         TabIndex        =   16
         Top             =   30
         Width           =   900
      End
      Begin VB.OptionButton optAuthor 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Author"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   195
         TabIndex        =   15
         Top             =   315
         Width           =   900
      End
      Begin VB.OptionButton optCode 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   195
         TabIndex        =   14
         Top             =   600
         Width           =   900
      End
      Begin VB.OptionButton optSeq 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Indicated sequence"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   195
         TabIndex        =   13
         Top             =   885
         Width           =   2220
      End
      Begin VB.OptionButton optRef 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Reference"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   195
         TabIndex        =   12
         Top             =   1155
         Width           =   2220
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Sort by"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1875
      Left            =   3960
      TabIndex        =   10
      Top             =   30
      Width           =   3270
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Sequence"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   900
      Left            =   3945
      TabIndex        =   9
      Top             =   1920
      Width           =   3300
   End
   Begin VB.CommandButton cmdExportToSpreadsheet 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Spreadsheet"
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
      Left            =   2265
      Picture         =   "frmPrintingOptions_Inv.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2460
      Width           =   1380
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Print"
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
      Left            =   1545
      Picture         =   "frmPrintingOptions_Inv.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1380
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&PDF"
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
      Left            =   840
      Picture         =   "frmPrintingOptions_Inv.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2460
      Width           =   1380
   End
   Begin VB.CommandButton cmdPickingSlip 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Picking slip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1395
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3210
      Width           =   1980
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   495
      Left            =   1950
      TabIndex        =   1
      Text            =   "1"
      Top             =   1185
      Width           =   660
   End
   Begin VB.CheckBox optSetSeqDef 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Set this choice as default"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   4650
      TabIndex        =   3
      Top             =   2910
      Width           =   2445
   End
   Begin VB.ComboBox cboCurr 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1410
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   375
      Width           =   1860
   End
   Begin VB.Label LabelTip 
      BackStyle       =   0  'Transparent
      Caption         =   "TIP: To skip this form, simply hold down the CTRL key when clicking  'Print' on the previous form"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   465
      TabIndex        =   20
      Top             =   3915
      Width           =   7350
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Copies to print"
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
      Height          =   285
      Left            =   1410
      TabIndex        =   4
      Top             =   915
      Width           =   1800
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Print in this currency"
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
      Height          =   285
      Left            =   1425
      TabIndex        =   2
      Top             =   105
      Width           =   1800
   End
End
Attribute VB_Name = "frmPrintingOptions_Inv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCurrentForeignCurrency As a_Currency
Dim oInvoice As a_Invoice
Dim flgLoading As Boolean

Public Sub ComponentObject(pInvoice As a_Invoice)
    On Error GoTo errHandler
Dim oDOC As a_DocumentControl
    Set oInvoice = pInvoice
    Set oDOC = oPC.Configuration.DocumentControls.FindDC(oInvoice.constDOCCODE)
    If oDOC Is Nothing Then
        txtQty = "1"
    Else
        txtQty = CStr(oPC.Configuration.DocumentControls.FindDC(oInvoice.constDOCCODE).QtyCopies)
    End If

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_Inv.ComponentObject(pInvoice)", pInvoice
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_Inv.ComponentObject(pInvoice)", pInvoice
End Sub

Private Sub cboCurr_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    Set oCurrentForeignCurrency = oPC.Configuration.Currencies.FindByDescription(cboCurr)
    oInvoice.BeginEdit
    oInvoice.CurrencyID = oCurrentForeignCurrency.ID
    oInvoice.ApplyEdit
    oInvoice.RecalculateAllLines
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_Inv.cboCurr_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_Inv.cboCurr_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub SortDetailLines()
    On Error GoTo errHandler
Dim strSrtSeq As String
    If optTitle Then
        oInvoice.InvoiceLines.SortInvoiceLines enTitle, optASC
        strSrtSeq = "Title"
    ElseIf optAuthor Then
        oInvoice.InvoiceLines.SortInvoiceLines enAuthor, optASC
        strSrtSeq = "Author"
    ElseIf optCode Then
        oInvoice.InvoiceLines.SortInvoiceLines enCode, optASC
        strSrtSeq = "Code"
    ElseIf optRef Then
        oInvoice.InvoiceLines.SortInvoiceLines enRef, optASC
        strSrtSeq = "Ref"
    ElseIf optSeq Then
        oInvoice.InvoiceLines.SortInvoiceLines enSequence, optASC
        strSrtSeq = "SeqNum"
    End If
        
    If optSetSeqDef = 1 Then
        SaveSetting App.EXEName, "PrintSettings", "InvoiceSequenceField", strSrtSeq
        SaveSetting App.EXEName, "PrintSettings", "InvoiceSequenceSeq", IIf(optASC, "ASCEND", "DESCEND")
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_Inv.SortDetailLines"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_Inv.SortDetailLines"
End Sub

Private Sub cmdExportToSpreadsheet_Click()
Dim sFilename As String
    Screen.MousePointer = vbHourglass
    cmdPreview.Enabled = False
    SortDetailLines
    If oInvoice.ExportToSpreadsheet(False, sFilename) = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
    End If
    If MsgBox("Spreadsheet file saved in: " & sFilename & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
        OpenFileWithApplication sFilename, enExcel
    End If
    Screen.MousePointer = vbDefault
    Unload Me


End Sub

Private Sub cmdPickingSlip_Click()
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim i As Long
Dim arPL As New arInvoicePickList
Dim rs As ADODB.Recordset

  '  If oInvoice.Status = stISSUED And oPC.AllowsInvoicePicking Then
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
        Set rs = New ADODB.Recordset
        rs.Open "SELECT * FROm vzInvoiceLine_Fetch WHERE IL_TR_ID = " & oInvoice.InvoiceID & " ORDER BY P_EAN,P_TITLE", oPC.COShort, adOpenForwardOnly
        arPL.component rs, oInvoice.Customer.NameAndCode(50), oInvoice.DOCCodeF, oInvoice.DocDateF
        arPL.Show vbModal
  '  End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_Inv.cmdPickingSlip_Click", , EA_NORERAISE
    HandleError
End Sub


'Private Sub cmdPickingSlip_Click()
'Dim OpenResult As Integer
'Dim i As Long
'Dim arPL As New arInvoicePickList
'Dim rs As ADODB.Recordset
'
'  '  If oInvoice.Status = stISSUED And oPC.AllowsInvoicePicking Then
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'        Set rs = New ADODB.Recordset
'        rs.Open "SELECT * FROm vzInvoiceLine_Fetch WHERE IL_TR_ID = " & oInvoice.InvoiceID & " ORDER BY P_EAN", oPC.COSHORT, adOpenForwardOnly
'        arPL.Component rs, oInvoice.Customer.NameAndCode(50), oInvoice.DOCCodeF, oInvoice.DocDateF
'        arPL.Show vbModal
'  '  End If
'End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    
    If oInvoice.Status = stISSUED And oPC.AllowsInvoicePicking Then
        If MsgBox("This document is still in the picking stage. Are you sure you want to print the formatted document?", vbQuestion + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    cmdPrint.Enabled = False
    SortDetailLines
 '+++++++++
    If oInvoice.ExportToXML(False, "", False, enPrint, CInt(txtQty), , , , True) = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
    End If
    Unload Me
    Screen.MousePointer = vbDefault
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_Inv.cmdPrint_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_Inv.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdPreview_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    cmdPreview.Enabled = False
    SortDetailLines
    If oInvoice.ExportToXML(False, True, False, enView, CInt(txtQty)) = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Cannot view, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
    End If
    Unload Me
    Screen.MousePointer = vbDefault
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_Inv.cmdPreview_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_Inv.cmdPreview_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadCurrs()
    On Error GoTo errHandler
Dim oCurr As a_Currency
Dim oItem As ListItem
Dim i As Integer
    For Each oCurr In oPC.Configuration.Currencies
        Me.cboCurr.AddItem oCurr.Description
    Next
    If oInvoice.ForeignCurrency.Description > "" Then
        cboCurr = oInvoice.ForeignCurrency.Description
    Else
        cboCurr = oPC.Configuration.DefaultCurrency.Description
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_Inv.LoadCurrs"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_Inv.LoadCurrs"
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim strInvSeqField As String
Dim strInvSeq As String

    flgLoading = True
    LoadCurrs
    strInvSeqField = GetSetting(App.EXEName, "PrintSettings", "InvoiceSequenceField", "Title")
    strInvSeq = GetSetting(App.EXEName, "PrintSettings", "InvoiceSequenceSeq", "Title")
    Select Case strInvSeqField
    Case "Title"
        optTitle = True
    Case "Author"
        optAuthor = True
    Case "Code"
        optCode = True
    Case "Ref"
        optRef = True
    Case "SeqNum"
        optSeq = True
    End Select
    Select Case strInvSeq
    Case "ASCEND"
        optASC = True
    Case Else
        optDESC = True
    End Select
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_Inv.Form_Load"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_Inv.Form_Load", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If Not IsNumeric(txtQty) Then Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_Inv.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
