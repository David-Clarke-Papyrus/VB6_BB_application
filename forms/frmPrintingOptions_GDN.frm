VERSION 5.00
Begin VB.Form frmPrintingOptions_GDN 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Goods delivery note print"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture 
      BackColor       =   &H00D3D3CB&
      Height          =   435
      Index           =   1
      Left            =   3900
      ScaleHeight     =   375
      ScaleWidth      =   2850
      TabIndex        =   17
      Top             =   2190
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
      Left            =   3840
      ScaleHeight     =   1440
      ScaleWidth      =   2910
      TabIndex        =   11
      Top             =   315
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
      Left            =   3735
      TabIndex        =   10
      Top             =   45
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
      Left            =   3720
      TabIndex        =   9
      Top             =   1935
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
      Left            =   1830
      Picture         =   "frmPrintingOptions_GDN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2430
      Width           =   1365
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
      Left            =   1125
      Picture         =   "frmPrintingOptions_GDN.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1785
      Width           =   1365
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
      Left            =   420
      Picture         =   "frmPrintingOptions_GDN.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2430
      Width           =   1365
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
      Left            =   825
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3285
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
      Left            =   1455
      TabIndex        =   1
      Text            =   "1"
      Top             =   1245
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
      Left            =   4290
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
      Left            =   915
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   435
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
      Left            =   15
      TabIndex        =   20
      Top             =   3750
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
      Left            =   915
      TabIndex        =   4
      Top             =   975
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
      Left            =   930
      TabIndex        =   2
      Top             =   165
      Width           =   1800
   End
End
Attribute VB_Name = "frmPrintingOptions_GDN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCurrentForeignCurrency As a_Currency
Dim oDocument As a_GDN
Dim flgLoading As Boolean

Public Sub ComponentObject(pGDN As a_GDN)
    On Error GoTo errHandler
Dim oDOC As a_DocumentControl
    Set oDocument = pGDN
    Set oDOC = oPC.Configuration.DocumentControls.FindDC(oDocument.constDOCCODE)
    If oDOC Is Nothing Then
        txtQty = "1"
    Else
        txtQty = CStr(oPC.Configuration.DocumentControls.FindDC(oDocument.constDOCCODE).QtyCopies)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_GDN.ComponentObject(pGDN)", pGDN
End Sub

Private Sub cboCurr_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    Set oCurrentForeignCurrency = oPC.Configuration.Currencies.FindByDescription(cboCurr)
    oDocument.BeginEdit
    oDocument.CurrencyID = oCurrentForeignCurrency.ID
    oDocument.ApplyEdit
    oDocument.RecalculateAllLines
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_GDN.cboCurr_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub SortDetailLines()
    On Error GoTo errHandler
Dim strSrtSeq As String
    If optTitle Then
        oDocument.GDNLines.SortDocumentLines enTitle, optASC
        strSrtSeq = "Title"
    ElseIf optAuthor Then
        oDocument.GDNLines.SortDocumentLines enAuthor, optASC
        strSrtSeq = "Author"
    ElseIf optCode Then
        oDocument.GDNLines.SortDocumentLines enCode, optASC
        strSrtSeq = "Code"
    ElseIf optRef Then
        oDocument.GDNLines.SortDocumentLines enRef, optASC
        strSrtSeq = "Ref"
    ElseIf optSeq Then
        oDocument.GDNLines.SortDocumentLines enSequence, optASC
        strSrtSeq = "SeqNum"
    End If
        
    If optSetSeqDef = 1 Then
        SaveSetting App.EXEName, "PrintSettings", "InvoiceSequenceField", strSrtSeq
        SaveSetting App.EXEName, "PrintSettings", "InvoiceSequenceSeq", IIf(optASC, "ASCEND", "DESCEND")
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_GDN.SortDetailLines"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_GDN.SortDetailLines"
End Sub

Private Sub cmdPickingSlip_Click()
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim i As Long
Dim arPL As New arInvoicePickList
Dim rs As ADODB.Recordset

  '  If oDocument.Status = stISSUED And oPC.AllowsInvoicePicking Then
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
        Set rs = New ADODB.Recordset
        rs.Open "SELECT * FROm vzInvoiceLine_Fetch WHERE IL_TR_ID = " & oDocument.GDNID & " ORDER BY P_EAN,P_TITLE", oPC.COShort, adOpenForwardOnly
        arPL.component rs, oDocument.Customer.NameAndCode(50), oDocument.DOCCodeF, oDocument.DocDateF
        arPL.Show vbModal
  '  End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_GDN.cmdPickingSlip_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    
    If oDocument.Status = stISSUED And oPC.AllowsInvoicePicking Then
        If MsgBox("This document is still in the picking stage. Are you sure you want to print the formatted document?", vbQuestion + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    cmdPrint.Enabled = False
    SortDetailLines
     MsgBox "Clicked print"

    If oDocument.ExportToXML(False, True, False, enPrint, CInt(txtQty)) = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
    End If
    Unload Me
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_GDN.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdPreview_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    cmdPreview.Enabled = False
    SortDetailLines
    If oDocument.ExportToXML(False, True, False, enView, CInt(txtQty)) = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Cannot view, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
    End If
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_GDN.cmdPreview_Click", , EA_NORERAISE
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
    If oDocument.ForeignCurrency.Description > "" Then
        cboCurr = oDocument.ForeignCurrency.Description
    Else
        cboCurr = oPC.Configuration.DefaultCurrency.Description
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_GDN.LoadCurrs"
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_GDN.Form_Load", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If Not IsNumeric(txtQty) Then Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_GDN.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
