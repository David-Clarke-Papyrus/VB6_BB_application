VERSION 5.00
Begin VB.Form frmPrintingOptions_QU 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Print quotation"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7710
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture 
      BackColor       =   &H00D3D3CB&
      Height          =   435
      Index           =   1
      Left            =   3915
      ScaleHeight     =   375
      ScaleWidth      =   2850
      TabIndex        =   14
      Top             =   2235
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   45
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture 
      BackColor       =   &H00D3D3CB&
      Height          =   1500
      Index           =   0
      Left            =   3855
      ScaleHeight     =   1440
      ScaleWidth      =   2910
      TabIndex        =   8
      Top             =   360
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
      Left            =   3750
      TabIndex        =   7
      Top             =   90
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
      Left            =   3735
      TabIndex        =   6
      Top             =   1980
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
      Left            =   1905
      Picture         =   "frmPrintingOptions_QU.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1845
      Width           =   1500
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
      Left            =   4890
      TabIndex        =   3
      Top             =   2865
      Width           =   2445
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
      Left            =   1545
      TabIndex        =   2
      Text            =   "1"
      Top             =   600
      Width           =   660
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
      Left            =   345
      Picture         =   "frmPrintingOptions_QU.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1845
      Width           =   1500
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
      Left            =   1155
      Picture         =   "frmPrintingOptions_QU.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1185
      Width           =   1500
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
      Left            =   120
      TabIndex        =   17
      Top             =   3360
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
      Left            =   975
      TabIndex        =   4
      Top             =   315
      Width           =   1800
   End
End
Attribute VB_Name = "frmPrintingOptions_QU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim oCurrentForeignCurrency As a_Currency
Dim oQU As a_QU
Dim flgLoading As Boolean
Dim strSeqField As String
Dim strSeq As String
Dim strSrtSeq As String


Public Sub ComponentObject(pDoc As a_QU)
    On Error GoTo errHandler
Dim oDC As a_DocumentControl
    Set oQU = pDoc
    Set oDC = oPC.Configuration.DocumentControls.FindDC(oQU.constDOCCODE)
    If Not oDC Is Nothing Then
        txtQty = CStr(oPC.Configuration.DocumentControls.FindDC(oQU.constDOCCODE).QtyCopies)
    Else
        txtQty = "1"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_QU.ComponentObject(pDoc)", pDoc
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_QU.ComponentObject(pDoc)", pDoc
End Sub

Private Sub cmdExportToSpreadsheet_Click()
Dim sFilename As String
    Screen.MousePointer = vbHourglass
    cmdPreview.Enabled = False
    SortDetailLines
    If oQU.ExportToSpreadsheet(False, sFilename) = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
    End If
    If MsgBox("Spreadsheet file saved in: " & sFilename & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
        OpenFileWithApplication sFilename, enExcel
    End If
    Screen.MousePointer = vbDefault
    Unload Me

End Sub

Private Sub Form_Load()
    On Error GoTo errHandler

    flgLoading = True
    strSeqField = GetSetting("PBKS", "PrintSettings", "QUSequenceField", "")
    strSeq = GetSetting("PBKS", "PrintSettings", "QUSequenceSeq", "")
    Select Case strSeqField
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
    Select Case strSeq
    Case "ASCEND"
        optASC = True
    Case Else
        optDESC = True
    End Select
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_QU.Form_Load"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_QU.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_QU.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    cmdPrint.Enabled = False
    
    SortDetailLines
    
    If oQU.ExportToXML(False, True, False, enPrint, CInt(txtQty), , , True) = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
    End If
    Unload Me
    Screen.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_QU.cmdPrint_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_QU.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub SortDetailLines()
    On Error GoTo errHandler
    If optTitle Then
        oQU.QuoteLines.SortLines enTitle, optASC
        strSrtSeq = "Title"
    ElseIf optAuthor Then
        oQU.QuoteLines.SortLines enAuthor, optASC
        strSrtSeq = "Author"
    ElseIf optCode Then
        oQU.QuoteLines.SortLines enCode, optASC
        strSrtSeq = "Code"
    ElseIf optRef Then
        oQU.QuoteLines.SortLines enRef, optASC
        strSrtSeq = "Ref"
    ElseIf optSeq Then
        oQU.QuoteLines.SortLines enSequence, optASC
        strSrtSeq = "SeqNum"
    End If
        
    If optSetSeqDef = 1 Then
        SaveSetting "PBKS", "PrintSettings", "QUSequenceField", strSrtSeq
        SaveSetting "PBKS", "PrintSettings", "QUSequenceSeq", IIf(optASC, "ASCEND", "DESCEND")
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_QU.SortDetailLines"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_QU.SortDetailLines"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_QU.SortDetailLines"
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    cmdPreview.Enabled = False
    SortDetailLines
    If oQU.ExportToXML(False, True, False, enView, CInt(txtQty)) = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
    End If
    Screen.MousePointer = vbDefault
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_QU.cmdPreview_Click"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_QU.cmdPreview_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_QU.cmdPreview_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If Not IsNumeric(txtQty) Then Cancel = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_QU.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_QU.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

