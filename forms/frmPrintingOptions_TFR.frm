VERSION 5.00
Begin VB.Form frmPrintingOptions_TFR 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Inter-branch transfer"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7380
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00D3D3CB&
      Height          =   435
      Left            =   3405
      ScaleHeight     =   375
      ScaleWidth      =   2850
      TabIndex        =   9
      Top             =   2295
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   45
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture 
      BackColor       =   &H00D3D3CB&
      Height          =   1320
      Left            =   3345
      ScaleHeight     =   1260
      ScaleWidth      =   2910
      TabIndex        =   8
      Top             =   540
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
      Left            =   1650
      Picture         =   "frmPrintingOptions_TFR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1770
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
      Left            =   900
      Picture         =   "frmPrintingOptions_TFR.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1110
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
      Left            =   195
      Picture         =   "frmPrintingOptions_TFR.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1770
      Width           =   1380
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
      Left            =   1290
      TabIndex        =   0
      Text            =   "1"
      Top             =   525
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
      Left            =   540
      TabIndex        =   3
      Top             =   2535
      Width           =   2445
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
      Left            =   3225
      TabIndex        =   2
      Top             =   1995
      Width           =   3300
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
      Height          =   1725
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   3270
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
      Left            =   90
      TabIndex        =   16
      Top             =   3075
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
      Left            =   765
      TabIndex        =   4
      Top             =   255
      Width           =   1800
   End
End
Attribute VB_Name = "frmPrintingOptions_TFR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCurrentForeignCurrency As a_Currency
Dim oDOC As a_TF
Dim flgLoading As Boolean

Public Sub ComponentObject(pD As a_TF)
    On Error GoTo errHandler
Dim oDC As a_DocumentControl
    Set oDOC = pD
    Set oDC = oPC.Configuration.DocumentControls.FindDC(pD.constDOCCODE)
    If oDC Is Nothing Then
        txtQty = "1"
    Else
        txtQty = CStr(oPC.Configuration.DocumentControls.FindDC(pD.constDOCCODE).QtyCopies)
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_TFR.ComponentObject(pInvoice)", pD
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_TFR.ComponentObject(pD)", pD
End Sub

'Private Sub cboCurr_Click()
'    On Error GoTo errHandler
'    If flgLoading Then Exit Sub
'    Set oCurrentForeignCurrency = oPC.Configuration.Currencies.FindByDescription(cboCurr)
'    oDoc.BeginEdit
'    oDoc.CurrencyID = oCurrentForeignCurrency.ID
'    oDoc.ApplyEdit
'    oDoc.RecalculateAllLines
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_TFR.cboCurr_Click"
'End Sub

Private Sub SortDetailLines()
    On Error GoTo errHandler
Dim strSrtSeq As String
    If optTitle Then
        oDOC.TFLines.SortLines enTitle, optASC
        strSrtSeq = "Title"
    ElseIf optAuthor Then
        oDOC.TFLines.SortLines enAuthor, optASC
        strSrtSeq = "Author"
    ElseIf optCode Then
        oDOC.TFLines.SortLines enCode, optASC
        strSrtSeq = "Code"
    ElseIf optSeq Then
   '     oDoc.InvoiceLines.SortInvoiceLines enSequence, optASC
        strSrtSeq = "SeqNum"
    End If
        
    If optSetSeqDef = 1 Then
        SaveSetting App.EXEName, "PrintSettings", "TransferSequenceField", strSrtSeq
        SaveSetting App.EXEName, "PrintSettings", "TransferSequenceSeq", IIf(optASC, "ASCEND", "DESCEND")
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_TFR.SortDetailLines"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_TFR.SortDetailLines"
End Sub
Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass

    SortDetailLines

    If oDOC.ExportToXML(enPrint, "", "", CInt(txtQty), True) = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Cannot print document, possibly no document has been set up for this workstation." & vbCrLf & "Try setting a document up using the configuration form.", vbInformation, "Can't print"
    End If
    Unload Me
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_TFR.cmdPrint_Click"
End Sub
Private Sub cmdPreview_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    cmdPreview.Enabled = False
    SortDetailLines
    If oDOC.ExportToXML(enView, "", "", CInt(txtQty)) = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
    End If
    Unload Me
    Screen.MousePointer = vbDefault

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_TFR.cmdPreview_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_TFR.cmdPreview_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub Command1_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_TFR.Command1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdExportToSpreadsheet_Click()
    On Error GoTo errHandler
Dim sFilename As String
    Screen.MousePointer = vbHourglass
    cmdPreview.Enabled = False
    SortDetailLines
    If oDOC.ExportToSpreadsheet(False, sFilename) = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
    End If
    If MsgBox("Spreadsheet file saved in: " & sFilename & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
        OpenFileWithApplication sFilename, enExcel
    End If
    Screen.MousePointer = vbDefault
    Unload Me

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_TFR.cmdExportToSpreadsheet_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim strInvSeqField As String
Dim strInvSeq As String

    flgLoading = True
    strInvSeqField = GetSetting(App.EXEName, "PrintSettings", "InvoiceSequenceField", "Title")
    strInvSeq = GetSetting(App.EXEName, "PrintSettings", "InvoiceSequenceSeq", "Title")
    Select Case strInvSeqField
    Case "Title"
        optTitle = True
    Case "Author"
        optAuthor = True
    Case "Code"
        optCode = True
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
'    ErrorIn "frmPrintingOptions_TFR.Form_Load"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_TFR.Form_Load", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If Not IsNumeric(txtQty) Then Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_TFR.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
