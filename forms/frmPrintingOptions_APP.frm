VERSION 5.00
Begin VB.Form frmPrintingOptions_APP 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Appro Print"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7260
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture 
      BackColor       =   &H00D3D3CB&
      Height          =   435
      Index           =   1
      Left            =   3615
      ScaleHeight     =   375
      ScaleWidth      =   2850
      TabIndex        =   11
      Top             =   2055
      Width           =   2910
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
         TabIndex        =   13
         Top             =   45
         Width           =   1320
      End
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
         TabIndex        =   12
         Top             =   45
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture 
      BackColor       =   &H00D3D3CB&
      Height          =   1320
      Index           =   0
      Left            =   3555
      ScaleHeight     =   1260
      ScaleWidth      =   2910
      TabIndex        =   8
      Top             =   300
      Width           =   2970
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
         TabIndex        =   10
         Top             =   600
         Width           =   900
      End
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
         TabIndex        =   9
         Top             =   30
         Width           =   900
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
      Height          =   1725
      Left            =   3450
      TabIndex        =   7
      Top             =   0
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
      Left            =   3435
      TabIndex        =   6
      Top             =   1755
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
      Left            =   1800
      Picture         =   "frmPrintingOptions_APP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2070
      Width           =   1515
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Print"
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
      Left            =   975
      Picture         =   "frmPrintingOptions_APP.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1335
      Width           =   1515
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
      Left            =   255
      Picture         =   "frmPrintingOptions_APP.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2070
      Width           =   1515
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
      Left            =   3945
      TabIndex        =   1
      Top             =   2760
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
      Left            =   1440
      TabIndex        =   0
      Text            =   "1"
      Top             =   675
      Width           =   660
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
      Left            =   60
      TabIndex        =   14
      Top             =   3420
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
      Left            =   810
      TabIndex        =   5
      Top             =   345
      Width           =   1800
   End
End
Attribute VB_Name = "frmPrintingOptions_APP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim oCurrentForeignCurrency As a_Currency
Dim oAPP As a_APP
Dim flgLoading As Boolean


Public Sub ComponentObject(pCO As a_APP)
    On Error GoTo errHandler
Dim oDC As a_DocumentControl
    Set oAPP = pCO
    Set oDC = oPC.Configuration.DocumentControls.FindDC(oAPP.constDOCCODE)
    If Not oDC Is Nothing Then
        txtQty = CStr(oPC.Configuration.DocumentControls.FindDC(oAPP.constDOCCODE).QtyCopies)
    Else
        txtQty = "1"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_APP.ComponentObject(pCO)", pCO
End Sub

Private Sub cmdExportToSpreadsheet_Click()
Dim sFilename As String
    Screen.MousePointer = vbHourglass
    cmdPreview.Enabled = False
    SortDetailLines
    If oAPP.ExportToSpreadsheet(False, sFilename) = False Then
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
Dim strSeqField As String
Dim strSeq As String

    flgLoading = True
    strSeqField = GetSetting(App.EXEName, "PrintSettings", "ApproSequenceField", "Title")
    strSeq = GetSetting(App.EXEName, "PrintSettings", "ApproSequenceSeq", "Title")
    Select Case strSeqField
    Case "Title"
        optTitle = True
    Case "Code"
        optCode = True
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
'    ErrorIn "frmPrintingOptions_App.Form_Load"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_APP.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    cmdPrint.Enabled = False

    SortDetailLines
    
    If oPC.GetProperty("UseXMLPrintingForAPP") = "TRUE" Then
        If Not oAPP.ExportToXML(enPrint, , , , True) Then
            MsgBox "Cannot print document, possibly no document has been set up for this workstation." & vbCrLf & "Try setting a document up using the configuration form.", vbInformation, "Can't print"
        End If
    Else
        If Not oAPP.PrintAPP(txtQty) Then
            MsgBox "Cannot print document, possibly no document has been set up for this workstation." & vbCrLf & "Try setting a document up using the configuration form.", vbInformation, "Can't print"
        End If
    End If
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_APP.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub SortDetailLines()
    On Error GoTo errHandler
Dim strSrtSeq As String
    If optTitle Then
        oAPP.ApproLines.SortLines enTitle, optASC
        strSrtSeq = "Title"
    ElseIf optCode Then
        oAPP.ApproLines.SortLines enCode, optASC
        strSrtSeq = "Code"
    End If
        
    If optSetSeqDef = 1 Then
        SaveSetting App.EXEName, "PrintSettings", "ApproSequenceField", strSrtSeq
        SaveSetting App.EXEName, "PrintSettings", "ApproSequenceSeq", IIf(optASC, "ASCEND", "DESCEND")
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_App.SortDetailLines"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_APP.SortDetailLines"
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo errHandler
Dim frm As frmPrintPreview

    Screen.MousePointer = vbHourglass
    cmdPreview.Enabled = False
    SortDetailLines
    
    If oPC.GetProperty("UseXMLPrintingForAPP") = "TRUE" Then
        If Not oAPP.ExportToXML(enView) Then
            MsgBox "Cannot print document, possibly no document has been set up for this workstation." & vbCrLf & "Try setting a document up using the configuration form.", vbInformation, "Can't print"
        End If
    Else
        Set frm = New frmPrintPreview
        frm.Caption = "Preview " & oAPP.DOCCode
        frm.component oAPP.PrintAPP
        frm.Show vbModal
    End If
    Unload Me
    Screen.MousePointer = vbDefault
    
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_Inv.cmdPreview_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_APP.cmdPreview_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If Not IsNumeric(txtQty) Then Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_APP.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

