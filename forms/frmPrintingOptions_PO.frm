VERSION 5.00
Begin VB.Form frmPrintingOptions_PO 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Purchase order print"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7560
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture 
      BackColor       =   &H00D3D3CB&
      Height          =   435
      Index           =   1
      Left            =   3765
      ScaleHeight     =   375
      ScaleWidth      =   2850
      TabIndex        =   18
      Top             =   2220
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   45
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture 
      BackColor       =   &H00D3D3CB&
      Height          =   1500
      Index           =   0
      Left            =   3705
      ScaleHeight     =   1440
      ScaleWidth      =   2910
      TabIndex        =   12
      Top             =   345
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
      Left            =   3600
      TabIndex        =   11
      Top             =   75
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
      Left            =   3585
      TabIndex        =   10
      Top             =   1965
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
      Left            =   1935
      Picture         =   "frmPrintingOptions_PO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1695
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
      Left            =   1245
      Picture         =   "frmPrintingOptions_PO.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1035
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
      Left            =   525
      Picture         =   "frmPrintingOptions_PO.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1695
      Width           =   1380
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
      Left            =   4215
      TabIndex        =   6
      Top             =   2970
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
      Left            =   1650
      TabIndex        =   4
      Text            =   "1"
      Top             =   480
      Width           =   660
   End
   Begin VB.CommandButton cmdEDI 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Send"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   -270
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2610
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame currFrame 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Select currency"
      ForeColor       =   &H8000000D&
      Height          =   915
      Left            =   990
      TabIndex        =   0
      Top             =   2430
      Width           =   1935
      Begin VB.OptionButton optF 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Local currency"
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   180
         TabIndex        =   2
         Top             =   420
         Width           =   1605
      End
      Begin VB.OptionButton optL 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Local currency"
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   165
         TabIndex        =   1
         Top             =   180
         Width           =   1605
      End
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
      Left            =   30
      TabIndex        =   21
      Top             =   3555
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
      Left            =   1035
      TabIndex        =   5
      Top             =   165
      Width           =   1800
   End
End
Attribute VB_Name = "frmPrintingOptions_PO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oPO As a_PO
Dim flgLoading As Boolean
Dim strSeqField As String
Dim strSeq As String
Dim strSrtSeq As String

Public Sub ComponentObject(pPO As a_PO, Optional pPrintOrSend As enTransmitType)
    On Error GoTo errHandler
    Set oPO = pPO
    optL.Caption = oPC.Configuration.DefaultCurrency.Description
    optF.Visible = False
    optL.Value = True
    Me.currFrame.Visible = oPO.ISForeignCurrency  'Not (oPO.Supplier.DefaultCurrency Is oPC.Configuration.DefaultCurrency)
    If oPO.ISForeignCurrency Then
        optF.Caption = oPO.CaptureCurrency.Description
        optF.Visible = True
        optF.Value = True
        optF.Enabled = True
    End If
    If pPrintOrSend = enEDI Then
'        cmdPrint.Visible = False
'        cmdEDI.Top = 1800
'        cmdEDI.Left = 930
    Else
'        cmdPrint.Top = 1800
'        Me.cmdPreview.Top = 1800
        cmdEDI.Visible = False
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_PO.ComponentObject(pPO,pPrintOrSend)", Array(pPO, pPrintOrSend)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_PO.ComponentObject(pPO,pPrintOrSend)", Array(pPO, pPrintOrSend)
End Sub


Private Sub cmdEDI_Click()
    On Error GoTo errHandler
    oPO.GenerateSAANAMsg
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_PO.cmdEDI_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_PO.cmdEDI_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdExportToSpreadsheet_Click()
Dim sFilename As String
    Screen.MousePointer = vbHourglass
    cmdPreview.Enabled = False
    SortDetailLines
    If oPO.ExportToSpreadsheet(False, sFilename) = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
    End If
    If MsgBox("Spreadsheet file saved in: " & sFilename & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
        OpenFileWithApplication sFilename, enExcel
    End If
    Screen.MousePointer = vbDefault
    Unload Me

End Sub

Private Sub cmdPreview_Click()
    On Error GoTo errHandler
Dim Res
    Screen.MousePointer = vbHourglass
    cmdPreview.Enabled = False
    SortDetailLines
    If oPO.ExportToXML(Me.optF, enView, "", , , CInt(txtQty)) = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Cannot view. Probably a purchase order document has not been configured for this workstation." & vbCrLf & "Use the menu Settings>Configuration and the documents tab to correct this."
    End If
    Unload Me
    Screen.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_PO.cmdPreview_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_PO.cmdPreview_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim Res
    Screen.MousePointer = vbHourglass
    cmdPrint.Enabled = False
    SortDetailLines

    If oPO.ExportToXML(Me.optF, enPrint, "", "", , CInt(txtQty), True) = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Cannot print. Probably a purchase order document has not been configured for this workstation." & vbCrLf & "Use the menu Settings>Configuration and the documents tab to correct this."
    End If
    Unload Me
    Screen.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_PO.cmdPrint_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_PO.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    strSeqField = GetSetting("PBKS", "PrintSettings", "POSequenceField", "")
    strSeq = GetSetting("PBKS", "PrintSettings", "POSequenceSeq", "")
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
'    ErrorIn "frmPrintingOptions_PO.Form_Load"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_PO.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim lngTmp As Long
    If Not ConvertToLng(txtQty, lngTmp) Then
        Cancel = True
    End If
  '  If Not IsNumeric(txtQty) Then Cancel = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_PO.txtQty_Validate(Cancel)", Cancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_PO.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub SortDetailLines()
    On Error GoTo errHandler
Dim strSrtSeq As String
    If optTitle Then
        oPO.POLines.SortPOLines enTitle, optASC
        strSrtSeq = "Title"
    ElseIf optAuthor Then
        oPO.POLines.SortPOLines enAuthor, optASC
        strSrtSeq = "Author"
    ElseIf optCode Then
        oPO.POLines.SortPOLines enCode, optASC
        strSrtSeq = "Code"
    ElseIf optRef Then
        oPO.POLines.SortPOLines enRef, optASC
        strSrtSeq = "Ref"
    ElseIf optSeq Then
        oPO.POLines.SortPOLines enSequence, optASC
        strSrtSeq = "SeqNum"
    End If
        
    If optSetSeqDef = 1 Then
        SaveSetting "PBKS", "PrintSettings", "POSequenceField", strSrtSeq
        SaveSetting "PBKS", "PrintSettings", "POSequenceSeq", IIf(optASC, "ASCEND", "DESCEND")
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_PO.SortDetailLines"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_PO.SortDetailLines"
End Sub

