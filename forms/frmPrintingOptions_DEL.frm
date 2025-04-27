VERSION 5.00
Begin VB.Form frmPrintingOptions_DEL 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Print Goods received note"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture 
      BackColor       =   &H00D3D3CB&
      Height          =   435
      Index           =   1
      Left            =   3660
      ScaleHeight     =   375
      ScaleWidth      =   2850
      TabIndex        =   17
      Top             =   2085
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   45
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture 
      BackColor       =   &H00D3D3CB&
      Height          =   1320
      Index           =   0
      Left            =   3600
      ScaleHeight     =   1260
      ScaleWidth      =   2910
      TabIndex        =   12
      Top             =   330
      Width           =   2970
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
         TabIndex        =   16
         Top             =   885
         Width           =   2220
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
         TabIndex        =   14
         Top             =   315
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
         TabIndex        =   13
         Top             =   30
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
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
      Left            =   3495
      TabIndex        =   11
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
      Left            =   3480
      TabIndex        =   10
      Top             =   1785
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
      Left            =   1845
      Picture         =   "frmPrintingOptions_DEL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1875
      Width           =   1440
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
      Left            =   1080
      Picture         =   "frmPrintingOptions_DEL.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1230
      Width           =   1440
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
      Left            =   375
      Picture         =   "frmPrintingOptions_DEL.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1875
      Width           =   1440
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
      TabIndex        =   5
      Text            =   "1"
      Top             =   675
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
      Left            =   4050
      TabIndex        =   4
      Top             =   2850
      Width           =   2445
   End
   Begin VB.CommandButton cmdReserved 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Reserve list"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2835
      Width           =   2040
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Select currency"
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
      Height          =   1050
      Left            =   15
      TabIndex        =   0
      Top             =   3750
      Width           =   3510
      Begin VB.OptionButton optL 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Local currency"
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
         Height          =   405
         Left            =   120
         TabIndex        =   2
         Top             =   435
         Width           =   1605
      End
      Begin VB.OptionButton optF 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Local currency"
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
         Height          =   405
         Left            =   1815
         TabIndex        =   1
         Top             =   465
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
      Left            =   75
      TabIndex        =   20
      Top             =   3405
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
      Left            =   1005
      TabIndex        =   6
      Top             =   360
      Width           =   1800
   End
End
Attribute VB_Name = "frmPrintingOptions_DEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oDel As a_Delivery
Dim flgLoading As Boolean
Dim strSeqField As String
Dim strSeq As String
Dim strSrtSeq As String

Public Sub ComponentObject(pPO As a_Delivery)
    On Error GoTo errHandler
    Set oDel = pPO
    optL.Caption = oPC.Configuration.DefaultCurrency.Description
    optF.Visible = False
    optL.Value = True
    If oDel.ISForeignCurrency Then
        optF.Caption = oDel.CaptureCurrency.Description
        optF.Value = True
        optF.Enabled = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_DEL.ComponentObject(pPO)", pPO
End Sub


Private Sub cmdExportToSpreadsheet_Click()
Dim sFilename As String
    Screen.MousePointer = vbHourglass
    cmdPreview.Enabled = False
    SortDetailLines
    If oDel.ExportToSpreadsheet(False, sFilename) = False Then
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
Dim frm As frmPrintPreview
    Screen.MousePointer = vbHourglass
    cmdPreview.Enabled = False
    SortDetailLines
    If oPC.GetProperty("UseXMLPrintingForGRN") = "TRUE" Then
        If Not oDel.ExportToXML(enView) Then
            Screen.MousePointer = vbDefault
            MsgBox "Cannot print document, possibly no document has been set up for this workstation." & vbCrLf & "Try setting a document up using the configuration form.", vbInformation, "Can't print"
        End If
        Unload Me
    Else
        Set frm = New frmPrintPreview
        frm.Caption = "Preview " & oDel.DOCCode
        frm.component oDel.Print_Display(Me.optL = False)
        Screen.MousePointer = vbDefault
        frm.Show vbModal
        Unload Me
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_DEL.cmdPreview_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    cmdPrint.Enabled = False
    SortDetailLines
    If oPC.GetProperty("UseXMLPrintingForGRN") = "TRUE" Then
        If Not oDel.ExportToXML(enPrint, , , CInt(txtQty), True) Then
            Screen.MousePointer = vbDefault
            MsgBox "Cannot print document, possibly no document has been set up for this workstation." & vbCrLf & "Try setting a document up using the configuration form.", vbInformation, "Can't print"
        End If
    Else
        If Not oDel.PrintDEL(optL = False) Then
            Screen.MousePointer = vbDefault
            MsgBox "Cannot print document, possibly no document has been set up for this workstation." & vbCrLf & "Try setting a document up using the configuration form.", vbInformation, "Can't print"
        End If
    End If
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_DEL.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub PrintProductsAwaitedBYCOs()
    On Error GoTo errHandler
Dim oC As chex_COLAllocation
Dim ar As New arCOLSFulfilled
    Set oC = New chex_COLAllocation
    oC.Load oDel.TRID, True
    ar.component oC
    ar.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_DEL.PrintProductsAwaitedBYCOs"
End Sub

Private Sub cmdReserved_Click()
    On Error GoTo errHandler
    PrintProductsAwaitedBYCOs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_DEL.cmdReserved_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub SortDetailLines()
    On Error GoTo errHandler
Dim strSrtSeq As String
    If optTitle Then
        oDel.DeliveryLines.SortLines enTitle, optASC
        strSrtSeq = "Title"
    ElseIf optAuthor Then
        oDel.DeliveryLines.SortLines enAuthor, optASC
        strSrtSeq = "Author"
    ElseIf optCode Then
        oDel.DeliveryLines.SortLines enCode, optASC
        strSrtSeq = "Code"
    ElseIf optSeq Then
        oDel.DeliveryLines.SortLines enSequence, optASC
        strSrtSeq = "SeqNum"
    End If
        
    If optSetSeqDef = 1 Then
        SaveSetting "PBKS", "PrintSettings", "DELSequenceField", strSrtSeq
        SaveSetting "PBKS", "PrintSettings", "DELSequenceSeq", IIf(optASC, "ASCEND", "DESCEND")
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_DEL.SortDetailLines"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_DEL.SortDetailLines"
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    strSeqField = GetSetting("PBKS", "PrintSettings", "DELSequenceField", "")
    strSeq = GetSetting("PBKS", "PrintSettings", "DELSequenceSeq", "")
    Select Case strSeqField
    Case "Title"
        optTitle = True
    Case "Author"
        optAuthor = True
    Case "Code"
        optCode = True
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
  '  txtQty = oPC.configuration
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_DEL.Form_Load", , EA_NORERAISE
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
'    ErrorIn "frmPrintingOptions_DEL.txtQty_Validate(Cancel)", Cancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_DEL.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

