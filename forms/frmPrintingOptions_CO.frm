VERSION 5.00
Begin VB.Form frmPrintingOptions_CO 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Order Print"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7305
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture 
      BackColor       =   &H00D3D3CB&
      Height          =   435
      Index           =   1
      Left            =   3675
      ScaleHeight     =   375
      ScaleWidth      =   2850
      TabIndex        =   15
      Top             =   2205
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   45
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture 
      BackColor       =   &H00D3D3CB&
      Height          =   1500
      Index           =   0
      Left            =   3615
      ScaleHeight     =   1440
      ScaleWidth      =   2910
      TabIndex        =   10
      Top             =   330
      Width           =   2970
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
         TabIndex        =   18
         Top             =   1155
         Width           =   2220
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
      Height          =   1875
      Left            =   3510
      TabIndex        =   9
      Top             =   60
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
      Left            =   3495
      TabIndex        =   8
      Top             =   1950
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
      Left            =   1770
      Picture         =   "frmPrintingOptions_CO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1695
      Width           =   1410
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
      Left            =   1035
      Picture         =   "frmPrintingOptions_CO.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   1410
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
      Left            =   330
      Picture         =   "frmPrintingOptions_CO.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1695
      Width           =   1410
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
      Left            =   1470
      TabIndex        =   3
      Text            =   "1"
      Top             =   480
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
      Left            =   4545
      TabIndex        =   2
      Top             =   2985
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
      Left            =   855
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4170
      Visible         =   0   'False
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
      Left            =   45
      TabIndex        =   19
      Top             =   3345
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
      Left            =   930
      TabIndex        =   4
      Top             =   210
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
      Left            =   870
      TabIndex        =   1
      Top             =   3900
      Visible         =   0   'False
      Width           =   1800
   End
End
Attribute VB_Name = "frmPrintingOptions_CO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flgLoading As Boolean
Dim oCO As a_CO

Public Sub ComponentObject(pCO As a_CO)
    On Error GoTo errHandler
    Set oCO = pCO
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_CO.ComponentObject(pCO)", pCO
End Sub

Private Sub cmdExportToSpreadsheet_Click()
Dim sFilename As String
    Screen.MousePointer = vbHourglass
    cmdPreview.Enabled = False
    SortDetailLines
    If oCO.ExportToSpreadsheet(False, sFilename) = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
    End If
    If MsgBox("Spreadsheet file saved in: " & sFilename & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
        OpenFileWithApplication sFilename, enExcel
    End If
    Screen.MousePointer = vbDefault
    Unload Me

End Sub

'Private Sub cboCurr_Click()
'    Set oCurrentForeignCurrency = oPC.Configuration.Currencies.FindByDescription(cboCurr)
'    oCO.BeginEdit
'    oCO.CurrencyID = oCurrentForeignCurrency.ID
'    oCO.ApplyEdit
'End Sub

Private Sub cmdPreview_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    cmdPreview.Enabled = False
    If oPC.GetProperty("UseXMLPrintingForCO") = "TRUE" Then
        SortDetailLines
        If Not oCO.ExportToXML("", enView, , , CInt(txtQty)) Then
            Screen.MousePointer = vbDefault
            MsgBox "Cannot print document, possibly no document has been set up for this workstation." & vbCrLf & "Try setting a document up using the configuration form.", vbInformation, "Can't print"
        End If
    Else
        Dim frm As frmPrintPreview
        Set frm = New frmPrintPreview
        frm.Caption = "Preview " & oCO.DOCCode
        frm.component oCO.Print_Display
        Screen.MousePointer = vbDefault
        frm.Show vbModal
    End If
    Unload Me
    Screen.MousePointer = vbDefault
    
    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_CO.cmdPreview_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    cmdPrint.Enabled = False
    If oPC.GetProperty("UseXMLPrintingForCO") = "TRUE" Then
        SortDetailLines
        If Not oCO.ExportToXML("", enPrint, , , CInt(txtQty), True) Then
            Screen.MousePointer = vbDefault
            MsgBox "Cannot print document, possibly no document has been set up for this workstation." & vbCrLf & "Try setting a document up using the configuration form.", vbInformation, "Can't print"
        End If
    Else
        If Not oCO.PrintCO Then
            Screen.MousePointer = vbDefault
            MsgBox "Cannot print document, possibly no document has been set up for this workstation." & vbCrLf & "Try setting a document up using the configuration form.", vbInformation, "Can't print"
        End If
    End If
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_CO.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub SortDetailLines()
    On Error GoTo errHandler
Dim strSrtSeq As String
    If optTitle Then
        oCO.COLines.SortInvoiceLines enTitle, optASC
        strSrtSeq = "Title"
    ElseIf optAuthor Then
        oCO.COLines.SortInvoiceLines enAuthor, optASC
        strSrtSeq = "Author"
    ElseIf optCode Then
        oCO.COLines.SortInvoiceLines enCode, optASC
        strSrtSeq = "Code"
    ElseIf optRef Then
        oCO.COLines.SortInvoiceLines enRef, optASC
        strSrtSeq = "Ref"
    ElseIf optSeq Then
        oCO.COLines.SortInvoiceLines enSequence, optASC
        strSrtSeq = "SeqNum"
    End If
        
    If optSetSeqDef = 1 Then
        SaveSetting App.EXEName, "PrintSettings", "COSequenceField", strSrtSeq
        SaveSetting App.EXEName, "PrintSettings", "COSequenceSeq", IIf(optASC, "ASCEND", "DESCEND")
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPrintingOptions_CO.SortDetailLines"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_CO.SortDetailLines"
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim strSeqField As String
Dim strSeq As String

    flgLoading = True
    strSeqField = GetSetting(App.EXEName, "PrintSettings", "COSequenceField", "Title")
    strSeq = GetSetting(App.EXEName, "PrintSettings", "COSequenceSeq", "Title")
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

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintingOptions_CO.Form_Load", , EA_NORERAISE
    HandleError
End Sub
