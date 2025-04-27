VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEntire 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Create wants list from all books on database"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSince 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1500
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   1065
      Width           =   2175
   End
   Begin VB.CommandButton cmdFile 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5475
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   585
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   135
      Top             =   3825
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Export to . . ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1500
      TabIndex        =   2
      Top             =   1530
      Width           =   2985
      Begin VB.OptionButton optBiblio 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Bibliofind"
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
         Height          =   255
         Left            =   1410
         TabIndex        =   9
         Top             =   480
         Width           =   1485
      End
      Begin VB.OptionButton optABE 
         BackColor       =   &H00D3D3CB&
         Caption         =   "ABE"
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
         Height          =   255
         Left            =   330
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3585
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4305
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3585
      Width           =   1155
   End
   Begin VB.ComboBox cboOperators 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1500
      TabIndex        =   1
      Top             =   660
      Width           =   2175
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1500
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   4155
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Captured since"
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
      Left            =   0
      TabIndex        =   13
      Top             =   1095
      Width           =   1395
   End
   Begin VB.Label lblFilename 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   1500
      TabIndex        =   11
      Top             =   2760
      Width           =   3945
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Save to file:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   2820
      Width           =   1125
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operator"
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
      Left            =   270
      TabIndex        =   7
      Top             =   690
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   270
      TabIndex        =   6
      Top             =   285
      Width           =   1125
   End
End
Attribute VB_Name = "frmEntire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tlOperators As z_TextList
Dim tlCatalogues As z_TextList
Dim strFilename As String
Dim lngOperatorID As Long
Dim lngCatalogueID As Long
Dim strOPTo As String


Private Sub cboOperators_Click()
    On Error GoTo errHandler
   lngOperatorID = tlOperators.key(cboOperators.Text)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEntire.cboOperators_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEntire.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFile_Click()
    On Error GoTo errHandler
Dim fs As Scripting.FileSystemObject
    Set fs = New Scripting.FileSystemObject
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\Export files") Then
        fs.CreateFolder oPC.SharedFolderRoot & "\Export files"
    End If
    CD1.InitDir = oPC.SharedFolderRoot & "\Export files"
    CD1.FLAGS = cdlOFNCreatePrompt
    CD1.ShowSave
    If fs.FileExists(CD1.FileName) Or (CD1.FileName = "") Then
        MsgBox "You must specify a new file name!", vbInformation, "Invalid filename"
    Else
        strFilename = CD1.FileName
        If Not InStr(1, strFilename, ".", vbTextCompare) Then
            strFilename = strFilename & IIf(Right(strFilename, 4) = ".txt", "", ".txt")
        End If
        Me.lblFilename = strFilename
        Me.cmdOK.Enabled = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEntire.cmdFile_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim ctrl As OptionButton
Dim oOp As a_Operation
Dim strOpDesc As String
Dim lngResult As Long
Dim oBat As z_SQL
Dim oEx As a_Export
Dim strStatus As String

    strOpDesc = "Exporting entire database formatted for wants list for " & strOPTo & vbCrLf & "and saving the data in " & strFilename
    If MsgBox("Confirm:" & vbCrLf & strOpDesc, vbOKCancel Or vbInformation, "Confirm") = vbOK Then
        Screen.MousePointer = vbHourglass
        Set oOp = New a_Operation
        oOp.BeginEdit
        oOp.StartedAt = Now()
        oOp.TypeID = Export
        oOp.NominalDate = Date
        oOp.OperatorID = lngOperatorID
        oOp.Fullreport = strOpDesc
        oOp.ApplyEdit lngResult
        Set oEx = New a_Export
        If strOPTo = "ABE" Then
            lngResult = oEx.ExportWantsListToABE(strFilename, CDate(Me.txtSince))
        ElseIf strOPTo = "Bibliofind" Then
            lngResult = oEx.ExportWantsListToBib(strFilename)
        End If
    End If
    Screen.MousePointer = vbDefault
    Set oEx = Nothing
    Set oOp = Nothing
    Me.cmdOK.Enabled = False
    oPC.Configuration.BeginEdit
    oPC.Configuration.LastWantsExportDate = Date
    oPC.Configuration.ApplyEdit strStatus
    Me.txtSince = oPC.Configuration.LastWantsExportDate
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEntire.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Set tlOperators = New z_TextList
    tlOperators.Load ltStaff
    LoadCombo cboOperators, tlOperators
 '   Set tlCatalogues = New z_TextList
 '   tlCatalogues.Load "Catalogues_tl"
 '   LoadCombo cboCatalogues, tlCatalogues
    lngOperatorID = tlOperators.key(cboOperators.Text)
    Me.txtSince = Format(oPC.Configuration.LastWantsExportDate, "dd/mm/yyyy")
    Me.optABE = True
    strOPTo = "ABE"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEntire.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub OptABE_Click()
    On Error GoTo errHandler
    strOPTo = "ABE"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEntire.OptABE_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub OptBiblio_Click()
    On Error GoTo errHandler
    strOPTo = "Bibliofind"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEntire.OptBiblio_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtSince_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not IsDate(txtSince)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmEntire.txtSince_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
