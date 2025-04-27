VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPSDetails 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Export catalogue"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6675
   StartUpPosition =   1  'CenterOwner
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
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   690
      Width           =   3975
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
      Left            =   1665
      TabIndex        =   2
      Top             =   1095
      Width           =   2175
   End
   Begin VB.TextBox txtCatalogues 
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
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   285
      Width           =   3975
   End
   Begin VB.TextBox txtAdj 
      Alignment       =   2  'Center
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
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmPSDetails.frx":0000
      Top             =   2190
      Width           =   630
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      Caption         =   "Export to"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   870
      Left            =   1665
      TabIndex        =   6
      Top             =   2685
      Width           =   3945
      Begin VB.OptionButton optBiblio 
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3CB&
         Caption         =   "Biblio"
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
         Height          =   345
         Left            =   2265
         TabIndex        =   13
         Top             =   330
         Width           =   1005
      End
      Begin VB.OptionButton optABE 
         Appearance      =   0  'Flat
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
         Height          =   345
         Left            =   660
         TabIndex        =   12
         Top             =   330
         Value           =   -1  'True
         Width           =   1035
      End
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
      Height          =   585
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1515
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5955
      Top             =   2925
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3735
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
      Height          =   465
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3735
      Width           =   1155
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "percent"
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
      Left            =   2370
      TabIndex        =   16
      Top             =   2235
      Width           =   1515
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
      Height          =   585
      Left            =   1680
      TabIndex        =   3
      Top             =   1515
      Width           =   3945
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Price adjustment"
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
      Left            =   45
      TabIndex        =   15
      Top             =   2235
      Width           =   1515
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Catalogues"
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
      Left            =   435
      TabIndex        =   14
      Top             =   330
      Width           =   1125
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Save to file"
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
      Left            =   435
      TabIndex        =   11
      Top             =   1575
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
      Left            =   435
      TabIndex        =   10
      Top             =   1125
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
      Left            =   435
      TabIndex        =   9
      Top             =   720
      Width           =   1125
   End
End
Attribute VB_Name = "frmPSDetails"
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
Dim strFilter As String

'Private Sub cboCatalogues_Click()
'   lngCatalogueID = tlCatalogues.Key(cboCatalogues.Text)
'End Sub

Private Sub cboOperators_Click()
    On Error GoTo errHandler
   lngOperatorID = tlOperators.Key(cboOperators.text)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPSDetails.cboOperators_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPSDetails.cmdCancel_Click", , EA_NORERAISE
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
    Set fs = Nothing
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPSDetails.cmdFile_Click", , EA_NORERAISE
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
Dim dblPriceFactor As Double

    strOpDesc = "Exporting items found in catalogue " & Me.txtCatalogues & vbCrLf & "formatted for " & strOPTo & vbCrLf & "and saving the data in " & strFilename
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
        Set oBat = New z_SQL
        Set oEx = New a_Export
        dblPriceFactor = CDbl(txtAdj)
        If strOPTo = "ABE" Then
            lngResult = oEx.Export_CAT_ABE(strFilename, strFilter, dblPriceFactor)
        ElseIf strOPTo = "Bibliofind" Then
            lngResult = oEx.Export_CAT_BIBLIO(strFilename, strFilter, dblPriceFactor)
        End If
 '       MsgBox "Done"
    End If
    Screen.MousePointer = vbDefault
    Set oEx = Nothing
    Set oBat = Nothing
    Set oOp = Nothing
    Me.cmdOK.Enabled = False
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPSDetails.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Set tlOperators = New z_TextList
    tlOperators.Load ltStaff
    LoadCombo cboOperators, tlOperators
    Set tlCatalogues = New z_TextList
    tlCatalogues.Load ltCatalogue
 '   LoadCombo cboCatalogues, tlCatalogues
    lngOperatorID = tlOperators.Key(cboOperators.text)
    strOPTo = "ABE"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPSDetails.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set tlOperators = Nothing
    Set tlCatalogues = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPSDetails.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub OptABE_Click()
    On Error GoTo errHandler
    strOPTo = "ABE"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPSDetails.OptABE_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub OptBiblio_Click()
    On Error GoTo errHandler
    strOPTo = "Bibliofind"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPSDetails.OptBiblio_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtAdj_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    
    If Not IsNumeric(txtAdj) Then
        MsgBox "Entry must be numeric. The price is multiplied by this factor" & vbCrLf & "Values betweem 0 and 1 reduce the prices, greater than 1 increase the prices."
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPSDetails.txtAdj_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtCatalogues_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim oSelection2 As z_Selection
Dim colCatsToInclude As Collection
Dim i As Long

    Set oSelection2 = New z_Selection
    Set colCatsToInclude = New Collection
    oSelection2.Parse txtCatalogues
    oSelection2.Load colCatsToInclude
    If colCatsToInclude.Count > 0 Then
        strFilter = "WHERE Arg_CATALID = "
        For i = 1 To colCatsToInclude.Count
            strFilter = strFilter & tlCatalogues.Key(colCatsToInclude(i))
            If colCatsToInclude.Count > i Then
                strFilter = strFilter & " OR Arg_CATALID = "
            End If
        Next
    End If
    Me.cmdOK.Enabled = (colCatsToInclude.Count > 0)
    Set oSelection2 = Nothing
    Set colCatsToInclude = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPSDetails.txtCatalogues_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
