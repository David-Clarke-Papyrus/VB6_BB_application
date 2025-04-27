VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCapturedSince 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Books captured since"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSince 
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
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "frmCapturedSince.frx":0000
      Top             =   420
      Width           =   2115
   End
   Begin VB.CommandButton cmdFile 
      BackColor       =   &H00C4BCA4&
      Caption         =   "···"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5580
      MaskColor       =   &H00CDCFAD&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4200
      Top             =   4350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Export to . . ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1560
      TabIndex        =   2
      Top             =   2100
      Width           =   2985
      Begin VB.OptionButton optBiblio 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Bibliofind"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2580
      MaskColor       =   &H00CDCFAD&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4290
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1380
      MaskColor       =   &H00CDCFAD&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4290
      Width           =   1155
   End
   Begin VB.ComboBox cboOperators 
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
      Left            =   1560
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1530
      Width           =   2175
   End
   Begin VB.TextBox txtDescription 
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
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmCapturedSince.frx":0006
      Top             =   990
      Width           =   4155
   End
   Begin VB.Label lblFilename 
      BackColor       =   &H00E8E8E8&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1560
      TabIndex        =   12
      Top             =   3360
      Width           =   3945
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Captured since"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   90
      TabIndex        =   11
      Top             =   450
      Width           =   1365
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Save to file:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   300
      TabIndex        =   10
      Top             =   3420
      Width           =   1125
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Operator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   330
      TabIndex        =   7
      Top             =   1560
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   330
      TabIndex        =   6
      Top             =   1020
      Width           =   1125
   End
End
Attribute VB_Name = "frmCapturedSince"
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
   lngOperatorID = tlOperators.Key(cboOperators.text)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFile_Click()
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
        Me.lblFilename = strFilename
        Me.cmdOK.Enabled = True
    End If
End Sub

Private Sub cmdOK_Click()
Dim ctrl As OptionButton
Dim oOp As a_Operation
Dim strOpDesc As String
Dim lngResult As Long
Dim oBat As z_SQL
Dim oEx As a_Export

    strOpDesc = "Exporting items captured since " & Me.txtSince & vbCrLf & "formatted for " & strOPTo & vbCrLf & "and saving the data in " & strFilename
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
        lngResult = oBat.RunProc("AllCopiesCapturedsince", Array(CDate(Me.txtSince)), "Exporting . . .")
        Set oEx = New a_Export
        If strOPTo = "ABE" Then
    '        lngResult = oEx.ExportToABE(strFilename)
        ElseIf strOPTo = "Bibliofind" Then
    '        lngResult = oEx.ExportToBib(strFilename)
        End If
        MsgBox "Done"
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
End Sub

Private Sub Form_Load()
    Set tlOperators = New z_TextList
    tlOperators.Load ltStaff
    LoadCombo cboOperators, tlOperators
 '   Set tlCatalogues = New z_TextList
 '   tlCatalogues.Load "Catalogues_tl"
 '   LoadCombo cboCatalogues, tlCatalogues
    lngOperatorID = tlOperators.Key(cboOperators.text)
    Me.optABE = True
    strOPTo = "ABE"
End Sub

Private Sub OptABE_Click()
    strOPTo = "ABE"
End Sub

Private Sub OptBiblio_Click()
    strOPTo = "Bibliofind"
End Sub

Private Sub txtSince_Validate(Cancel As Boolean)
    If Not IsDate(txtSince) Then
        MsgBox "Invalid date"
        Cancel = True
    End If
End Sub
