VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_Step_2 
   BackColor       =   &H00E8E8DD&
   Caption         =   "Step 2 - Import scanned files"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNext_To_3 
      BackColor       =   &H00D8D9C4&
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5865
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4635
      Width           =   840
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   165
      Left            =   1275
      TabIndex        =   2
      Top             =   4590
      Visible         =   0   'False
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdImportSimple 
      BackColor       =   &H00D8D9C4&
      Caption         =   "Import scanned files"
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
      Left            =   1755
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2835
      Width           =   3345
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   75
      Top             =   2940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.txt"
      DialogTitle     =   "Locate scanner files"
      MaxFileSize     =   30000
   End
   Begin VB.Label lblSharedFolder 
      BackStyle       =   0  'Transparent
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
      Height          =   240
      Left            =   900
      TabIndex        =   6
      Top             =   2070
      Width           =   5130
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Note. The files to be scanned should all be in the default folder "
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
      Height          =   360
      Left            =   540
      TabIndex        =   5
      Top             =   1650
      Width           =   5925
   End
   Begin VB.Label lblFilename 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   750
      Left            =   1275
      TabIndex        =   4
      Top             =   3705
      Width           =   4455
   End
   Begin VB.Label lbl_Step_1_Msg 
      BackStyle       =   0  'Transparent
      Caption         =   "Click button to select the files you want to import."
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
      Height          =   525
      Left            =   270
      TabIndex        =   1
      Top             =   1125
      Width           =   5325
   End
End
Attribute VB_Name = "frm_Step_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oSA As a_Stktke
Attribute oSA.VB_VarHelpID = -1

Dim strFilename As String


Public Sub Component(pSA As a_Stktke)
    Set oSA = pSA
End Sub
Private Sub cmdCancel_Click()

End Sub

Private Sub cmdContinue_Click()
    If MsgBox("Cancel current stock-take ( you can start a new one afterwards", vbQuestion + vbYesNo) = vbYes Then
        'Cancel
    End If
End Sub

Private Sub cmdImportSimple_Click()
    On Error GoTo errHandler

Dim iresult As Integer
Dim fs As New Scripting.FileSystemObject
Dim lngBadRecords As Long
Dim dteEffectiveDate As Date
Dim lngLastSAID As Long
Dim rs As adodb.Recordset
Dim lngFilecount As Long
Dim Testar() As String
Dim strErrorFilenames As String
Dim i As Integer

    If Not fs.FolderExists(oPC.SharedFolderRoot & "\Stocktke") Then
        fs.CreateFolder oPC.SharedFolderRoot & "\Stocktke"
    End If
    CD1.InitDir = oPC.SharedFolderRoot & "\Stocktke"
    CD1.FLAGS = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer
    CD1.CancelError = True
    CD1.Filter = "Text Files (*.txt)|*.txt"
    CD1.FileName = ""
    CD1.ShowOpen
    If CD1.FileName = "" Then
        MsgBox "You must specify an existing file name!", vbInformation, "Invalid filename"
    Else
        strFilename = CD1.FileName
    End If
    
    Testar = Split(strFilename, Chr(0))
    strErrorFilenames = ""
    If UBound(Testar) = 0 Then 'There is only one file
            If StripToAlphanumeric(fs.GetBaseName(strFilename)) <> fs.GetBaseName(strFilename) Then
                strErrorFilenames = fs.GetFileName(strFilename)
            End If
    Else
        For i = 1 To UBound(Testar)
            If StripToAlphanumeric(fs.GetBaseName(Testar(i))) <> fs.GetBaseName(Testar(i)) Then
                strErrorFilenames = strErrorFilenames & IIf(Len(strErrorFilenames) > 0, vbCrLf, "") & Testar(i)
            End If
        Next
    End If

    If strErrorFilenames > "" Then
        MsgBox "Some files have names containing invalid characters. Names should only contain alphabetic or digit characters." & vbCrLf & "Please fix names before continuing." & vbCrLf & "Invalid names follow:" & vbCrLf & strErrorFilenames, vbInformation, "Can't do this"
        Exit Sub
    End If
    
    Me.PB1.Visible = True
    
    Screen.MousePointer = vbHourglass
    Me.Refresh
    
    
    oSA.ImportSimple strFilename, lngBadRecords, lngFilecount 'creates a stocktake and imports the data into STOCKTAKE_WORKC consolidated by filename
    Me.lblFilename = "Checking for missing items"
    DoEvents
    Screen.MousePointer = vbHourglass
    oSA.PrepareMissingData
    
    lblFilename.Caption = ""
    lngLastSAID = oSA.TransactionID
    PB1.Visible = False
    DoEvents
    Screen.MousePointer = vbDefault
    MsgBox "Import complete: " & CStr(lngFilecount) & " files imported", vbOKOnly, "Status"

    Me.Refresh
EXIT_Handler:
    Me.PB1.Visible = False
    Exit Sub
errHandler:
    ErrPreserve
    If Err = 32755 Then   'Cancel selected in CD1
        Exit Sub
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frm_Step_2.cmdImportSimple_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oSA_ImportFile(pFilename As String)
    lblFilename.Caption = "Importing . . . " & pFilename
    DoEvents
End Sub
Private Sub oSA_LineCOuntChange(pCnt As Long)
    PB1.Value = pCnt
End Sub
Private Sub oSA_MaxImportRows(pMax As Long)
    PB1.Max = pMax
    PB1.Min = 0
    PB1.Value = 0
    PB1.Visible = True
End Sub
Private Sub oSA_FinishedImporting()
    PB1.Visible = False
    lblFilename.Caption = ""
    DoEvents
End Sub
Private Sub cmdToStep3_Click()
    Set frm3 = New frm_Step_3
    frm3.Show
    Unload Me
End Sub

Private Sub cmdNext_To_3_Click()
    Set frm3 = New frm_Step_3
    frm3.Component oSA
    frm3.Show
    Unload Me
End Sub


Private Sub Form_Load()
    lblSharedFolder.Caption = oPC.SharedFolderRoot & "\STOCKTKE)"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not UnloadMode = 1 Then
        If MsgBox("Do you want to close the stock-take application?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
            Cancel = True
        End If
    End If
End Sub
