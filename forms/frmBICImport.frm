VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBICImport 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Import BIC Codes from Bookdata CD"
   ClientHeight    =   3765
   ClientLeft      =   7755
   ClientTop       =   1080
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7200
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdUpdaterecords 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Update product records from Bookfind BIC Codes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2085
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2475
      Width           =   2880
   End
   Begin VB.CommandButton cmdFind 
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
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   405
      Width           =   585
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Import BIC codes"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1650
      Width           =   2865
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   630
      Top             =   1260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".RT"
      DialogTitle     =   "Find BIC codes file"
      FileName        =   "BIC"
      Filter          =   "Bookfind .RT file"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Find File containing BIC codes (usually in the same folder as the Papyrus database)"
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
      Height          =   600
      Left            =   660
      TabIndex        =   3
      Top             =   165
      Width           =   4335
   End
   Begin VB.Label lblBICSourceFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Nothing>"
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
      Height          =   375
      Left            =   645
      TabIndex        =   2
      Top             =   795
      Width           =   5715
   End
End
Attribute VB_Name = "frmBICImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strBICSource As String
Private Sub cmdFind_Click()
    On Error GoTo errHandler
Dim fs As New Scripting.FileSystemObject
    CD1.DefaultExt = ".csv"
    CD1.InitDir = oPC.SharedFolderRoot & "\Data\"
    CD1.FLAGS = cdlOFNFileMustExist Or cdlOFNReadOnly
    CD1.ShowOpen
    If CD1.FileName = "" Then
        MsgBox "You must specify a file name!", vbInformation, "Invalid filename"
    Else
        strBICSource = CD1.FileName
        lblBICSourceFile.Caption = strBICSource
        Me.cmdOK.Enabled = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBICImport.cmdFind_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    If MsgBox("Do you want to refresh all the BIC codes from Bookdata?", vbQuestion + vbYesNo, "Confirm") = -vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    DropAndCreateBICTable
    ImportBIC (strBICSource)
    oPC.Configuration.Reload
    Screen.MousePointer = vbDefault
    MsgBox "Import complete", vbExclamation, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBICImport.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdUpdaterecords_Click()
    On Error GoTo errHandler
'Dim oBF As a_BookFind
    If MsgBox("Do you want to update all the book records' BIC codes from Bookdata?", vbQuestion + vbYesNo, "Confirm") = -vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Set oBF = New a_BookFind
    oBF.UpdateBICCodes
    Set oBF = Nothing
    Screen.MousePointer = vbDefault
    MsgBox "Update complete", vbExclamation, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBICImport.cmdUpdaterecords_Click", , EA_NORERAISE
    HandleError
End Sub
