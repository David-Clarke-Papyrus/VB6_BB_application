VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Catalogue production"
   ClientHeight    =   4665
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboCurr 
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
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   600
      Width           =   2205
   End
   Begin VB.ComboBox cboCats 
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
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   600
      Width           =   1635
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   645
      Left            =   0
      TabIndex        =   4
      Top             =   4020
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   1138
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   8811
            MinWidth        =   8820
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Generate catalogue"
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
      Height          =   615
      Left            =   2670
      TabIndex        =   2
      Top             =   2550
      Width           =   2865
   End
   Begin VB.CommandButton cmdFT 
      Caption         =   ">>>"
      Height          =   375
      Left            =   7410
      TabIndex        =   0
      Top             =   1890
      Width           =   585
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4830
      Top             =   1230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".DOC"
      DialogTitle     =   "Save to new file"
      Filter          =   "Word Document (*.Doc)|*.doc"
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   465
      Left            =   150
      TabIndex        =   5
      Top             =   3360
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   820
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label4 
      Caption         =   "Currency"
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
      Left            =   2070
      TabIndex        =   9
      Top             =   300
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Catalogue"
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
      Left            =   210
      TabIndex        =   7
      Top             =   300
      Width           =   3375
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Nothing>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   150
      TabIndex        =   3
      Top             =   1890
      Width           =   7095
   End
   Begin VB.Label Label1 
      Caption         =   "Save catalogue document as . . . "
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
      Left            =   150
      TabIndex        =   1
      Top             =   1530
      Width           =   3375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oCat As PapCat.z_Catalog
Attribute oCat.VB_VarHelpID = -1
Dim strTemplate As String
Dim strFilename As String
Dim flgShowWORD As Boolean

Dim tlCats As z_TextList
Dim tlCurr As z_TextList
Dim lngCurrID As Long
Dim lngCATID As Long
Dim fs As Scripting.FileSystemObject



Private Sub cboCats_Change()
    If tlCats.Key(cboCats.Text) > 0 Then
        lngCATID = tlCats.Key(cboCats.Text)
    End If

End Sub

Private Sub cboCats_Click()
    If tlCats.Key(cboCats.Text) > 0 Then
        lngCATID = tlCats.Key(cboCats.Text)
    End If
End Sub
Private Sub cboCurr_Change()
    If tlCurr.Key(cboCurr.Text) > 0 Then
        lngCurrID = tlCurr.Key(cboCurr.Text)
    End If

End Sub

Private Sub cboCurr_Click()
    If tlCurr.Key(cboCurr.Text) > 0 Then
        lngCurrID = tlCurr.Key(cboCurr.Text)
    End If
End Sub

Private Sub cmdFT_Click()
Dim fs As New Scripting.FileSystemObject
    
    If Not fs.FolderExists(gPapyConn.DatabaseFolder & "\Catalogues") Then
        fs.CreateFolder gPapyConn.DatabaseFolder & "\Catalogues"
    End If
    CD1.InitDir = gPapyConn.DatabaseFolder & "\Catalogues"
    CD1.Flags = cdlOFNCreatePrompt
    CD1.ShowSave
    If fs.FileExists(CD1.FileName) Or (CD1.FileName = "") Then
        MsgBox "You must specify a new file name!", vbInformation, "Invalid filename"
    Else
        strFilename = CD1.FileName
        Me.Label2 = strFilename
        Me.cmdOK.Enabled = True
    End If
End Sub

Private Sub cmdOK_Click()
On Error GoTo ERR_Handler
Dim dteStart, dteEnd As Date
Dim fs As Scripting.FileSystemObject

    Me.SB1.Panels(2).Text = "Generating catalogue, please wait . . ."
    
    Screen.MousePointer = vbHourglass
    dteStart = Now()
    
    Set oCat = New z_Catalog
    Me.cboCats.Enabled = False
    Me.cmdFT.Enabled = False
    Me.cmdOK.Enabled = False
    oCat.SetCurrency (lngCurrID)
    oCat.PrepareData lngCATID
    oCat.ExportToWORD strTemplate, strFilename, True, flgShowWORD
    Me.PB1.Visible = False
    MsgBox "Catalogue done!", vbOKOnly, "Status"
    Me.cboCats.Enabled = True
    Me.cmdFT.Enabled = True
    Screen.MousePointer = vbDefault
    Me.SB1.Panels(2).Text = ""
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    Resume
End Sub

Private Sub Form_Load()
    PB1.Align = vbAlignBottom
    PB1.Visible = False
    strTemplate = GetSetting(App.Title, "Settings", "Template", "")
    flgShowWORD = GetSetting(App.Title, "Settings", "ShowWORD", True)
    SB1.Panels(1).Text = "Template: " & strTemplate & IIf(flgShowWORD, " (visible)", "(Background)")
    Set tlCats = New z_TextList
    tlCats.Load "GetCatalogues_tl"
    LoadCombo cboCats, tlCats
    lngCATID = tlCats.Key(cboCats.Text)
    Set tlCurr = New z_TextList
    tlCurr.Load "GetCurrencies"
    LoadCombo cboCurr, tlCurr
    lngCurrID = tlCurr.Key(cboCurr.Text)
    
End Sub


Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuSettings_Click()
Dim frm As New frmSettings
    frm.Component strTemplate, flgShowWORD
    frm.Show vbModal
    strTemplate = frm.Template
    flgShowWORD = frm.ShowWORD
    SB1.Panels(1).Text = "Template: " & strTemplate & IIf(flgShowWORD, " (visible)", "(Background)")
    Unload frm
End Sub


Private Sub oCat_MaxRecs(i As Long)
    Me.PB1.Max = i
    Me.PB1.Min = 0
    Me.PB1.Visible = True
End Sub

Private Sub oCat_Status(i As Long)
    Me.PB1.Value = i
End Sub
