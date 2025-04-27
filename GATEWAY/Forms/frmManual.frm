VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmManual 
   Caption         =   "DOLE registered services:  Upload (manual control)"
   ClientHeight    =   6825
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   10245
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrowerData 
      BackColor       =   &H00CCC8BB&
      Caption         =   "&Display first 1000 grower rows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   945
      Width           =   3120
   End
   Begin VB.CommandButton cmdBeloqDIP 
      BackColor       =   &H00CCC8BB&
      Caption         =   "&Display first 1000 below DIP rows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   495
      Width           =   3135
   End
   Begin VB.CommandButton cmdImportBelowDIP 
      BackColor       =   &H00CCC8BB&
      Caption         =   "&Import below DIP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   945
      Width           =   2370
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00CCC8BB&
      Caption         =   "Test report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7035
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6045
      Width           =   1890
   End
   Begin VB.CommandButton cmdGenPDFs 
      BackColor       =   &H00CCC8BB&
      Caption         =   "Generate PDFs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7020
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5550
      Width           =   1890
   End
   Begin VB.CommandButton cmdFTP 
      BackColor       =   &H00CCC8BB&
      Caption         =   "&Export to FTP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7035
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5010
      Width           =   1890
   End
   Begin VB.CommandButton cmdSplit 
      BackColor       =   &H00CCC8BB&
      Caption         =   "&Split data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7065
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4470
      Width           =   1890
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00CCC8BB&
      Caption         =   "&Import grower data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   495
      Width           =   2370
   End
   Begin VB.CommandButton cmdImport 
      BackColor       =   &H00CCC8BB&
      Caption         =   "&Import payment data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   45
      Width           =   2385
   End
   Begin VB.TextBox txtError 
      Height          =   960
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   5235
      Width           =   6450
   End
   Begin VB.CommandButton cmdPaymentData 
      BackColor       =   &H00CCC8BB&
      Caption         =   "&Display first 1000 payment rows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   45
      Width           =   3120
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00CCC8BB&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4470
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   6330
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   6405
      Visible         =   0   'False
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11360
            MinWidth        =   11360
         EndProperty
      EndProperty
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Height          =   2655
      Left            =   60
      OleObjectBlob   =   "frmManual.frx":0000
      TabIndex        =   4
      Top             =   1575
      Width           =   10695
   End
   Begin VB.Label Label1 
      Caption         =   "Error messages"
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
      Left            =   60
      TabIndex        =   10
      Top             =   4965
      Width           =   2490
   End
End
Attribute VB_Name = "frmManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XA As XArrayDB
Private Sub cmdGrowerData_Click()
    On Error GoTo errHandler
    LoadArray "B"
    G.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual.cmdGrowerData_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdOK_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdBeloqDIP_Click()
    On Error GoTo errHandler
    LoadArray "C"
    G.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual.cmdBeloqDIP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFTP_Click()
    On Error GoTo errHandler
Dim oEx As New z_Export
    WaitMsg "Uploading to FTP site . . .", True, Me
    oEx.SendFiles
    WaitMsg "", False, Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual.cmdFTP_Click", , EA_NORERAISE
    HandleError
End Sub





Private Sub cmdSplit_Click()
    On Error GoTo errHandler
Dim oSplit As New z_Split
    WaitMsg "Splitting data . . .", True, Me
    oSplit.ExporttoFile Date
    WaitMsg "", False, Me
    Me.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual.cmdSplit_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub LoadArray(pType As String)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oImp As New z_Import
Dim lngIndex As Long
Dim i As Integer
Dim j As Integer
Dim strCNote As String
Dim strError As String
Dim limit As Long

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    strError = ""
    oImp.LoadFromAS400 rs, strError, pType
    If strError = "" Then
      XA.Clear
      XA.ReDim 0, rs.RecordCount - 1, 0, rs.Fields.Count - 1
      i = 0
      For j = 0 To rs.Fields.Count - 1
        If j < G.Columns.Count Then
          G.Columns.Item(j).Caption = rs.Fields(j).Name
          G.Columns.Item(j).Width = 500
          End If
      Next j
      If rs.RecordCount > 5000 Then
        limit = 5000
    Else
        limit = rs.RecordCount
    End If
      Do While i < limit 'rs.RecordCount
          For j = 0 To rs.Fields.Count - 1
              XA.Value(i, j) = FNS(rs.Fields(j))
          Next j
          i = i + 1
          rs.MoveNext
      Loop
    '  XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
      G.Array = XA
      txtError = ""
    Else
        Me.txtError = strError
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual.LoadArray(pType)", pType
End Sub




Private Sub Command2_Click()
    On Error GoTo errHandler
Dim oImp As New z_Import
Dim strError As String

    strError = ""
    oImp.ImportFROMAS400_GR strError
    If strError <> "" Then
        txtError = strError
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual.Command2_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
    Set XA = New XArrayDB
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuConfiguration_Click()
    On Error GoTo errHandler
Dim frm As New frmConfiguration
    oPC.Configuration.BeginEdit
    frm.Show vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual.mnuConfiguration_Click", , EA_NORERAISE
    HandleError
End Sub







Private Sub mnuDeleteALl_Click()
    On Error GoTo errHandler
Dim oImp As New z_Import
    oImp.DeleteAllData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual.mnuDeleteALl_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuExit_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual.mnuExit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuImport_Click()
    On Error GoTo errHandler
Dim oImp As New z_Import
  '  oImp.LoadFromAS400 (0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmManual.mnuImport_Click", , EA_NORERAISE
    HandleError
End Sub

