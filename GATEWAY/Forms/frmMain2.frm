VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmMain2 
   Caption         =   "DOLE registered services:  Web receptacle"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13260
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   13260
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAddDed 
      Alignment       =   1  'Right Justify
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
      Left            =   9495
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   5250
      Width           =   1230
   End
   Begin VB.TextBox txtDiffs 
      Alignment       =   1  'Right Justify
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
      Left            =   9480
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4845
      Width           =   1230
   End
   Begin VB.TextBox txtAdvRec 
      Alignment       =   1  'Right Justify
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
      Left            =   9465
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4440
      Width           =   1230
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
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
      Left            =   9435
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   4065
      Width           =   1230
   End
   Begin TrueOleDBGrid60.TDBGrid GAdvance 
      Height          =   585
      Left            =   7185
      OleObjectBlob   =   "frmMain2.frx":0000
      TabIndex        =   6
      Top             =   5835
      Width           =   1890
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboEntity 
      Height          =   315
      Left            =   345
      OleObjectBlob   =   "frmMain2.frx":5F01
      TabIndex        =   2
      Top             =   375
      Width           =   1665
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   6180
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
      Top             =   6090
      Width           =   13260
      _ExtentX        =   23389
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
   Begin EXCOMBOBOXLibCtl.ComboBox cboCNote 
      Height          =   315
      Left            =   2400
      OleObjectBlob   =   "frmMain2.frx":7127
      TabIndex        =   3
      Top             =   390
      Width           =   1665
   End
   Begin TrueOleDBGrid60.TDBGrid GFinal 
      Height          =   750
      Left            =   9300
      OleObjectBlob   =   "frmMain2.frx":834D
      TabIndex        =   7
      Top             =   5715
      Width           =   1395
   End
   Begin TrueOleDBGrid60.TDBGrid GCombo 
      Height          =   1635
      Left            =   390
      OleObjectBlob   =   "frmMain2.frx":DC50
      TabIndex        =   8
      Top             =   2415
      Width           =   10575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fruit specifications"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   390
      TabIndex        =   11
      Top             =   1980
      Width           =   3180
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Advance payment details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   3630
      TabIndex        =   10
      Top             =   1980
      Width           =   3270
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Final payment details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   6960
      TabIndex        =   9
      Top             =   1980
      Width           =   4005
   End
   Begin VB.Label Label2 
      Caption         =   "Consignment note"
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
      Left            =   2460
      TabIndex        =   5
      Top             =   165
      Width           =   1605
   End
   Begin VB.Label Label1 
      Caption         =   "Entity / Farm code"
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
      Left            =   420
      TabIndex        =   4
      Top             =   150
      Width           =   1545
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuImport 
         Caption         =   "&Import"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuDeleteALl 
         Caption         =   "&Delete all data"
      End
      Begin VB.Menu mnuShowFiles 
         Caption         =   "&Show names of all imported files"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
   End
End
Attribute VB_Name = "frmMain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents oImport As z_Import
Attribute oImport.VB_VarHelpID = -1
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim XC As XArrayDB
Dim tlCnotes As z_TextListSimple

Public Property Get Cnotes() As z_TextListSimple
    Set Cnotes = tlCnotes
End Property




Private Sub cboCNote_SelectionChanged()
    LoadArray1
    GAdvance.ReBind
    
    LoadArray2
    GFinal.ReBind
   
    LoadArray3
    Me.GCombo.ReBind
    
End Sub
Private Sub LoadArray1()
Dim rs As ADODB.Recordset
Dim oRet As New z_Retrieval
Dim lngIndex As Long
Dim i As Integer
Dim strCNote As String

    strCNote = cboCNote.Items.CellCaption(cboCNote.Items.SelectedItem, 0)
    Set rs = oRet.GetAdvances(strCNote)
    XA.Clear
    XA.ReDim 1, rs.RecordCount, 1, 10
    i = 1
    Do While Not rs.EOF
         XA.Value(i, 1) = Format(FNS(rs.Fields("RailDate")), "dd/mm/yyyy")
         XA.Value(i, 2) = FNS(rs.Fields("StatementDate"))  ', "dd/mm/yyyy"
         XA.Value(i, 3) = FNS(rs.Fields("GD_Variety"))
         XA.Value(i, 4) = FNS(rs.Fields("GD_Pack"))
         XA.Value(i, 5) = FNS(rs.Fields("GD_Grade"))
         XA.Value(i, 6) = FNS(rs.Fields("GD_Brand"))
         XA.Value(i, 7) = FNS(rs.Fields("GD_Count"))
         XA.Value(i, 8) = FNN(rs.Fields("SumOfQty_adv"))
         XA.Value(i, 9) = FNDBL(rs.Fields("SumOfAmt_Adv")) / FNN(rs.Fields("SumOfQty_Adv"))
         XA.Value(i, 10) = FNDBL(rs.Fields("SumOfAmt_Adv"))
        i = i + 1
        rs.MoveNext
    Loop
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASSCEND, XTYPE_STRING
    GAdvance.Array = XA
End Sub
Private Sub LoadArray2()
Dim rs As ADODB.Recordset
Dim oRet As New z_Retrieval
Dim lngIndex As Long
Dim i As Integer
Dim strCNote As String

    strCNote = cboCNote.Items.CellCaption(cboCNote.Items.SelectedItem, 0)
    Set rs = oRet.GetFinals(strCNote)
    XB.Clear
    XB.ReDim 1, rs.RecordCount, 1, 10
    i = 1
    Do While Not rs.EOF
         XB.Value(i, 1) = FNS(rs.Fields("PD_StatementRef"))
         XB.Value(i, 2) = Format(rs.Fields("StatementDateF"), "dd/mm/yyyy")
         XB.Value(i, 3) = FNS(rs.Fields("PD_Variety"))
         XB.Value(i, 4) = FNS(rs.Fields("PD_Pack"))
         XB.Value(i, 5) = FNS(rs.Fields("PD_Grade"))
         XB.Value(i, 6) = FNS(rs.Fields("PD_Brand"))
         XB.Value(i, 7) = FNS(rs.Fields("PD_Count"))
         XB.Value(i, 8) = FNN(rs.Fields("SumOfQty_Fin"))
         XB.Value(i, 9) = FNDBL(rs.Fields("SumOfAmt_Fin")) / FNN(rs.Fields("SumOfQty_Fin"))
         XB.Value(i, 10) = FNDBL(rs.Fields("SumOfAmt_Fin"))
        i = i + 1
        rs.MoveNext
    Loop
    XB.QuickSort 1, XB.UpperBound(1), 1, XORDER_ASSCEND, XTYPE_STRING
    GFinal.Array = XB
End Sub
Private Sub LoadArray3()
Dim rs As ADODB.Recordset
Dim oRet As New z_Retrieval
Dim lngIndex As Long
Dim i As Integer
Dim strCNote As String
Dim Final_tot As Double
Dim AdvancesRecovered As Double
Dim Differences_Tot As Double
Dim AddDeds_tot As Double

    strCNote = cboCNote.Items.CellCaption(cboCNote.Items.SelectedItem, 0)
    Set rs = oRet.GetBoth(strCNote)
    XC.Clear
    XC.ReDim 1, rs.RecordCount, 1, 20
    i = 1
    Final_tot = 0
    rs.MoveFirst
    Do While Not rs.EOF
        
         XC.Value(i, 1) = FNS(rs.Fields("Variety"))
         XC.Value(i, 2) = FNS(rs.Fields("Pack"))
         XC.Value(i, 3) = FNS(rs.Fields("Grade"))
         XC.Value(i, 4) = FNS(rs.Fields("Brand"))
         XC.Value(i, 5) = FNS(rs.Fields("Cnt"))
         XC.Value(i, 6) = Format(FNS(rs.Fields("RailDate")), "dd/mm/yyyy")
       '  XC.Value(i, 7) = FNS(rs.Fields("StatementDate"))  ', "dd/mm/yyyy"
         XC.Value(i, 8) = FNNF(rs.Fields("SumOfQty_Adv"))
        XC.Value(i, 9) = FNDBLF(rs.Fields("Rate_Adv"))
         XC.Value(i, 10) = FNDBLF(rs.Fields("SumOfAmt_Adv"))
     '    XC.Value(i, 11) = FNS(rs.Fields("STatementDateF"))
      '   XC.Value(i, 12) = FNS(rs.Fields("PD_StatementRef"))
         XC.Value(i, 13) = FNNF(rs.Fields("SumOfQty_Fin"))
        XC.Value(i, 14) = FNDBLF(rs.Fields("Rate_Fin"))
         XC.Value(i, 15) = FNDBLF(rs.Fields("SumOfAmt_Fin"))
         XC.Value(i, 16) = FNN(rs.Fields("Colour"))
        i = i + 1
        Final_tot = Final_tot + FNDBL(rs.Fields("SumOfAmt_Fin"))
        AdvancesRecovered = FNDBL(rs.Fields("SumOfFinAdvRec"))
        Differences_Tot = FNDBL(rs.Fields("SumOfDiffs"))
        AddDeds_tot = FNDBL(rs.Fields("AddDed"))
        rs.MoveNext
    Loop
  '  XC.QuickSort 1, XC.UpperBound(1), 1, XORDER_ASSCEND, XTYPE_STRING
    GCombo.Array = XC
    Me.txtTotal = Final_tot
    Me.txtAdvRec = AdvancesRecovered
    Me.txtDiffs = Differences_Tot
    Me.txtAddDed = AddDeds_tot
End Sub

Private Sub cboEntity_SelectionChanged()
Dim ar() As String
Dim sEntity As String
    sEntity = cboEntity.Items.CellCaption(cboEntity.Items.SelectedItem, 0)
    Set tlCnotes = Nothing
    Set tlCnotes = New z_TextListSimple
    tlCnotes.Load ltCNote, sEntity
    tlCnotes.CollectionAsArray ar
    cboCNote.BeginUpdate
    cboCNote.PutItems ar
    cboCNote.EndUpdate
End Sub


Private Sub Form_Initialize()
Dim ar() As String
    Set XA = New XArrayDB
    Set XB = New XArrayDB
    Set XC = New XArrayDB
    oPC.Entities.CollectionAsArray ar
    SetupCBOs
    cboEntity.BeginUpdate
    cboEntity.PutItems ar
    cboEntity.EndUpdate
'    cboEntity.Items.SelectItem(cboEntity.Items(0)) = True
End Sub


Private Sub Form_Terminate()
    Set tlCnotes = Nothing
    Set XA = Nothing
    Set XB = Nothing
    Set XC = Nothing
End Sub

Private Sub GCombo_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If XC(Bookmark, 16) < 1 Then
        RowStyle.BackColor = RGB(244, 248, 182)
    Else
        RowStyle.BackColor = RGB(209, 248, 182)
    End If
End Sub

Private Sub mnuDeleteALl_Click()
Dim oI As z_Import
    If MsgBox("You want to delete all data in the database." & vbCrLf & "This will require uploading data from the .CSV files again!" & vbCrLf & "Do you want to do this?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Exit Sub
    End If
    Set oI = New z_Import
    oI.DeleteAllData
    Set oI = Nothing
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuImport_Click()
Dim strError As String

    Set oImport = New z_Import
    WaitMsg "Deleting data . . .", True, Me
    oImport.DeleteAllData
    WaitMsg "", False, Me
    WaitMsg "Importing files . . .", True, Me
    oImport.ImportDelimitedFiles
    If oImport.Validate_LD_Seq(strError) = False Or oImport.ValidatePerType(strError) = False Then
        MsgBox "The import in invalid or incomplete for the following reason(s)" & vbCrLf & vbCrLf & strError, vbCritical, "Validation"
    End If
    WaitMsg "", False, Me
End Sub

Private Sub mnuShowFiles_Click()
Dim frm As New frmFileList
Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM vALLFiles ORDER BY FN", oPC.CO, adOpenForwardOnly, adLockReadOnly
    
    frm.Component rs
    frm.Show vbModal
    
End Sub

Private Sub oImport_InitProgress(pVal As Integer)
    PB1.Max = pVal
    PB1.Visible = True
End Sub

Private Sub oImport_UpdateProgress(pVal As Integer)
    PB1.Value = pVal
    If PB1.Value = PB1.Max Then
        PB1.Visible = False
    End If
End Sub





Private Sub SetupCBOs()
    cboEntity.BeginUpdate
    cboEntity.WidthList = 150
    cboEntity.HeightList = 162
    cboEntity.AllowSizeGrip = True
    cboEntity.AutoDropDown = True
    cboEntity.Columns.Add ""
    cboEntity.Columns(0).Width = 140
    cboEntity.BackColorLock = Me.BackColor
    cboEntity.EndUpdate
    
    cboCNote.BeginUpdate
    cboCNote.WidthList = 150
    cboCNote.HeightList = 162
    cboCNote.AllowSizeGrip = True
    cboCNote.AutoDropDown = True
    cboCNote.Columns.Add ""
    cboCNote.Columns(0).Width = 140
    cboCNote.BackColorLock = Me.BackColor
    cboCNote.EndUpdate
End Sub

