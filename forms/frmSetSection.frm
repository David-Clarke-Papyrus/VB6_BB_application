VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetSection 
   Caption         =   "Set category"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCHange 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Change category"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3990
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3150
      Width           =   2175
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboSection 
      Height          =   390
      Left            =   150
      OleObjectBlob   =   "frmSetSection.frx":0000
      TabIndex        =   0
      Top             =   3150
      Width           =   3780
   End
   Begin MSComctlLib.ListView lvwLines 
      Height          =   2265
      Left            =   150
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   165
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   3995
      View            =   3
      SortOrder       =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483635
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Categories"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Please note: Any existing categories for these products will be removed."
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   150
      TabIndex        =   4
      Top             =   2490
      Width           =   7365
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   150
      TabIndex        =   1
      Top             =   2895
      Width           =   915
   End
End
Attribute VB_Name = "frmSetSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ar() As String
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim tlSections As z_TextList
Public Sub component(pIDs As String)
    On Error GoTo errHandler
Dim strSQL As String
Dim oSQL As New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    'strSQL = "SELECT P_TITLE,dbo.CodeF(P_CODE,P_EAN,0) as CodeF,PT_CODE,P_ID,P_MainAuthor from tPRODUCT LEFT JOIN tPT ON P_PRODUCTTYPE_ID = tPT.PT_ID WHERE P_ID IN (" & pIDs & ")"
    strSQL = "SELECT P_TITLE,dbo.CodeF(P_CODE,P_EAN,0) as CodeF,dbo.FlattenSections(P_ID),P_ID,P_MainAuthor from tPRODUCT  WHERE P_ID IN (" & pIDs & ")"
    oSQL.RunGetRecordset strSQL, enText, "", "", rs

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSetSection.component(pIDs)", pIDs
End Sub
Private Sub LoadGrid()
    On Error GoTo errHandler
Dim li As ListItem
    
    Do While Not rs.eof
        Set li = lvwLines.ListItems.Add
        li.text = FNS(rs.fields(1))
        li.SubItems(1) = FNS(rs.fields(0))
        li.SubItems(2) = FNS(rs.fields(2))
        li.SubItems(3) = FNS(rs.fields(3))
        li.Checked = True
        rs.MoveNext
    Loop
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSetSection.LoadGrid"
End Sub
Private Sub setcombo()
    On Error GoTo errHandler

    Set tlSections = Nothing
    Set tlSections = New z_TextList
    tlSections.Load ltSectionsActive

    cboSection.BeginUpdate
    tlSections.CollectionAsArray ar
    cboSection.PutItems ar
    cboSection.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSetSection.setcombo"
End Sub

Private Sub SetupcboArray()
    On Error GoTo errHandler
    cboSection.BeginUpdate
    cboSection.WidthList = 270
    cboSection.HeightList = 262
    cboSection.AllowSizeGrip = True
    cboSection.AutoDropDown = True
    cboSection.SelForeColor = vbRed
    cboSection.Columns.Add "Section"
    cboSection.Columns.Add ""
    cboSection.Columns(0).Width = 245
    cboSection.Columns(1).Width = 0
    cboSection.BackColorLock = Me.BackColor
    cboSection.EndUpdate

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSetSection.SetupcboArray"
End Sub

Private Sub cmdCHange_Click()
    On Error GoTo errHandler
Dim i As Integer
Dim cnt As Long

    If Not cboSection.Value > "" Then
        MsgBox "Select a section from the drop-down list first.", vbInformation, "Can't do this"
        Exit Sub
    End If
    cnt = 0
    For i = 1 To lvwLines.ListItems.Count
        If lvwLines.ListItems.Item(i).Checked Then
            oPC.COShort.execute "Delete FROM tProductSection FROM tPRODUCTSECTION JOIN tDICT ON PSEC_SEC_ID = DICT_ID WHERE PSEC_P_ID = '" & lvwLines.ListItems.Item(i).SubItems(3) & "' AND ISNULL(DICT_SYSTEM,'') = ''"
            oPC.COShort.execute "INSERT INTO tProductSection (PSEC_P_ID,PSEC_SEC_ID,PSEC_PRIORITY) VALUES ('" & lvwLines.ListItems.Item(i).SubItems(3) & "'," & tlSections.Key(cboSection.Value) & ",99)"
            cnt = cnt + 1
        End If
    Next i
    MsgBox CStr(cnt) & " items updated", vbInformation, "Status"
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSetSection.cmdCHange_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    SetupcboArray
    setcombo
    LoadGrid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSetSection.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSetSection.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
