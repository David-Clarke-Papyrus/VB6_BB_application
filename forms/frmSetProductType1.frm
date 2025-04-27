VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetProductType 
   Caption         =   "Set product type"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin EXCOMBOBOXLibCtl.ComboBox cboProductType 
      Height          =   315
      Left            =   375
      OleObjectBlob   =   "frmSetProductType1.frx":0000
      TabIndex        =   3
      Top             =   3015
      Width           =   3150
   End
   Begin VB.CommandButton cmdCHange 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Change product type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvwLines 
      Height          =   2265
      Left            =   150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
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
         Name            =   "Arial"
         Size            =   11.25
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
         Text            =   "Product type"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Product type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   150
      TabIndex        =   1
      Top             =   2730
      Width           =   1545
   End
End
Attribute VB_Name = "frmSetProductType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ar() As String
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim tlProductTypes As z_TextList
Public Sub Component(pIDs As String)
Dim strSQL As String
Dim oSQL As New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT P_TITLE,dbo.CodeF(P_CODE,P_EAN,0) as CodeF,PT_CODE,P_ID,P_MainAuthor from tPRODUCT LEFT JOIN tPT ON P_PRODUCTTYPE_ID = tPT.PT_ID WHERE P_ID IN (" & pIDs & ")"
    oSQL.RunGetRecordset strSQL, enText, "", "", rs

End Sub
Private Sub LoadGrid()
Dim li As ListItem
    
    Do While Not rs.EOF
        Set li = lvwLines.ListItems.Add
        li.Text = FNS(rs.Fields(1))
        li.SubItems(1) = FNS(rs.Fields(0))
        li.SubItems(2) = FNS(rs.Fields(2))
        li.SubItems(3) = FNS(rs.Fields(3))
        li.Checked = True
        rs.MoveNext
    Loop
    
End Sub
Private Sub setcombo()

    Set tlProductTypes = Nothing
    Set tlProductTypes = New z_TextList
    tlProductTypes.Load ltProductTypeActive

    cboProductType.BeginUpdate
    tlProductTypes.CollectionAsArray ar
    cboProductType.PutItems ar
    cboProductType.EndUpdate
End Sub

Private Sub SetupcboArray()
    cboProductType.BeginUpdate
    cboProductType.WidthList = 270
    cboProductType.HeightList = 262
    cboProductType.AllowSizeGrip = True
    cboProductType.AutoDropDown = True
    cboProductType.SelForeColor = vbRed
    cboProductType.Columns.Add "Product type"
    cboProductType.Columns.Add ""
    cboProductType.Columns(0).Width = 245
    cboProductType.Columns(1).Width = 0
    cboProductType.BackColorLock = Me.BackColor
    cboProductType.EndUpdate

End Sub

Private Sub cmdCHange_Click()
Dim i As Integer
Dim cnt As Long

    If Not cboProductType.Value > "" Then
        MsgBox "Select a product type from the drop-down list first.", vbInformation, "Can't do this"
        Exit Sub
    End If
    cnt = 0
    For i = 1 To lvwLines.ListItems.Count
        If lvwLines.ListItems.Item(i).Checked Then
            oPC.COShort.Execute "UPDATE tPRODUCT SET P_PRODUCTTYPE_ID = " & tlProductTypes.Key(cboProductType.Value) & " WHERE P_ID = '" & lvwLines.ListItems.Item(i).SubItems(3) & "'"
            cnt = cnt + 1
        End If
    Next i
    MsgBox CStr(cnt) & " items updated", vbInformation, "Status"
    Unload Me
End Sub

Private Sub Form_Load()
    SetupcboArray
    setcombo
    LoadGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub
