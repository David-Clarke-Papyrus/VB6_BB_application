VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWriteDown 
   Caption         =   "Write-down stock"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   8670
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Report"
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
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3675
      Width           =   1275
   End
   Begin VB.TextBox txtTotals 
      Alignment       =   1  'Right Justify
      Height          =   705
      Left            =   4455
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   135
      Width           =   3990
   End
   Begin VB.TextBox txtCostPrice 
      Alignment       =   2  'Center
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   4860
      TabIndex        =   4
      Text            =   "R 0.00"
      Top             =   3270
      Width           =   1320
   End
   Begin VB.CommandButton cmdCHange 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Write down"
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
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3210
      Width           =   1275
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboSection 
      Height          =   390
      Left            =   360
      OleObjectBlob   =   "frmWriteDown.frx":0000
      TabIndex        =   1
      Top             =   435
      Width           =   3780
   End
   Begin MSComctlLib.ListView lvwLines 
      Height          =   2265
      Left            =   360
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   885
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3995
      View            =   3
      SortOrder       =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Cost"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Qty on hand"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblTotals 
      Height          =   315
      Left            =   375
      TabIndex        =   6
      Top             =   3240
      Width           =   3075
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "New cost price"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   3150
      TabIndex        =   5
      Top             =   3315
      Width           =   1605
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   360
      TabIndex        =   2
      Top             =   180
      Width           =   915
   End
End
Attribute VB_Name = "frmWriteDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ar() As String
Dim rs As adodb.Recordset
Dim OpenResult As Integer
Dim tlSections As z_TextList
Dim dblCostValue As Double
Dim dblCostValueExVat As Double
Dim lngCost As Long

Public Sub Component()
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

End Sub


Private Sub setcombo()

    Set tlSections = Nothing
    Set tlSections = New z_TextList
    tlSections.Load ltSectionsActive

    cboSection.BeginUpdate
    tlSections.CollectionAsArray ar
    cboSection.PutItems ar
    cboSection.EndUpdate
End Sub

Private Sub SetupcboArray()
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

End Sub

Private Sub cboSection_Click()
    Set rs = Nothing
    Set rs = New adodb.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT dbo.CodeF(P_CODE,P_EAN,0) as codeF,LEFT(P_TITLE,40) + '/' + P_MainAuthor as description,P_COST,P_QtyOnHand,CostValue,CostValueExVAT,OldCostValue,OldCostValueExVAT FROM vWriteDownList WHERE DICT_ID = " & tlSections.Key(cboSection.Value), oPC.COShort, adOpenDynamic
    Me.lblTotals.Caption = CStr(rs.RecordCount) & " items"
    LoadGrid
End Sub

Private Sub cmdCHange_Click()
Dim strSQL As String

    If MsgBox("Confirm you want to write down the stock listed to " & Format(CDbl(lngCost) / 100, "###,##0.00"), vbOKCancel + vbQuestion, "Confirm action") = vbCancel Then
        Exit Sub
    End If
    strSQL = "UPDATE tPRODUCT SET P_SPECIAL = a.P_COST FROM tPRODUCT a JOIN vWriteDownList b ON a.P_ID = b.P_ID WHERE b.DICT_ID = " & tlSections.Key(cboSection.Value)
    oPC.COShort.Execute strSQL
    
    strSQL = "UPDATE tPRODUCT SET P_Cost = " & CStr(lngCost) & " FROM tPRODUCT a JOIN vWriteDownList b ON a.P_ID = b.P_ID WHERE b.DICT_ID = " & tlSections.Key(cboSection.Value)
    oPC.COShort.Execute strSQL
    
    MsgBox "The writedown is complete."
    
End Sub

Private Sub cmdReport_Click()
Dim rptWriteDowns As arWriteDowns
    Set rs = Nothing
    Set rs = New adodb.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT dbo.CodeF(P_CODE,P_EAN,0) as codeF,LEFT(P_TITLE,40) + '/' + P_MainAuthor as description,P_COST,P_QtyOnHand,CostValue,CostValueExVAT,p_Special,OldCostValue,OldCostValueExVAT FROM vWriteDownList WHERE DICT_ID = " & tlSections.Key(cboSection.Value), oPC.COShort, adOpenDynamic
    
    Set rptWriteDowns = New arWriteDowns
    rptWriteDowns.Component rs, "Write-downs done on " & Format(Date, "dd/mm/yyyy") & " for set " & cboSection.Value
    rptWriteDowns.Show vbModal

End Sub

Private Sub Form_Load()
    SetupcboArray
    setcombo

End Sub

Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

Private Sub LoadGrid()
Dim li As ListItem
    dblCostValue = 0
    dblCostValueExVat = 0
    lvwLines.ListItems.Clear
    Do While Not rs.EOF
        Set li = lvwLines.ListItems.Add
        li.Text = FNS(rs.Fields(0))
        li.SubItems(1) = FNS(rs.Fields(1))
        li.SubItems(2) = Format(FNDBL(rs.Fields(2)), "###,##0.00")
        li.SubItems(3) = FNS(rs.Fields(3))
        dblCostValue = dblCostValue + FNDBL(rs.Fields("CostValue"))
        dblCostValueExVat = dblCostValueExVat + FNDBL(rs.Fields("CostValueExVat"))
        rs.MoveNext
    Loop
    txtTotals.Text = "Value at cost: " & Format(dblCostValue, "###,##0.00")
End Sub

Private Sub txtCostPrice_GotFocus()
    txtCostPrice = CStr(lngCost)
End Sub

Private Sub txtCostPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    
    txtCostPrice = Trim(txtCostPrice)
    If Trim$(txtCostPrice) = "" Then
        lngCost = 0
    ElseIf (Not ConvertToLng(txtCostPrice, lngCost)) Then
        Cancel = True
        Exit Sub
    End If
    txtCostPrice = Format(CDbl(lngCost) / 100, "###,##0.00")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtCostPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

