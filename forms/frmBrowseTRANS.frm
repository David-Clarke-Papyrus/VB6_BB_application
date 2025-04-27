VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBrowseTRANS 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Browse credit notes"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvw 
      Height          =   3450
      Left            =   45
      TabIndex        =   7
      Top             =   1455
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   6085
      SortKey         =   2
      View            =   3
      Arrange         =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Customer"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Invoice No."
         Object.Width           =   2294
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Invoice Date"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "SortTag"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   45
      TabIndex        =   4
      ToolTipText     =   "Select any one criteria.  If using dates, a selection between dates is catered for"
      Top             =   -75
      Width           =   4950
      Begin VB.ComboBox cboSince 
         Height          =   345
         ItemData        =   "frmBrowseTRANS.frx":0000
         Left            =   1200
         List            =   "frmBrowseTRANS.frx":0013
         TabIndex        =   8
         Text            =   "Last week"
         Top             =   1020
         Width           =   1890
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Default         =   -1  'True
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
         Left            =   3195
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   735
         UseMaskColor    =   -1  'True
         Width           =   1425
      End
      Begin VB.TextBox txtCNNum 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1185
         TabIndex        =   2
         Top             =   630
         Width           =   1500
      End
      Begin VB.TextBox txtTP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1185
         TabIndex        =   0
         Top             =   255
         Width           =   500
      End
      Begin VB.ComboBox cboTP 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1710
         TabIndex        =   1
         Top             =   255
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dated in"
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   60
         TabIndex        =   9
         Top             =   1035
         Width           =   1050
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Order no."
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   90
         TabIndex        =   6
         Top             =   645
         Width           =   1050
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Customer"
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   210
         TabIndex        =   5
         Top             =   285
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmBrowseTRANS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mcol As c_CNs
Dim dispCN As d_CN
Dim tlCustomer As z_TextList
Dim lngTPID As Long
Dim strCNNum As String
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim strDate1 As String
Dim strDate2 As String
Dim blnNoRecordsReturned As Boolean

Dim ofrm As frmInvoicePreview
Dim ofrmCN As frmCNPreview
Private Sub cboSince_Change()
    If cboSince <> "<none>" Then
        Me.cboTP = ""
        Me.txtCNNum = ""
    End If
End Sub

Private Sub cboSince_DblClick()
    Me.cboSince.ListIndex = 0
End Sub

Private Sub cboTP_LostFocus()
    If cboTP.ListIndex > -1 Then
        lngTPID = tlCustomer.Key(cboTP)
    End If
End Sub

Private Sub cboTP_Validate(Cancel As Boolean)
    If cboTP.ListIndex > -1 Then
        Me.cmdFind.Enabled = True
    End If
End Sub

Private Sub cmdFind_Click()

    On Error GoTo ERR_Handler
    blnNoRecordsReturned = False
    
    Set mcol = Nothing
    Set mcol = New c_CNs
    MousePointer = vbHourglass
    Me.lvw.ListItems.Clear
    
    Select Case cboSince
    Case "<any date>"
        dteDate1 = CDate("1995-01-01")
        dteDate2 = DateAdd("d", 1, Date)
    Case "Last week"
        dteDate1 = DateAdd("d", -7, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case "Last month"
        dteDate1 = DateAdd("m", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case "Last quarter"
        dteDate1 = DateAdd("q", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case "Last year"
        dteDate1 = DateAdd("yyyy", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    End Select
    
    mcol.Load blnNoRecordsReturned, lngTPID, "", strCNNum, dteDate1, dteDate2
    
    If blnNoRecordsReturned Then
        MsgBox "No records found", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        GoTo EXIT_Handler
    End If
    
    LoadListView

EXIT_Handler:
    MousePointer = vbDefault
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub

Private Sub cmdFind_LostFocus()
    LoadControls
End Sub

Private Sub Lvw_AfterLabelEdit(Cancel As Integer, NewString As String)
Cancel = True
End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub Lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' When a ColumnHeader object is clicked, the ListView control is
    ' sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1
    
    lvw.SortKey = ColumnHeader.Index - 1
    ' Set Sorted to True to sort the list.
    lvw.Sorted = True
    If lvw.SortOrder = lvwAscending Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvw_DblClick()
Dim lngID As Long
Dim blnEdit As Boolean
    Set ofrmCN = New frmCNPreview
    lngID = val(lvw.SelectedItem.Key)
    ofrmCN.Component lngID    ', False
    ofrmCN.Show
End Sub

Private Sub Lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then lvw_DblClick
End Sub

'Private Sub txtDate1_Change()
'   strDate1 = txtDate1
'End Sub

'Private Sub txtDate1_Validate(Cancel As Boolean)
'    If strDate1 > "" Then
'        dteDate1 = CDate(strDate1)
'        cmdFind.Enabled = True
'    End If
'End Sub
'
'Private Sub txtDate2_Change()
'    strDate2 = txtDate2
'End Sub

Private Sub Form_Load()
    Set tlCustomer = New z_TextList
    Set mcol = New c_CNs
    Set dispCN = New d_CN
    Me.Top = 50
    Me.Left = 50
    Me.Width = 5300
    Me.Height = 5800
    LoadControls
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tlCustomer = Nothing
    Set mcol = Nothing
    Set dispCN = Nothing
    Set ofrm = Nothing
End Sub

Private Sub txtDate2_Validate(Cancel As Boolean)
    If IsNull(dteDate1) Then
        MsgBox "Please ensure that a date is entered in the first box" & vbCrLf & "before entering a date in the second date box" _
                    , vbOKOnly, "Papyrus Invoices Information"
  '      txtDate2 = ""
  '      txtDate1.SetFocus
    ElseIf IsDate(strDate2) Then
        dteDate2 = CDate(strDate2)
        cmdFind.Enabled = True
    End If
End Sub

Private Sub txtCNNum_Change()
    strCNNum = txtCNNum
End Sub

Private Sub txtCNNum_Validate(Cancel As Boolean)
    If txtCNNum > "" Then
        cmdFind.Enabled = True
    End If
End Sub

Private Sub txtTP_LostFocus()
    If Len(txtTP) <> 0 Then
        Set tlCustomer = Nothing
        Set tlCustomer = New z_TextList
        tlCustomer.Load ltCustomer, Me.txtTP
        LoadCombo Me.cboTP, tlCustomer
 '       retval = SendMessage(Me.cboTP.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
    End If
End Sub

Private Sub LoadControls()
'    txtDate1 = ""
'    txtDate2 = ""
    txtCNNum = ""
    txtTP = ""
    strDate1 = ""
    strDate2 = ""
    strCNNum = ""
    lngTPID = 0
    
    cboTP.ListIndex = -1
End Sub

Private Sub LoadListView()
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvw.ListItems.Clear
    For i = 1 To mcol.Count
        Set objItm = Me.lvw.ListItems.Add
        With objItm
            .Key = mcol(i).TRID & "K"
            .Text = mcol(i).TPName & (IIf(Len(Trim(mcol(i).TPACCNo)) <= 1, "", "(" & Trim(mcol(i).TPACCNo) & ")"))
           ' .SubItems(1) = mcol(i).TPName
            .SubItems(1) = mcol(i).Ref
            .SubItems(2) = mcol(i).TRDateF
            .SubItems(3) = mcol(i).DateForSort
            If mcol(i).statusF = "VOID" Then
                objItm.ForeColor = vbBlack
                .ListSubItems(1).ForeColor = vbBlack
                .ListSubItems(2).ForeColor = vbBlack
            ElseIf mcol(i).statusF = "IN PROCESS" Then
                objItm.ForeColor = vbRed
                .ListSubItems(1).ForeColor = vbRed
                .ListSubItems(2).ForeColor = vbRed
            End If
        End With
    Next i
End Sub


