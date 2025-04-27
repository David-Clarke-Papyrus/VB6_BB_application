VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBrowsePOs 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Browse purchase orders"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   8310
   Begin VB.CommandButton cmdAdv 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Advanced"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   90
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5175
      UseMaskColor    =   -1  'True
      Width           =   930
   End
   Begin VB.CommandButton cmdSim 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Simple"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5910
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1605
      UseMaskColor    =   -1  'True
      Width           =   930
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
      Height          =   1485
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Select any one criteria.  If using dates, a selection between dates is catered for"
      Top             =   0
      Width           =   5295
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5640
         TabIndex        =   10
         Top             =   465
         Width           =   1530
      End
      Begin VB.ComboBox cboTP 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1920
         TabIndex        =   6
         Top             =   255
         Width           =   3135
      End
      Begin VB.TextBox txtTP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   5
         Top             =   255
         Width           =   500
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   4
         Top             =   630
         Width           =   1500
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3600
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   750
         UseMaskColor    =   -1  'True
         Width           =   1440
      End
      Begin VB.ComboBox cboSince 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmBrowseSOrders.frx":0000
         Left            =   1320
         List            =   "frmBrowseSOrders.frx":0013
         TabIndex        =   2
         Text            =   "Last week"
         Top             =   1005
         Width           =   1890
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Product code"
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
         Left            =   5715
         TabIndex        =   11
         Top             =   195
         Width           =   1290
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Supplier"
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
         Height          =   360
         Left            =   375
         TabIndex        =   9
         Top             =   285
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Supp. Ord. No."
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
         Height          =   360
         Left            =   15
         TabIndex        =   8
         Top             =   660
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Date"
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
         Height          =   360
         Left            =   780
         TabIndex        =   7
         Top             =   1035
         Width           =   480
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3570
      Left            =   105
      TabIndex        =   0
      Top             =   1590
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   6297
      SortKey         =   3
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
         Text            =   "Supplier"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   2294
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "srttg"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmBrowsePOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mcol As c_POs
Dim dPO As d_PO
Dim tlSupplier As z_TextList
Dim lngTPID As Long
Dim strNum As String
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim strDate1 As String
Dim strDate2 As String
Dim blnNoRecordsReturned As Boolean

Dim ofrm As frmPOPreview
Private Sub cboSince_Change()
    If cboSince <> "<none>" Then
        Me.cboTP = ""
        Me.txtNum = ""
    End If
End Sub

Private Sub cboSince_DblClick()
    Me.cboSince.ListIndex = 0
End Sub

Private Sub cboTP_LostFocus()
    If cboTP.ListIndex > -1 Then
        lngTPID = tlSupplier.Key(cboTP)
    End If
End Sub

Private Sub cboTP_Validate(Cancel As Boolean)
    If cboTP.ListIndex > -1 Then
        Me.cmdFind.Enabled = True
    End If
End Sub

Private Sub cmdAdv_Click()
    Me.Width = 8000
    Frame1.Width = 7500
End Sub
Private Sub cmdSim_Click()
    txtCode = ""
    Width = 5700
    Height = 5800
    Frame1.Width = 5295
End Sub

Private Sub cmdFind_Click()

    On Error GoTo ERR_Handler
    blnNoRecordsReturned = False
    
    Set mcol = Nothing
    Set mcol = New c_POs
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
    
    mcol.Load blnNoRecordsReturned, lngTPID, strNum, dteDate1, dteDate2, Trim$(txtCode)
    
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

Private Sub Lvw_DblClick()
Dim lngID As Long
Dim blnEdit As Boolean
    Set ofrm = New frmPOPreview
    lngID = val(lvw.SelectedItem.Key)
    ofrm.Component lngID    ', False
    ofrm.Show
End Sub

Private Sub Lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Lvw_DblClick
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
    Set tlSupplier = New z_TextList
    Set mcol = New c_POs
    Set dPO = New d_PO
    Me.Top = 50
    Me.Left = 50
    Me.Width = 5700
    Me.Height = 6070
    LoadControls
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tlSupplier = Nothing
    Set mcol = Nothing
    Set dPO = Nothing
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

Private Sub txtNum_Change()
    strNum = txtNum
End Sub

Private Sub txtNum_Validate(Cancel As Boolean)
    If txtNum > "" Then
        cmdFind.Enabled = True
    End If
End Sub

Private Sub txtTP_LostFocus()
    If Len(txtTP) <> 0 Then
        Set tlSupplier = Nothing
        Set tlSupplier = New z_TextList
        tlSupplier.Load ltSupplier, Me.txtTP
        LoadCombo Me.cboTP, tlSupplier
 '       retval = SendMessage(Me.cboTP.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
    End If
End Sub

Private Sub LoadControls()
'    txtDate1 = ""
'    txtDate2 = ""
    txtNum = ""
    txtTP = ""
    strDate1 = ""
    strDate2 = ""
    strNum = ""
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
            .Text = mcol(i).TPName
            .SubItems(1) = mcol(i).TRCode
            .SubItems(2) = mcol(i).TRDateF
            .SubItems(3) = mcol(i).CaptureDateForSort
            If mcol(i).status = "VOID" Then
                objItm.ForeColor = vbBlack
                .ListSubItems(1).ForeColor = vbBlack
                .ListSubItems(2).ForeColor = vbBlack
            ElseIf mcol(i).status = "IN PROCESS" Then
                objItm.ForeColor = vbRed
                .ListSubItems(1).ForeColor = vbRed
                .ListSubItems(2).ForeColor = vbRed
            End If
        End With
    Next i
End Sub




