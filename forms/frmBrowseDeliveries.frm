VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBrowseDELS 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Browse deliveries"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   5325
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
      Top             =   -45
      Width           =   4950
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
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   1200
         TabIndex        =   6
         Text            =   "Last quarter"
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
         Height          =   705
         Left            =   3135
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   630
         UseMaskColor    =   -1  'True
         Width           =   1680
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
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   1185
         TabIndex        =   4
         Top             =   630
         Width           =   1860
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
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   1185
         TabIndex        =   3
         Top             =   255
         Width           =   500
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
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   1710
         TabIndex        =   2
         Top             =   255
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dated in"
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
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   1050
         Width           =   1050
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Delivery No."
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
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   990
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
         Height          =   255
         Left            =   195
         TabIndex        =   7
         Top             =   285
         Width           =   885
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3255
      Left            =   105
      TabIndex        =   0
      Top             =   1545
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Customer"
         Object.Width           =   3775
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Appro number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmBrowseDELS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mcol As c_DELs
Dim dDEL As d_DEL
Dim tlSupplier As z_TextList
Dim lngTPID As Long
Dim strNum As String
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim strDate1 As String
Dim strDate2 As String
Dim blnNoRecordsReturned As Boolean

Dim ofrm As frmDELPreview
Private Sub cboSince_Change()
    If cboSince <> "<none>" Then
        Me.cboTP = ""
        Me.txtnum = ""
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

Private Sub cmdFind_Click()

    On Error GoTo ERR_Handler
    blnNoRecordsReturned = False
    
    Set mcol = Nothing
    Set mcol = New c_DELs
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
    
    mcol.Load blnNoRecordsReturned, lngTPID, strNum, dteDate1, dteDate2
    
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

Private Sub lvw_AfterLabelEdit(Cancel As Integer, NewString As String)
Cancel = True
End Sub

Private Sub lvw_BeforeLabelEdit(Cancel As Integer)
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
    Set ofrm = New frmDELPreview
    lngID = Val(lvw.SelectedItem.Key)
    ofrm.Component lngID    ', False
    ofrm.Show
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
    Set tlSupplier = New z_TextList
    Set mcol = New c_DELs
    Set dDEL = New d_DEL
    Me.Top = 50
    Me.Left = 50
    Me.Width = 5300
    Me.Height = 5800
    LoadControls
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tlSupplier = Nothing
    Set mcol = Nothing
    Set dDEL = Nothing
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
    strNum = txtnum
End Sub

Private Sub txtNum_Validate(Cancel As Boolean)
    If txtnum > "" Then
        cmdFind.Enabled = True
    End If
End Sub

Private Sub txtTP_LostFocus()
    If Len(txtTP) <> 0 Then
        Set tlSupplier = Nothing
        Set tlSupplier = New z_TextList
        tlSupplier.Load ltSupplier, Me.txtTP
        LoadCombo Me.cboTP, tlSupplier
    End If
End Sub

Private Sub LoadControls()
    txtnum = ""
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
            .Text = mcol(i).TPName & (IIf(Len(Trim(mcol(i).TPACCNo)) <= 1, "", "(" & Trim(mcol(i).TPACCNo) & ")"))
            .SubItems(1) = mcol(i).TPName
            .SubItems(1) = mcol(i).TRCode
            .SubItems(2) = Format(mcol(i).TRDate, "Short Date")
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




