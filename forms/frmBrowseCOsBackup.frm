VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmBrowseCOsBackup 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Browse orders"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
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
   ScaleHeight     =   5700
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
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
      Left            =   90
      TabIndex        =   4
      ToolTipText     =   "Select any one criteria.  If using dates, a selection between dates is catered for"
      Top             =   -75
      Width           =   4950
      Begin CoolButtonControl.CoolButton CB1 
         Height          =   300
         Left            =   2655
         TabIndex        =   10
         Top             =   795
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   529
         BackColor       =   14737632
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Normal"
         Style           =   1
         BackStyle       =   0
      End
      Begin VB.ComboBox cboSince 
         Height          =   345
         ItemData        =   "frmBrowseCOsBackup.frx":0000
         Left            =   990
         List            =   "frmBrowseCOsBackup.frx":0013
         TabIndex        =   7
         Text            =   "Last week"
         Top             =   1020
         Width           =   1530
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
         Height          =   480
         Left            =   3900
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   870
         UseMaskColor    =   -1  'True
         Width           =   930
      End
      Begin VB.TextBox txtCONum 
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
         Left            =   975
         TabIndex        =   2
         Top             =   645
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
         Left            =   975
         TabIndex        =   0
         Top             =   255
         Width           =   500
      End
      Begin VB.ComboBox cboTP 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1500
         TabIndex        =   1
         Top             =   225
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dated in"
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   195
         TabIndex        =   8
         Top             =   1035
         Width           =   690
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Order no."
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   660
         Width           =   780
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Customer"
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   75
         TabIndex        =   5
         Top             =   285
         Width           =   810
      End
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Height          =   3345
      Left            =   135
      OleObjectBlob   =   "frmBrowseCOsBackup.frx":0053
      TabIndex        =   9
      Top             =   1545
      Width           =   4875
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   3390
      Left            =   120
      Top             =   1530
      Width           =   4905
   End
End
Attribute VB_Name = "frmBrowseCOsBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mcol As c_COs
Dim dispCO As d_CO
Dim tlCustomer As z_TextList
Dim lngTPID As Long
Dim strInvoiceNum As String
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim strDate1 As String
Dim strDate2 As String
Dim enCOType As enumCOType
Dim blnNoRecordsReturned As Boolean
Dim flgLoading As Boolean
Dim ofrm As frmInvoicePreview
Dim ofrmCO As frmCOPreview
Dim XA As New XArrayDB

Private Sub CB1_Click()
    enCOType = OptionLoop(enCOType, 2)
    Select Case enCOType
    Case enNormalCO
        CB1.Caption = "Normal"
    Case enWant
        CB1.Caption = "Wants"
    End Select
End Sub

Private Sub cboSince_Change()
    If cboSince <> "<none>" Then
        Me.cboTP = ""
        Me.txtCONum = ""
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
    Set mcol = New c_COs
    MousePointer = vbHourglass
  '  Me.lvwCOs.ListItems.Clear
    
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
    
    mcol.Load blnNoRecordsReturned, enCOType, lngTPID, strInvoiceNum, dteDate1, dteDate2
    
    If blnNoRecordsReturned Then
        MsgBox "No records found", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        XA.Clear
        Grid.ReBind
        GoTo EXIT_Handler
    End If
    
    LoadArray
    Grid.ReBind

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

Private Sub lvwCOs_AfterLabelEdit(Cancel As Integer, NewString As String)
Cancel = True
End Sub

Private Sub lvwCOs_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

'Private Sub lvwCOs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    ' When a ColumnHeader object is clicked, the ListView control is
'    ' sorted by the subitems of that column.
'    ' Set the SortKey to the Index of the ColumnHeader - 1
'
'    lvwCOs.SortKey = ColumnHeader.Index - 1
'    ' Set Sorted to True to sort the list.
'    lvwCOs.Sorted = True
'    If lvwCOs.SortOrder = lvwAscending Then
'        lvwCOs.SortOrder = lvwDescending
'    Else
'        lvwCOs.SortOrder = lvwAscending
'    End If
'End Sub
'
'
'Private Sub lvwCOs_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then lvwCOs_DblClick
'End Sub

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
    enCOType = enNormalCO
    Set tlCustomer = New z_TextList
    Set mcol = New c_COs
    Set dispCO = New d_CO
    Me.Top = 50
    Me.Left = 50
    Me.Width = 5300
    Me.Height = 5900
    LoadControls
    CB1.Visible = oPC.Configuration.SupportsWants
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tlCustomer = Nothing
    Set mcol = Nothing
    Set dispCO = Nothing
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


'Private Sub optNormal_Click()
'    If flgLoading Then Exit Sub
'    iCOType = enNormalCO
'End Sub
'
'Private Sub optStandingOrders_Click()
'    If flgLoading Then Exit Sub
'    iCOType = enStanding
'End Sub
'
'Private Sub optWants_Click()
'    If flgLoading Then Exit Sub
'    iCOType = enWant
'End Sub

Private Sub txtCONum_Change()
    strInvoiceNum = txtCONum
End Sub

Private Sub txtCONum_Validate(Cancel As Boolean)
    If txtCONum > "" Then
        cmdFind.Enabled = True
        cboSince = "<any date>"
    End If
End Sub

Private Sub txtTP_LostFocus()
    If Len(txtTP) <> 0 Then
        Set tlCustomer = Nothing
        Set tlCustomer = New z_TextList
        tlCustomer.Load ltCustomer, Me.txtTP
        LoadCombo Me.cboTP, tlCustomer
        cboTP.ListIndex = 0
        Me.cboSince = "<any date>"
 '       retval = SendMessage(Me.cboTP.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
    End If
End Sub

Private Sub LoadControls()
    flgLoading = True
    txtCONum = ""
    txtTP = ""
    strDate1 = ""
    strDate2 = ""
    strInvoiceNum = ""
    lngTPID = 0
    Select Case enCOType
    Case enNormalCO
        CB1.Caption = "Normal"
    Case enWant
        CB1.Caption = "Wants"
    End Select
    flgLoading = False
End Sub

'Private Sub LoadListView()
'Dim objItm As ListItem
'Dim i As Integer
'Dim tmp As String
'
'    lvwCOs.ListItems.Clear
'    For i = 1 To mcol.Count
'        Set objItm = Me.lvwCOs.ListItems.Add
'        With objItm
'            .Key = mcol(i).TRID & "K"
'            .Text = mcol(i).TPName & (IIf(Len(Trim(mcol(i).TPACCNo)) <= 1, "", "(" & Trim(mcol(i).TPACCNo) & ")"))
'           ' .SubItems(1) = mcol(i).TPName
'            .SubItems(1) = mcol(i).TRCode
'            .SubItems(2) = mcol(i).TRDateF
'            .SubItems(3) = mcol(i).CaptureDateForSort
'            If mcol(i).StatusF = "VOID" Then
'                objItm.ForeColor = vbBlack
'                .ListSubItems(1).ForeColor = vbBlack
'                .ListSubItems(2).ForeColor = vbBlack
'            ElseIf mcol(i).StatusF = "IN PROCESS" Then
'                objItm.ForeColor = vbRed
'                .ListSubItems(1).ForeColor = vbRed
'                .ListSubItems(2).ForeColor = vbRed
'            End If
'        End With
'    Next i
'End Sub
Private Sub LoadArray()
Dim objItem As d_Customer
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.Clear
    XA.ReDim 1, mcol.Count, 1, 6
    For i = 1 To mcol.Count
        With objItem
            XA.Value(i, 1) = mcol(i).TPName & (IIf(Len(Trim(mcol(i).TPACCNo)) <= 1, "", "(" & Trim(mcol(i).TPACCNo) & ")"))
            XA.Value(i, 2) = mcol(i).TRCode
            XA.Value(i, 3) = mcol(i).TRDateF
            XA.Value(i, 4) = mcol(i).CaptureDateForSort
            XA.Value(i, 5) = mcol(i).TRID & "K"
            XA.Value(i, 6) = mcol(i).StatusF
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 4, XORDER_ASCEND, XTYPE_STRING
    Grid.Array = XA
End Sub

Private Sub Grid_DblClick()
Dim lngID As Long
Dim blnEdit As Boolean
    Set ofrmCO = New frmCOPreview
    lngID = val(XA(Grid.Row + 1, 5))
    ofrmCO.Component lngID    ', False
    ofrmCO.Show
End Sub
Private Sub Grid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If XA(Bookmark, 6) > "VOID" Then
        RowStyle.ForeColor = vbBlack
    End If
    If XA(Bookmark, 6) = "IN PROCESS" Then
        RowStyle.ForeColor = vbRed
    End If
'            If mcol(i).StatusF = "VOID" Then
'                objItm.ForeColor = vbBlack
'                .ListSubItems(1).ForeColor = vbBlack
'                .ListSubItems(2).ForeColor = vbBlack
'            ElseIf mcol(i).StatusF = "IN PROCESS" Then
'                objItm.ForeColor = vbRed
'                .ListSubItems(1).ForeColor = vbRed
'                .ListSubItems(2).ForeColor = vbRed
'            End If

End Sub

