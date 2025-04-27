VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmBrowseCustomers 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse customers"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8625
   Icon            =   "frmBrowseCustomers.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEmailInsertList 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Create EMail insert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6330
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2415
      Width           =   2040
   End
   Begin VB.CommandButton cmdDeselectAll 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Deselect all"
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2535
      Width           =   1200
   End
   Begin VB.CommandButton cmdSelectAll 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Select all"
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
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2535
      Width           =   1200
   End
   Begin VB.CommandButton cmdAddSelected 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Add selected to current list"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3750
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   8025
      Width           =   1500
   End
   Begin VB.CommandButton cmdManage 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&View current list"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   8010
      Width           =   1155
   End
   Begin VB.CommandButton cmdLists 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Select current list"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   8010
      Width           =   1155
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   7305
      Picture         =   "frmBrowseCustomers.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7965
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Search in address for . . ."
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
      Height          =   1380
      Left            =   5925
      TabIndex        =   6
      Top             =   15
      Width           =   2460
      Begin VB.TextBox txtAddress 
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
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "Enter an address fragment and click FIND."
         Top             =   300
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddress 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Fin&d"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   885
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Click to find all customers mwith an address containing . . ."
         Top             =   660
         UseMaskColor    =   -1  'True
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Select any one criteria.  If using dates, a selection between dates is catered for"
      Top             =   15
      Width           =   5565
      Begin VB.OptionButton optOR 
         BackColor       =   &H00C4BCA4&
         Caption         =   "or"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4980
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1065
         Width           =   480
      End
      Begin VB.OptionButton optAnd 
         BackColor       =   &H00C4BCA4&
         Caption         =   "and"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   570
         Value           =   -1  'True
         Width           =   480
      End
      Begin VB.ComboBox cboStores 
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
         Left            =   75
         TabIndex        =   20
         Top             =   1890
         Width           =   2100
      End
      Begin VB.ComboBox cboIG3 
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
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   "Select a customer grouping"
         Top             =   1245
         Width           =   1725
      End
      Begin VB.ComboBox cboIG2 
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
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Select a customer grouping"
         Top             =   847
         Width           =   1725
      End
      Begin VB.ComboBox cboCT 
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
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Select a customer type or allow any customer to be found"
         Top             =   1155
         Width           =   1725
      End
      Begin VB.ComboBox cboIG1 
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
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Select a customer grouping"
         Top             =   450
         Width           =   1725
      End
      Begin VB.CommandButton cmdFind1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3915
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Click to find all customers matching the retrictions entered."
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   1005
      End
      Begin VB.TextBox txtArg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   90
         TabIndex        =   0
         ToolTipText     =   "Enter A/C number, name, start of name, telephone number or end of telephone number. Hit ENTER to fetch."
         Top             =   480
         Width           =   2550
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Originating store"
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
         Left            =   105
         TabIndex        =   21
         Top             =   1620
         Width           =   1560
      End
      Begin VB.Label lblRecordsFound 
         BackStyle       =   0  'Transparent
         Height          =   330
         Left            =   1845
         TabIndex        =   17
         Top             =   1155
         Width           =   1260
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer type"
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
         Height          =   270
         Left            =   105
         TabIndex        =   15
         Top             =   900
         Width           =   1290
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Interest group"
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
         Height          =   270
         Left            =   3495
         TabIndex        =   13
         Top             =   195
         Width           =   1290
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Look for name, phone or ACno."
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
         Height          =   270
         Left            =   90
         TabIndex        =   5
         Top             =   240
         Width           =   2565
      End
   End
   Begin TrueOleDBGrid60.TDBGrid CustGrid 
      Height          =   4965
      Left            =   105
      OleObjectBlob   =   "frmBrowseCustomers.frx":0635
      TabIndex        =   1
      Top             =   2910
      Width           =   8265
   End
   Begin VB.Label lblDefaultListName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   345
      Left            =   1335
      TabIndex        =   10
      Top             =   8085
      Width           =   2325
   End
End
Attribute VB_Name = "frmBrowseCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCust As c_C_Customer
Dim dispCust As d_C_Customer
Dim lngTPID As Long
Dim strACCNum As String
Dim oCust As a_Customer
Dim blnNoRecordsReturned As Boolean
Dim XA As New XArrayDB
'#If H_CENTRAL <> 1 Then
Dim ofrm As frmCustomerPreview
'#End If
Dim ofrmLoy As frmLoyaltyPreview
Dim CustomerTypes_tl As z_TextList
Dim InterestGroups_tl As z_TextList
Dim arDir(1 To 6) As Integer


Private Sub cmdAddress_Click()
    On Error GoTo errHandler
    FindByAddress
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.cmdAddress_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAddSelected_Click()
    On Error GoTo errHandler
Dim i As Long
Dim strSQL As String
    If lngDefaultListID = 0 Then
        MsgBox "You must select a customer list first.", , "Can't do this"
    Else
        For i = 1 To XA.UpperBound(1)
            If XA(i, 1) = True Then
        'For i = 1 To CustGrid.SelBookmarks.Count
                strSQL = "INSERT INTO tLISTITEM (LISTITEM_LIST_ID,LISTITEM_TP_ID) VALUES (" & lngDefaultListID & "," & CLng(XA(i, 7)) & ")"
                oPC.CO.Execute strSQL
            End If
        Next i
    End If
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217873 Then Resume Next
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.cmdAddSelected_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdFind1_Click()
    On Error GoTo errHandler
    Find
    If XA.UpperBound(1) > 1 Then
        lblRecordsFound.Caption = XA.UpperBound(1) & " Records"
    Else
        lblRecordsFound.Caption = XA.UpperBound(1) & " Record"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.cmdFind1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdLists_Click()
    On Error GoTo errHandler
Dim frm As New frmLists
    frm.Show vbModal
    If lngDefaultListID > 0 Then
        lblDefaultListName.Caption = strDefaultListName
    Else
        lblDefaultListName.Caption = ""
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.cmdLists_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdManage_Click()
    On Error GoTo errHandler
Dim frm As New frmListsManage
    frm.Show vbModal
    Unload frm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.cmdManage_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSelectAll_Click()
Dim i As Integer
    For i = 1 To XA.UpperBound(1)
        XA(i, 1) = True
    Next
    Me.CustGrid.Refresh
End Sub
Private Sub cmdDeselectAll_Click()
Dim i As Integer
    For i = 1 To XA.UpperBound(1)
        XA(i, 1) = False
    Next
    Me.CustGrid.Refresh
End Sub

Private Sub cmdEmailInsertList_Click()
Dim oTF As New z_TextFile
Dim strLine As String
Dim oCust As d_C_Customer
Dim rs As ADODB.Recordset

    oTF.OpenTextFile oPC.SharedFolderRoot & "\EMAIL_LIST.CSV"
    For Each oCust In cCust
        Set rs = New ADODB.Recordset
        rs.Open "SELECT ADD_EMail from tADD WHERE ADD_TP_ID = '" & oCust.ID & "'", oPC.CO, adOpenForwardOnly, adLockReadOnly
        Do While Not rs.EOF
            strLine = oCust.Appell & "," & oCust.Initials & "," & oCust.Name & "," & oCust.AcNo & "," & FNS(rs!ADD_EMAIL)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        
        oTF.WriteToTextFile strLine
    Next
    oTF.CloseTextFile
    Set oTF = Nothing
    MsgBox "File exported to " & oPC.SharedFolderRoot & "\EMAIL_LIST.CSV", , "Status"
    
    
    
    
End Sub

Private Sub CustGrid_LostFocus()
    CustGrid.Update
End Sub

Private Sub CustGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuBrowseCustomerPopup   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.CustGrid_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, X, Y), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub CustGrid_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = 13 Then
        CustGrid_DblClick
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.CustGrid_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub FindByAddress()
    On Error GoTo errHandler
Dim bRecsFound As Boolean
    blnNoRecordsReturned = False
    Set cCust = Nothing
    Set cCust = New c_C_Customer
    MousePointer = vbHourglass
    cCust.LoadForAddress bRecsFound, txtAddress
    If blnNoRecordsReturned Then
        MsgBox "No records found", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        GoTo EXIT_Handler
    End If
    LoadArray
    CustGrid.ReBind
    If XA.UpperBound(1) > 1 Then
        lblRecordsFound.Caption = XA.UpperBound(1) & " Records"
    Else
        lblRecordsFound.Caption = XA.UpperBound(1) & " Record"
    End If
EXIT_Handler:
    MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.FindByAddress"
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub CustGrid_DblClick()
    On Error GoTo errHandler
Dim lngID As Long
Dim blnEdit As Boolean
    If IsNull(CustGrid.Bookmark) Then Exit Sub
    lngID = Val(XA(CustGrid.Bookmark, 7))
    Set oCust = Nothing
    Set oCust = New a_Customer
    oCust.Load lngID
 '   If oCust.IsLoyaltyClubMember Then
'        Set ofrmLoy = New frmLoyaltyPreview
'        ofrmLoy.Component oCust    ', False
'        ofrmLoy.Show
'#If H_CENTRAL <> 1 Then
 '   Else
        Set ofrm = New frmCustomerPreview
        ofrm.Component oCust    ', False
        ofrm.Show
'#End If
 '   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.CustGrid_DblClick", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdAdv_Click()
    On Error GoTo errHandler
        txtAddress = ""
        Width = 4800
        Height = 6300

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.cmdAdv_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Find()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    Set cCust = Nothing
    Set cCust = New c_C_Customer
    cCust.LoadEasy txtArg, 0, oPC.Configuration.Stores_tl.Key(cboStores), InterestGroups_tl.Key(cboIG1), InterestGroups_tl.Key(cboIG2), InterestGroups_tl.Key(cboIG3), CustomerTypes_tl.Key(cboCT), 0, IIf(optAnd = True, "AND", "OR")
    LoadArray
    CustGrid.ReBind
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.Find"
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    
    'sets up the defalt sort directions
    arDir(1) = 0
    arDir(2) = 1
    arDir(3) = 1
    arDir(4) = 1
    arDir(5) = 1
    arDir(6) = 1

    
    Me.top = 50
    Me.left = 50
    Width = 9000
    Height = 9150
    Set CustomerTypes_tl = New z_TextList
    Set InterestGroups_tl = New z_TextList
    
    oPC.Configuration.LoadStores_tl "<ANY>"
    LoadCombo cboStores, oPC.Configuration.Stores_tl
    
    CustomerTypes_tl.Load ltCustomerTypeActive, , "<ALL>"
    InterestGroups_tl.Load ltInterestGroupActive, , "<ANY>"
    LoadCombo cboIG1, InterestGroups_tl
    LoadCombo cboIG2, InterestGroups_tl
    LoadCombo cboIG3, InterestGroups_tl
    LoadCombo cboCT, CustomerTypes_tl
    cboCT = CustomerTypes_tl.Item("0")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set oCust = Nothing
    Set cCust = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_C_Customer
Dim itmList As ListItem
Dim lngIndex As Long
    XA.ReDim 1, cCust.Count, 1, 7
    For lngIndex = 1 To cCust.Count
        With objItem
            Set objItem = cCust.Item(lngIndex)
            XA.Value(lngIndex, 2) = objItem.Fullname2
            XA.Value(lngIndex, 3) = objItem.AcNo
            XA.Value(lngIndex, 4) = objItem.Cellf
            XA.Value(lngIndex, 5) = objItem.SalesQty
            XA.Value(lngIndex, 6) = objItem.SalesValue
            XA.Value(lngIndex, 7) = objItem.ID
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    CustGrid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.LoadArray"
End Sub




Private Sub txtArg_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = 13 Then
        Find
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = 13 Then
        FindByAddress
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.txtAddress_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Public Sub AddToList()
    On Error GoTo errHandler
Dim i As Long
Dim strSQL As String
    If lngDefaultListID = 0 Then
        MsgBox "You must select a customer list first.", , "Can't do this"
    Else
        For i = 1 To CustGrid.SelBookmarks.Count
            strSQL = "INSERT INTO tLISTITEM (LISTITEM_LIST_ID,LISTITEM_TP_ID) VALUES (" & lngDefaultListID & "," & CLng(XA(CustGrid.SelBookmarks(i - 1), 7)) & ")"
            oPC.CO.Execute strSQL
        Next i
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.AddToList"
End Sub

Public Sub RemoveFromList()
    On Error GoTo errHandler
    MsgBox "remove"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.RemoveFromList"
End Sub
Private Sub CustGrid_HeadClick(ByVal ColIndex As Integer)
Static Direction As Variant

If ColIndex = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
'    If Direction = 0 Then
'        Direction = 1
'    Else
'        Direction = 0
'    End If
    If arDir(ColIndex + 1) = 1 Then
        arDir(ColIndex + 1) = 0
    Else
        arDir(ColIndex + 1) = 1
    End If
    
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, arDir(ColIndex + 1), GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    CustGrid.Refresh
    Screen.MousePointer = vbDefault

End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    Select Case ColIndex
        Case 1, 2, 3, 4
            GetRowType = XTYPE_STRING
'        Case 3, 4
'            GetRowType = XTYPE_DATE
    End Select
End Function


