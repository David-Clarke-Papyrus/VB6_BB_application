VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBrowseAPPRs 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Browse Appro Returns"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5505
   ScaleWidth      =   5955
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
      Height          =   1905
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Select any one criteria.  If using dates, a selection between dates is catered for"
      Top             =   120
      Width           =   5655
      Begin VB.ComboBox cboSince 
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
         ItemData        =   "frmBrowseApproReturns.frx":0000
         Left            =   1680
         List            =   "frmBrowseApproReturns.frx":0013
         TabIndex        =   10
         Text            =   "Last quarter"
         Top             =   1395
         Width           =   1890
      End
      Begin VB.TextBox txtCode 
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
         Height          =   330
         Left            =   1680
         TabIndex        =   9
         Top             =   1005
         Width           =   1935
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
         Left            =   3840
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   1425
      End
      Begin VB.TextBox txtISBN 
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
         Left            =   1680
         TabIndex        =   4
         Top             =   630
         Width           =   1980
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
         Left            =   1680
         TabIndex        =   3
         Top             =   255
         Width           =   500
      End
      Begin VB.ComboBox cboTP 
         Appearance      =   0  'Flat
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
         Left            =   2190
         TabIndex        =   2
         Top             =   240
         Width           =   3090
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
         Height          =   360
         Left            =   420
         TabIndex        =   11
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Code"
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
         Left            =   855
         TabIndex        =   8
         Top             =   1035
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contains Book"
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
         Left            =   60
         TabIndex        =   7
         Top             =   660
         Width           =   1410
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Customer"
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
         Left            =   585
         TabIndex        =   6
         Top             =   285
         Width           =   885
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3315
      Left            =   135
      TabIndex        =   0
      Top             =   2115
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5847
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
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmBrowseAPPRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tlCustomer As z_TextList
Dim cAppRet As c_APPRs
Dim dAPPR As d_APPR

Dim lngTPID As Long
Dim strApproRetNum As String
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim strDate1 As String
Dim strDate2 As String
Dim blnNoRecordsReturned As Boolean
Dim strISBN As String

Dim ofrm As frmAPPRPreview

Private Sub cmdFind_Click()

    On Error GoTo ERR_Handler
    blnNoRecordsReturned = False
    
    Set cAppRet = Nothing
    Set cAppRet = New c_APPRs
    MousePointer = vbHourglass
    Me.lvw.ListItems.Clear
   
    Select Case cboSince
    Case "<any date>"
        dteDate1 = CDate("1995-01-01")
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
   
    cAppRet.Load blnNoRecordsReturned, lngTPID, strApproRetNum, strISBN, dteDate1, dteDate2
    
    If blnNoRecordsReturned Then
        MsgBox "No records found", vbOKOnly + vbInformation, "Papyrus Appro Returns Information"
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

'Dim ofrm As frmApproPreview

Private Sub Form_Load()

    Set tlCustomer = New z_TextList
    Set cAppRet = New c_APPRs
    Set dAPPR = New d_APPR
    Me.Top = 50
    Me.Left = 50
    Me.Width = 6075
    Me.Height = 5910
    LoadControls
    
End Sub

Private Sub LoadControls()
    txtCode = ""
    txtTP = ""
    txtISBN = ""
    strApproRetNum = ""
    strDate1 = ""
    strDate2 = ""
    strISBN = ""
    lngTPID = 0
    
    cboTP.ListIndex = -1
End Sub

Private Sub LoadListView()
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvw.ListItems.Clear
    For i = 1 To cAppRet.Count
        Set objItm = Me.lvw.ListItems.Add
        With objItm
            .Key = cAppRet(i).TRID & "K"
            .Text = cAppRet(i).TPName
            .SubItems(1) = cAppRet(i).TRDateF
            .SubItems(2) = cAppRet(i).TRCode
            If cAppRet(i).statusF = "VOID" Then
                objItm.ForeColor = CL_DARKBLUE
            ElseIf cAppRet(i).statusF = "IN PROCESS" Then
                objItm.ForeColor = vbRed
            End If
        End With
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tlCustomer = Nothing
    Set cAppRet = Nothing
    Set dAPPR = Nothing
End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub Lvw_DblClick()
Dim lngID As Long
Dim blnEdit As Boolean
    Set ofrm = New frmAPPRPreview
    lngID = val(lvw.SelectedItem.Key)
    ofrm.Component lngID    ', False
    ofrm.Show
End Sub

Private Sub txtCode_Change()
    strApproRetNum = txtCode
End Sub

Private Sub txtISBN_Change()
    strISBN = txtISBN
End Sub

Private Sub txtTP_LostFocus()
    If Len(txtTP) <> 0 Then
        Set tlCustomer = Nothing
        Set tlCustomer = New z_TextList
        tlCustomer.Load ltCustomer, Me.txtTP
        LoadCombo Me.cboTP, tlCustomer
    End If
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

Private Sub cmdFind_LostFocus()
'    lngTPID = 0
    LoadControls
End Sub
