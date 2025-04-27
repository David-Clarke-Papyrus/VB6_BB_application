VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmConfiguration 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Configuration"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   10245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   6240
      Picture         =   "frmConfiguration.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6330
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7500
      Picture         =   "frmConfiguration.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6345
      Width           =   1260
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5880
      Left            =   180
      TabIndex        =   3
      Top             =   330
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   10372
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   7
      TabHeight       =   741
      BackColor       =   13882315
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "frmConfiguration.frx":0B14
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtLookupSeq"
      Tab(0).Control(1)=   "cboLocalCountry"
      Tab(0).Control(2)=   "txtVATRate"
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(5)=   "Label3"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Companies"
      TabPicture(1)   =   "frmConfiguration.frx":0B30
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRemove"
      Tab(1).Control(1)=   "cmdDefault"
      Tab(1).Control(2)=   "cmdEditComp"
      Tab(1).Control(3)=   "cmdAddComp"
      Tab(1).Control(4)=   "lvwCompanies"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Currencies"
      TabPicture(2)   =   "frmConfiguration.frx":0B4C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdRemCurr"
      Tab(2).Control(1)=   "cmdDefaultCurr"
      Tab(2).Control(2)=   "cmdEditCurr"
      Tab(2).Control(3)=   "cmdAddCurr"
      Tab(2).Control(4)=   "cmdLocal"
      Tab(2).Control(5)=   "lvwCurrencies"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Stores"
      TabPicture(3)   =   "frmConfiguration.frx":0B68
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lvwStores"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Command1"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdSetDefaultStore"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdStoreEdit"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdAddStore"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin VB.CommandButton cmdRemCurr 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Remove"
         Height          =   405
         Left            =   -72585
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5220
         Width           =   1095
      End
      Begin VB.CommandButton cmdDefaultCurr 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Set as default"
         Height          =   345
         Left            =   -66870
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         Width           =   1485
      End
      Begin VB.CommandButton cmdEditCurr 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Edit"
         Height          =   405
         Left            =   -73695
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   5220
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddCurr 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
         Height          =   405
         Left            =   -74790
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5220
         Width           =   1095
      End
      Begin VB.CommandButton cmdLocal 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Set as local"
         Height          =   345
         Left            =   -66870
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   960
         Width           =   1485
      End
      Begin VB.TextBox txtLookupSeq 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   5
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -68775
         TabIndex        =   18
         Top             =   2190
         Width           =   1140
      End
      Begin VB.ComboBox cboLocalCountry 
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
         Left            =   -69765
         TabIndex        =   16
         Top             =   1530
         Width           =   2835
      End
      Begin VB.CommandButton cmdAddStore 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
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
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdStoreEdit 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Edit"
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
         Left            =   1365
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdSetDefaultStore 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Set as default"
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
         Left            =   6870
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   735
         Width           =   1500
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Remove"
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
         Left            =   2475
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4545
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Remove"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   -72555
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4515
         Width           =   1095
      End
      Begin VB.CommandButton cmdDefault 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Set as default"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -68010
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Width           =   1470
      End
      Begin VB.CommandButton cmdEditComp 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   -73650
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4515
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddComp 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4515
         Width           =   1095
      End
      Begin VB.TextBox txtVATRate 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   5
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73020
         TabIndex        =   4
         Top             =   1110
         Width           =   915
      End
      Begin MSComctlLib.ListView lvwCompanies 
         Height          =   3615
         Left            =   -74775
         TabIndex        =   10
         Top             =   780
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6376
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14416635
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Default company"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwStores 
         Height          =   3735
         Left            =   210
         TabIndex        =   15
         Top             =   750
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   6588
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14416635
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Default company"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwCurrencies 
         Height          =   4500
         Left            =   -74820
         TabIndex        =   25
         Top             =   585
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   7938
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14416635
         BorderStyle     =   1
         Appearance      =   1
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
            Text            =   "Name"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Symbol"
            Object.Width           =   776
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Format string"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Factor"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Default"
            Object.Width           =   1834
         EndProperty
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Book lookup sequence (e.g. BF or BFWH or WH or WHBF). The books are looked for on the sources in the sequence indicated."
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
         Height          =   540
         Left            =   -74835
         TabIndex        =   19
         Top             =   2190
         Width           =   5880
      End
      Begin VB.Label Label2 
         Caption         =   "Local country"
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
         Left            =   -69765
         TabIndex        =   17
         Top             =   1260
         Width           =   1755
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "VAT Rate"
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
         Left            =   -74850
         TabIndex        =   5
         Top             =   1170
         Width           =   1755
      End
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H00D3D3CB&
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   195
      TabIndex        =   2
      Top             =   6285
      Width           =   1875
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oConfig As a_c_Configuration
Dim flgLoading As Boolean
Dim tlCountries As New z_TextList
Dim lngOperatorID As Long

Private Sub EnableOK(pOK As Boolean)
    Me.cmdOK.Enabled = pOK
End Sub
Private Sub oConfig_Valid(pErrors As String, Status As Boolean)
    EnableOK Status
    lblErrors = pErrors
End Sub

Public Sub Component(poConfig As a_c_Configuration)
    Set oConfig = poConfig
    oConfig.BeginEdit
End Sub
Private Sub LoadControls()
Dim rs As New ADODB.Recordset

    flgLoading = True
    Me.txtVATRate = oConfig.vatRate
    Me.txtLookupSeq = oConfig.LookupSeq
    tlCountries.Load ltCountry
    LoadCombo Me.cboLocalCountry, tlCountries
    cboLocalCountry = tlCountries.Item(CStr(oConfig.LocalCountryID))
    
    
    flgLoading = False
    FillStoresList
    FillCurrencyList
    SSTab1.Tab = 0
End Sub
Private Sub FillStoresList()
Dim objItem As a_Store
Dim itmList As ListItem
Dim lngIndex As Long
    Me.lvwStores.ListItems.Clear
    For lngIndex = 1 To oConfig.Stores.Count
        With objItem
            Set objItem = oConfig.Stores.Item(lngIndex)
            Set itmList = lvwStores.ListItems.Add(Key:=objItem.Key)
            With itmList
                .Text = objItem.Description
                If objItem.IsDeleted Then .Text = .Text & "(deleted)"
                If objItem.IsNew Then .Text = .Text & "(New)"
                .SubItems(1) = objItem.code
                If objItem.ID = oConfig.DefaultStoreID Then .SubItems(1) = "default"
            End With
        End With
    Next
End Sub

Private Sub cboLocalCountry_Click()
    If flgLoading Then Exit Sub
    oConfig.LocalCountryID = tlCountries.Key(cboLocalCountry)
End Sub







Private Sub cmdAddStore_Click()
Dim oStore As a_Store
Dim frm As frmStore

    Set oStore = oConfig.Stores.Add
    Set frm = New frmStore
    frm.Component oStore
    frm.Show vbModal
    FillStoresList
End Sub

Private Sub cmdCancel_Click()
    oConfig.CancelEdit
    Unload Me
End Sub




Private Sub cmdOK_Click()
Dim strStatus As String

    oConfig.ApplyEdit strStatus
    If strStatus <> "" Then
        strStatus = "The save operation has not been successful for the following reason:" & vbCrLf & vbCrLf & strStatus & vbCrLf & vbCrLf & "Either select Cancel or correct the data and select OK again."
        MsgBox strStatus
    Else
        Unload Me
    End If
End Sub

Private Sub cmdRemCurr_Click()
    On Error GoTo errHandler
Dim oCurr As a_Currency
Dim lngResult As Long
Dim oSQL As New z_SQL
Dim bCanDelete As Boolean

    bCanDelete = (oSQL.QtyDocumentsUsingCurrency(oConfig.Currencies.Item(lvwCurrencies.SelectedItem.Key).ID) = 0)
    If bCanDelete Then
        oConfig.Currencies.Item(lvwCurrencies.SelectedItem.Key).Delete
        FillCurrencyList
    Else
        MsgBox "This currency is associated with documents in your database. You cannot delete it." & vbCrLf _
        & "You should first merge this currency with another, then you will be able to delete it.", vbInformation + vbOKOnly, "Can't do this"
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdRemCurr_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSetDefaultStore_Click()
    oConfig.DefaultStoreID = oConfig.Stores(Val(lvwStores.SelectedItem.Key)).ID
    FillStoresList
End Sub

Private Sub cmdStoreEdit_Click()
    EditStore
End Sub


Private Sub Form_Load()
    LoadControls
End Sub



Private Sub EditStore()
Dim frm As New frmStore
Dim oStore As a_Store
Dim lngResult As Long

    Set oStore = New a_Store
    Set oStore = oConfig.Stores.Item(lvwStores.SelectedItem.Key)
    frm.Component oStore
    frm.Show vbModal
    FillStoresList
End Sub

Private Sub lvwStores_AfterLabelEdit(Cancel As Integer, NewString As String)
Cancel = True
End Sub

Private Sub lvwStores_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub txtLookupSeq_Change()
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oConfig.setLookupSequence txtLookupSeq
    If Err Then
      Beep
      intPos = txtLookupSeq.SelStart
      txtLookupSeq = oConfig.LookupSeq
      txtLookupSeq.SelStart = intPos - 1
    End If

End Sub
Private Sub txtLookupSeq_GotFocus()
    AutoSelect Controls("txtLookupSeq")
End Sub

Private Sub txtLookupSeq_LostFocus()
   txtLookupSeq.Text = oConfig.LookupSeq
End Sub


Private Sub txtVATRate_Change()
    If flgLoading Then Exit Sub
    
    oConfig.SetVATRate txtVATRate
End Sub
Private Sub txtVATRate_LostFocus()
   txtVATRate.Text = oConfig.vatRate
End Sub

Private Sub cmdAddCurr_Click()
    On Error GoTo errHandler
Dim oCurr As a_Currency
Dim frm As frmCurrency

    Set oCurr = oConfig.Currencies.Add
    Set frm = New frmCurrency
    frm.Component oCurr
    frm.Show vbModal
    FillCurrencyList

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdAddCurr_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEditCurr_Click()
    On Error GoTo errHandler
    EditCurr
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdEditCurr_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdDefaultCurr_Click()
    On Error GoTo errHandler
    oConfig.DefaultCurrencyID = oConfig.Currencies(lvwCurrencies.SelectedItem.Key).ID
    FillCurrencyList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdDefaultCurr_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdLocal_Click()
    On Error GoTo errHandler
    oConfig.LocalCurrencyID = oConfig.Currencies(lvwCurrencies.SelectedItem.Key).ID
    FillCurrencyList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdLocal_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub FillCurrencyList()
    On Error GoTo errHandler
Dim objItem As a_Currency
Dim itmList As ListItem
Dim lngIndex As Long

    lvwCurrencies.ListItems.Clear
    For lngIndex = 1 To oConfig.Currencies.Count
        With objItem
            Set objItem = oConfig.Currencies.Item(lngIndex)
            Set itmList = lvwCurrencies.ListItems.Add(Key:=objItem.Key)
            With itmList
                .Text = objItem.Description
                If objItem.IsDeleted Then .Text = .Text & "(deleted)"
                If objItem.IsNew Then .Text = .Text & "(New)"
                .SubItems(1) = objItem.Symbol
                .SubItems(2) = objItem.FormatString
                .SubItems(3) = objItem.FactorF & "/" & objItem.FactorINVF
                If oConfig.DefaultCurrencyID = oConfig.LocalCurrencyID Then
                    If objItem.ID = oConfig.DefaultCurrencyID Then .SubItems(4) = "Default and Local"
                Else
                    If objItem.ID = oConfig.DefaultCurrencyID Then .SubItems(4) = "Default"
                    If objItem.ID = oConfig.LocalCurrencyID Then .SubItems(4) = "Local"
                End If
            '''''    If objItem.ID = oConfig.LocalCurrencyID Then .SubItems(2) = "default"
            End With
        End With
    Next


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.FillCurrencyList"
End Sub


Private Sub EditCurr()
    On Error GoTo errHandler
Dim frm As New frmCurrency
Dim oCurr As a_Currency
Dim lngResult As Long

    Set oCurr = New a_Currency
    Set oCurr = oConfig.Currencies.Item(Val(lvwCurrencies.SelectedItem.Key))
    frm.Component oCurr
    frm.Show vbModal
    FillCurrencyList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.EditCurr"
End Sub

