VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCopy 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Product copy"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Height          =   6165
      Left            =   2910
      TabIndex        =   18
      Top             =   105
      Width           =   5325
      Begin VB.TextBox txtSerialToCopyFrom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   2265
         TabIndex        =   24
         Top             =   5655
         Width           =   1125
      End
      Begin VB.TextBox txtFlagText 
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
         Height          =   690
         Left            =   270
         TabIndex        =   6
         Top             =   4875
         Width           =   4650
      End
      Begin VB.CommandButton cmdInsert 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Copy text from: "
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
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   5640
         Width           =   1950
      End
      Begin VB.TextBox txtComment 
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
         Height          =   1785
         Left            =   270
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2730
         Width           =   4890
      End
      Begin VB.TextBox txtDescription 
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
         Height          =   1980
         Left            =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   405
         Width           =   4875
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Flag text"
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
         Height          =   225
         Left            =   330
         TabIndex        =   23
         Top             =   4635
         Width           =   945
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Comment"
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
         Left            =   285
         TabIndex        =   20
         Top             =   2490
         Width           =   885
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Condition"
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
         Left            =   255
         TabIndex        =   19
         Top             =   150
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   660
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5715
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1635
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5715
      Width           =   975
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   2  'Center
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
      Left            =   210
      TabIndex        =   1
      Top             =   1980
      Width           =   1125
   End
   Begin VB.TextBox txtDateSold 
      Alignment       =   2  'Center
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
      Left            =   210
      TabIndex        =   10
      Top             =   1350
      Width           =   1125
   End
   Begin VB.TextBox txtDatePurchased 
      Alignment       =   2  'Center
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
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   1125
   End
   Begin VB.TextBox txtSerial 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Enabled         =   0   'False
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   150
      Width           =   585
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Catalogue entries"
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
      Height          =   3105
      Left            =   180
      TabIndex        =   16
      Top             =   2475
      Width           =   2460
      Begin VB.CommandButton cmdAdd 
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
         Height          =   420
         Left            =   1110
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   330
         Width           =   945
      End
      Begin VB.ComboBox cboCATAL 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   285
         TabIndex        =   2
         Top             =   345
         Width           =   825
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
         Height          =   345
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2655
         Width           =   945
      End
      Begin MSComctlLib.ListView lvwCE 
         Height          =   1890
         Left            =   105
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   750
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   3334
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cat. No."
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Price"
            Object.Width           =   2187
         EndProperty
      End
   End
   Begin VB.Label lblLocalPrice 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   1350
      TabIndex        =   17
      Top             =   1575
      Width           =   1110
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Left            =   225
      TabIndex        =   15
      Top             =   1755
      Width           =   750
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date sold"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1095
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date purchased"
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
      TabIndex        =   13
      Top             =   630
      Width           =   1440
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Copy number"
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
      Left            =   135
      TabIndex        =   12
      Top             =   195
      Width           =   1125
   End
End
Attribute VB_Name = "frmCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oCopy As a_Copy
Private oCopyFrom As a_Copy
Private flgLoading As Boolean
Dim tlCATAL As z_TextList

Public Sub component(pCopy As a_Copy, pCopyFrom As a_Copy)
    On Error GoTo errHandler
    Set oCopy = pCopy
    Set oCopyFrom = pCopyFrom
    oCopy.BeginEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.component(pCopy,pCopyFrom)", Array(pCopy, pCopyFrom)
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
    Me.txtSerial = oCopy.Serial
    Me.txtDatePurchased = oCopy.PurchaseDateF
    Me.txtDateSold = oCopy.SoldDateF
    Me.txtDescription = oCopy.Description
    Me.txtComment = oCopy.Comment
    Me.txtPrice = oCopy.PriceF
    Me.txtFlagText = oCopy.FlagText
    Me.lblLocalPrice.Caption = oCopy.LocalPriceF
    If Not oCopyFrom Is Nothing Then
        txtSerialToCopyFrom = oCopyFrom.SerialF
    Else
        txtSerialToCopyFrom = ""
        Me.cmdInsert.Enabled = False
    End If
    Set tlCATAL = New z_TextList
    tlCATAL.Load ltCatalogue
    LoadCombo Me.cboCATAL, tlCATAL
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.LoadControls"
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo errHandler
Dim oCE As a_CATALPI
    If cboCATAL = "" Then Exit Sub
    Set oCE = oCopy.CatalogueEntries.Add
    oCE.BeginEdit
    oCE.CATALID = tlCATAL.Key(Me.cboCATAL)
    oCE.Serial = cboCATAL
    oCE.Price = oCopy.Price
    oCE.ApplyEdit
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.cmdAdd_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdInsert_Click()
    On Error GoTo errHandler
Dim oLocalCopy As a_Copy
    If oCopyFrom Is Nothing Then
     '   Set oLocalCopy = oCopy.CopyBySerialNo(txtSerialToCopyFrom)
      Set oLocalCopy = oCopy.PreviousItem
    Else
        Set oLocalCopy = oCopyFrom
    End If
    If Not oLocalCopy Is Nothing Then
        oCopy.SetComment oLocalCopy.Comment
        oCopy.SetDescription oLocalCopy.Description
        oCopy.SetFlagtext oLocalCopy.FlagText
    Else
        MsgBox "There is no previous copy", , "Invalid action"
    End If
    LoadControls
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.cmdInsert_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemove_Click()
    On Error GoTo errHandler
Dim oCE As a_CATALPI
    Set oCE = oCopy.CatalogueEntries(lvwCE.SelectedItem.Key)
    oCE.BeginEdit
    oCE.Delete
    oCE.ApplyEdit
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.cmdRemove_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oCopy.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    oCopy.ApplyEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        TOP = 50
        Left = 50
        Width = 8600
        Height = 7000
    End If
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set tlCATAL = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwCE_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.lvwCE_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), EA_NORERAISE
    HandleError
End Sub

Private Sub lvwCE_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.lvwCE_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtComment_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtComment = oCopy.Comment
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtComment_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtComment_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCopy.SetComment txtComment
    If Err Then
      Beep
      intPos = txtComment.SelStart
      txtComment = oCopy.Comment
      txtComment.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtComment_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtComment_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCopy.SetComment(txtComment)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtComment_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtDescription_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtDescription = oCopy.Description
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtDescription_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDescription_Change()
    On Error GoTo errHandler
Dim intPos As Integer
Dim strTemp As String
Dim iresult As Integer

    On Error Resume Next
    oCopy.SetDescription (txtDescription)
 '   strTemp = SC.CheckText(txtCondition, iResult)
  '  txtCondition = strTemp
    If Err Then
      Beep
      intPos = txtDescription.SelStart
      txtDescription = oCopy.Description
      txtDescription.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtDescription_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDescription_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCopy.SetDescription(txtDescription)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtDescription_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtDatePurchased_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtDatePurchased
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtDatePurchased_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDatePurchased_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtDatePurchased = oCopy.PurchaseDateF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtDatePurchased_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDatePurchased_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCopy.SetPurchaseDate(txtDatePurchased)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtDatePurchased_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtDateSold_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtDateSold
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtDateSold_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtdatesold_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtDateSold = oCopy.SoldDateF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtdatesold_LostFocus", , EA_NORERAISE
    HandleError
End Sub
'Private Sub txtdatesold_Change()
'Dim intPos As Integer
'    On Error Resume Next
'    oCopy.SetSoldDate txtDateSold
'    If Err Then
'      Beep
'      intPos = Me.txtDateSold.SelStart
'      txtDateSold = oCopy.SoldDate
'      txtDateSold.SelStart = intPos - 1
'    End If
'End Sub
Private Sub txtdatesold_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCopy.SetSoldDate(txtDateSold)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtdatesold_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub LoadListView()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvwCE.ListItems.Clear
    For i = 1 To oCopy.CatalogueEntries.Count
        Set objItm = Me.lvwCE.ListItems.Add
        With objItm
            .Key = oCopy.CatalogueEntries(i).Key
            .text = oCopy.CatalogueEntries(i).Serial & IIf(oCopy.CatalogueEntries(i).IsDeleted, "(DEL)", "")
            .SubItems(1) = oCopy.CatalogueEntries(i).PriceF
        End With
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.LoadListView"
End Sub


Private Sub txtFlagText_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtFlagText = oCopy.FlagText
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtFlagText_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFlagText_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oCopy.SetFlagtext txtFlagText
    If Err Then
      Beep
      intPos = txtFlagText.SelStart
      txtFlagText = oCopy.FlagText
      txtFlagText.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtFlagText_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFlagText_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCopy.SetFlagtext(txtFlagText)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtFlagText_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtPrice_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtPrice
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtPrice_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPrice_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtPrice = oCopy.PriceF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtPrice_LostFocus", , EA_NORERAISE
    HandleError
End Sub
'Private Sub txtPrice_Change()
'Dim intPos As Integer
'    On Error Resume Next
'    oCopy.SetPrice txtPrice
'    If Err Then
'      Beep
'      intPos = txtPrice.SelStart
'      txtPrice = oCopy.PriceF
'      txtPrice.SelStart = intPos - 1
'    End If
'End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCopy.SetPrice(txtPrice)
    If Not Cancel Then
        lblLocalPrice.Caption = oCopy.LocalPriceF
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCopy.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

