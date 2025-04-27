VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdHocProduct 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Missing product - quick capture"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frCode 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Codes"
      ForeColor       =   &H8000000D&
      Height          =   1875
      Left            =   60
      TabIndex        =   27
      Top             =   45
      Width           =   3045
      Begin VB.TextBox txtCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000005&
         Height          =   285
         Left            =   165
         TabIndex        =   0
         Top             =   645
         Width           =   2715
      End
      Begin VB.TextBox txtEAN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00000005&
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Top             =   1320
         Width           =   2715
      End
      Begin VB.Label Label1 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Short code"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   165
         TabIndex        =   29
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D3D3CB&
         Caption         =   "EAN"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   150
         TabIndex        =   28
         Top             =   1035
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdSupplier 
      BackColor       =   &H00C4BCA4&
      Caption         =   "· · ·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5010
      Width           =   480
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Categories"
      ForeColor       =   &H8000000D&
      Height          =   2925
      Left            =   3255
      TabIndex        =   15
      Top             =   210
      Width           =   3435
      Begin VB.CommandButton cmdUP 
         BackColor       =   &H00C4BCA4&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3045
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2565
         Width           =   330
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00C4BCA4&
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2955
         Picture         =   "frmAdHocProduct.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   555
         Width           =   375
      End
      Begin VB.CommandButton cmdAddSection 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
         Height          =   405
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1080
         Width           =   750
      End
      Begin VB.ComboBox cboSection 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   150
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   570
         Width           =   2775
      End
      Begin VB.CommandButton cmdRemoveSection 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Remove"
         Height          =   405
         Left            =   2190
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1065
         Width           =   750
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   1320
         Left            =   150
         TabIndex        =   19
         Top             =   1515
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   2328
         SortKey         =   1
         View            =   3
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Section "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Priority"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   210
         TabIndex        =   18
         Top             =   315
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Product type"
      ForeColor       =   &H8000000D&
      Height          =   780
      Left            =   3345
      TabIndex        =   14
      Top             =   5745
      Visible         =   0   'False
      Width           =   3315
      Begin VB.PictureBox Picture 
         BackColor       =   &H00D3D3CB&
         Height          =   465
         Left            =   165
         ScaleHeight     =   405
         ScaleWidth      =   2895
         TabIndex        =   24
         Top             =   225
         Width           =   2955
         Begin VB.OptionButton optBook 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Book"
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   120
            TabIndex        =   26
            Top             =   75
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton optNonBook 
            BackColor       =   &H00D3D3CB&
            Caption         =   "General product"
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   1110
            TabIndex        =   25
            Top             =   75
            Width           =   1680
         End
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D7D1BF&
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5235
      Picture         =   "frmAdHocProduct.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4770
      Width           =   1455
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00000005&
      Height          =   285
      Left            =   270
      TabIndex        =   4
      Top             =   3600
      Width           =   2715
   End
   Begin VB.TextBox txtAuthor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00000005&
      Height          =   285
      Left            =   270
      TabIndex        =   3
      Top             =   2940
      Width           =   2715
   End
   Begin VB.TextBox txtDescription 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00000005&
      Height          =   285
      Left            =   270
      TabIndex        =   2
      Top             =   2265
      Width           =   2715
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboProductType 
      Height          =   330
      Left            =   240
      OleObjectBlob   =   "frmAdHocProduct.frx":0714
      TabIndex        =   5
      Top             =   4395
      Width           =   2760
   End
   Begin VB.Label lblSupplier 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   240
      TabIndex        =   22
      Top             =   4815
      Width           =   810
   End
   Begin VB.Label lblError 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   3330
      TabIndex        =   13
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Product type"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   240
      TabIndex        =   12
      Top             =   4110
      Width           =   1440
   End
   Begin VB.Label Label5 
      BackColor       =   &H00D3D3CB&
      Caption         =   "R.R.P."
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   270
      TabIndex        =   11
      Top             =   3315
      Width           =   1290
   End
   Begin VB.Label Label4 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Author"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   270
      TabIndex        =   10
      Top             =   2655
      Width           =   1290
   End
   Begin VB.Label Label3 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Description"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   270
      TabIndex        =   9
      Top             =   1980
      Width           =   1290
   End
End
Attribute VB_Name = "frmAdHocProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flgLoading As Boolean
Dim lngPTID As Long
Dim WithEvents oProd As a_Product
Attribute oProd.VB_VarHelpID = -1
Dim strCode As String
Dim oCurr As a_Currency
Dim mSupplierID As Long
Dim mSupplierName As String
Dim bCancelled As Boolean
Private Sub cmdRefresh_Click()
    oPC.Configuration.ReloadCategories
    LoadCombo cboSection, oPC.Configuration.Sections
    If cboSection.ListCount > 0 Then
        cboSection.ListIndex = 0
    Else
        cboSection.ListIndex = -1
    End If
End Sub
#If H_IsFRONTDESK = 0 Then
Private Sub cmdSupplier_Click()
    On Error GoTo errHandler
Dim frm As New frmBrowseSUppliers2
    frm.Show vbModal
    If frm.SupplierID > 0 Then
        oProd.SupplierID = frm.SupplierID
        oProd.LastSupplierName = frm.SupplierName
        Me.lblSupplier = oProd.LastSupplierName
    Else
        MsgBox "No supplier selected.", vbOKOnly, "Warning"
    End If
    Unload frm

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.cmdSupplier_Click", , EA_NORERAISE
    HandleError
End Sub
#End If

Private Sub cmdUP_Click()
    On Error GoTo errHandler
    If oProd.ProductSections.Key(lvw.SelectedItem) <> oPC.Configuration.WebExportID And _
            InStr(1, lvw.SelectedItem, "Multibuy") = 0 Then
        oProd.ProductSections.Mark oProd.ProductSections.Key(lvw.SelectedItem)
        LoadPSECs
    Else
        MsgBox "You cannot assign a priority category to the multibuy category.", vbInformation, "Can't do this"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.cmdUP_Click", , EA_NORERAISE
    HandleError

End Sub

Private Sub Form_Load()
    AutoSizeDropDownWidth Me.cboSection

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
  
    If oProd.IsEditing Then
        If MsgBox("You want to close without creating a new product?", vbQuestion + vbYesNo, "Closing form") = vbNo Then
            Cancel = True
        Else
            oProd.CancelEdit
            bCancelled = True
            Me.Hide

        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Public Property Get IsCancelled() As Boolean
    IsCancelled = bCancelled
End Property





Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub oProd_RedisplayCodes()
    Me.txtCode.text = oProd.code
    Me.txtEAN.text = oProd.EAN
End Sub

Private Sub oProd_Valid(pMsg As String)
    On Error GoTo errHandler
    cmdSave.Enabled = (pMsg = "")
    lblError.Caption = pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.oProd_Valid(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub
Public Sub component(pCode As String, Optional pCurrency As a_Currency, Optional SupplierID As Long, Optional SupplierName As String)
    On Error GoTo errHandler
Dim oProdCode As New z_ProdCode

            bCancelled = False
    mSupplierID = SupplierID
    mSupplierName = SupplierName
    If mSupplierID > 0 Then
      oProd.SupplierID = mSupplierID
      oProd.LastSupplierName = mSupplierName
      lblSupplier.Caption = mSupplierName
    End If
    pCode = FNS(pCode)
    If IsISBN13(pCode) Then
        If oProdCode.LoadNew("", pCode, True, , , True) = True Then
            txtEAN = oProdCode.ISBN13
            txtCode = oProdCode.code
            oProd.SetEAN oProdCode.ISBN13
            oProd.SetCode oProdCode.code
        End If
    Else
        If oProdCode.LoadNew(pCode, "", True, , , True) = True Then
            txtEAN = oProdCode.ISBN13
            txtCode = oProdCode.code
            oProd.SetEAN oProdCode.ISBN13
            oProd.SetCode oProdCode.code
        Else
            txtCode = oProdCode.code
        End If
    
    End If
'    txtCode.Locked = True
'    txtEAN.Locked = True
    If Not pCurrency Is Nothing Then
        Set oCurr = pCurrency
    Else
        Set oCurr = oPC.Configuration.DefaultCurrency
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.component(pCODE,pCurrency)", Array(pCode, pCurrency)
End Sub

Private Sub SetupPT()
    On Error GoTo errHandler
    cboProductType.BeginUpdate
    cboProductType.WidthList = 190
    cboProductType.HeightList = 162
    cboProductType.AllowSizeGrip = True
    cboProductType.AutoDropDown = True
    cboProductType.SelForeColor = vbRed
    cboProductType.Columns.Add "Product type"
    cboProductType.Columns.Add "Seesafe"
    cboProductType.Columns(0).Width = 190
    cboProductType.Columns(1).Width = 0
    cboProductType.BackColorLock = Me.BackColor
    cboProductType.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.SetupPT"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errHandler
Dim oPT As New a_PT
    Dim strTemp As String
    If oProd.ValidateObject("") = False Then Exit Sub
    If oProd.RRP = 0 Then
        If MsgBox("The price has been set to 0. Do you want to continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    If Not oCurr Is Nothing Then
        oProd.ForeignOrderedCURRID = oCurr.ID
    End If
    If lngPTID > 0 Then
        oPT.Load lngPTID
        oProd.SetSPFROMRRP oPT
    End If
    oProd.SetProductTypeID lngPTID
    oProd.VATRate = oPC.Configuration.VATRate
    oProd.SetProductType "B"
    oProd.ApplyEdit , strTemp
    If strTemp > "" Then
        MsgBox "An error occurred while saving this product, the message is: " & vbCrLf & strTemp
    Else
        strCode = oProd.code
        Me.Hide
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
Dim ar() As String
Dim lngTemp As Long
Dim strPos As String

    flgLoading = True
    SetupPT
    ReDim ar(5)
    cboProductType.BeginUpdate
    oPC.Configuration.ProductTypes.CollectionAsArray ar
    cboProductType.PutItems ar
    cboProductType.EndUpdate
    lngPTID = oPC.Configuration.DefaultPT
    lngTemp = cboProductType.Items.FindItem(oPC.Configuration.ProductTypes.Item(lngPTID), 0)
    If lngTemp > 0 Then
        On Error Resume Next
        cboProductType.Items.SelectItem(cboProductType.Items.FindItem(oPC.Configuration.ProductTypes.Item(lngPTID), 0)) = True
        On Error GoTo errHandler
    End If
    Set oProd = New a_Product
    oProd.BeginEdit

    oProd.VATRate = oPC.Configuration.VATRate
    oProd.SetRRP "0"
    oProd.SetSP "0"
    oProd.SetProductTypeID lngPTID
    If optBook Then
        oProd.SetProductType "B"
    ElseIf Me.optNonBook Then
        oProd.SetProductType "G"
    End If
    LoadCombo cboSection, oPC.Configuration.Sections_Short
    LoadPSECs
    RestrictCustomerTypes

    flgLoading = False

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.Form_Initialize", , EA_NORERAISE, , "line number", Array(Erl())
    HandleError
End Sub

Private Sub cboProductType_SelectionChanged()
    On Error GoTo errHandler
    lngPTID = oPC.Configuration.ProductTypes.Key(cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.cboProductType_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub


Private Sub optBook_Click()
    On Error GoTo errHandler
    oProd.SetProductType "B"
    oProd.SetEAN txtEAN     'to force validation after product type change
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.optBook_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optNonBook_Click()
    On Error GoTo errHandler
    oProd.SetProductType "G"
    oProd.SetEAN txtEAN     'to force validation after product type change
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.optNonBook_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtAuthor_Validate(Cancel As Boolean)
Dim intPos As Integer
    On Error GoTo errHandler
    On Error Resume Next
    oProd.SetAuthor txtAuthor
    If Err Then
      Beep
      intPos = txtAuthor.SelStart
      txtAuthor = oProd.Author
      txtAuthor.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.txtAuthor_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtCode_Validate(Cancel As Boolean)
10        On Error GoTo errHandler
      Dim bOK As Boolean
          
20        bOK = oProd.SetCode(txtCode)
            
30          '6oProd.mobjValid.v ("CODE"
'40          MsgBox "The code you entered is invalid." & vbCrLf _
'                      & "Make sure you are not entering an ISBN-13 or EAN number in this field. They belong in the next box.", vbExclamation + vbOKOnly, "Warning"
'
'50              Cancel = True
'60              Exit Sub
'70          End If
'80        If Len(oProd.code) > 10 Then
'90            MsgBox "The code you entered is greater than ten characters long. This is possibly an error." & vbCrLf _
'                      & "Make sure you are not entering an ISBN-13 or EAN number in this field. They belong in the next box.", vbExclamation + vbOKOnly, "Warning"
'100       End If
110       Exit Sub
errHandler:
120       If ErrMustStop Then Debug.Assert False: Resume
130       ErrorIn "frmAdHocProduct.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE, , "Line number,Txtcode", Array(Erl(), txtCode)
140       HandleError
End Sub

Public Property Get code() As String
    code = oProd.EAN
End Property
Private Sub txtDescription_Validate(Cancel As Boolean)
Dim intPos As Integer
    On Error GoTo errHandler
    On Error Resume Next
    oProd.SetTitle txtDescription
    If Err Then
      Beep
      intPos = txtDescription.SelStart
      txtDescription = oProd.Title
      txtDescription.SelStart = intPos - 1
    End If
    Cancel = (txtDescription = "")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.txtDescription_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtEAN_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oProd.SetEAN txtEAN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.txtEAN_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtPrice_LostFocus()
    On Error GoTo errHandler
  '  txtPrice = oProd.RRPF
    If App.Title = "PBKS Manager" Then
        txtPrice = oProd.RRPF
    Else
        txtPrice = oProd.SPF
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.txtPrice_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If App.Title = "PBKS Manager" Then
        Cancel = Not oProd.SetRRP(txtPrice, oCurr)
        Cancel = Not oProd.SetSP(txtPrice, oCurr)
    Else
        Cancel = Not oProd.SetSP(txtPrice, oCurr)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub LoadPSECs()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Long
    
    lvw.ListItems.Clear
    For i = 1 To oProd.ProductSections.Count
        Set lstItem = lvw.ListItems.Add
        With oProd.ProductSections(i)
            lstItem.text = .Description
            If lstItem.Key = "" Then lstItem.Key = .Key
            lstItem.SubItems(1) = .PriorityF
        End With
    Next i
    
EXIT_Handler:
    Set lstItem = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.LoadPSECs"
End Sub


Private Sub cmdAddSection_Click()
    On Error GoTo errHandler
Dim oPSEC As New a_ProductSection
    If flgLoading Then Exit Sub
    If cboSection.ListIndex < 0 Then
        MsgBox "You must choose a section.", vbInformation, "Can't do this"
        Exit Sub
    End If
    If cboSection = "" Then
        MsgBox "You cannot add an empty section description.", vbInformation, "Can't do this"
        Exit Sub
    End If
    If InStr(1, cboSection, "Unallocated") > 0 Then
        MsgBox "You cannot add to the 'Unallocated' section.", vbInformation, "Can't do this"
        Exit Sub
    End If
    
    Set oPSEC = oProd.ProductSections.Add
    oPSEC.PID = oProd.PID
    oPSEC.SECID = oPC.Configuration.Sections.Key(cboSection)
    oPSEC.Description = cboSection
    If oProd.ProductSections.Count = 0 Or oProd.ProductSections.Count = 1 And oProd.MultibuyCode > "" Then
        oPSEC.Priority = 99
        oProd.MasterCategory = oPSEC.SECID
    End If
    oPSEC.ApplyEdit
    oPSEC.BeginEdit
    cboSection.RemoveItem cboSection.ListIndex
    If cboSection.ListCount > 0 Then
        cboSection.ListIndex = 0
    Else
        cboSection.ListIndex = -1
    End If
    LoadPSECs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.cmdAddSection_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdRemoveSection_Click()
Dim Res As Boolean

    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If lvw.ListItems.Count = 0 Then Exit Sub
    
    If Not oProd.ProductSections.Remove(oProd.ProductSections.Key(lvw.SelectedItem)) Then
        MsgBox "Cannot remove this category assignment, possibly it is the master category. First assign a new master category.", vbInformation + vbOKOnly, "Can't do this"
        Exit Sub
    End If
    If oPC.Configuration.Sections.Key(lvw.SelectedItem) <> 0 Then   'only if not a 'system' category like 'for web export'
        cboSection.AddItem lvw.SelectedItem
        cboSection.ListIndex = 0
        LoadPSECs
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.cmdRemoveSection_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub RestrictCustomerTypes()
    On Error GoTo errHandler
Dim oPSEC As a_ProductSection
Dim i As Integer

    For Each oPSEC In oProd.ProductSections
        For i = cboSection.ListCount To 1 Step -1
            cboSection.ListIndex = i - 1
            If oPSEC.Description = cboSection Then
                cboSection.RemoveItem cboSection.ListIndex
            End If
        Next
    Next
    If cboSection.ListCount > 0 Then
        cboSection.ListIndex = 0
    Else
        cboSection.ListIndex = -1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAdHocProduct.RestrictCustomerTypes"
End Sub

