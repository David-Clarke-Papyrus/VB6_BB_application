VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Begin VB.Form frmProductNonBook 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Stock"
   ClientHeight    =   6735
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11565
   ControlBox      =   0   'False
   Icon            =   "frmProductNonBook.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleMode       =   0  'User
   ScaleWidth      =   15255.96
   Begin VB.TextBox txtSpecialPrice 
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
      Left            =   1920
      TabIndex        =   28
      Top             =   5745
      Width           =   1380
   End
   Begin VB.TextBox txtCost 
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
      Left            =   1920
      TabIndex        =   27
      Top             =   5370
      Width           =   1380
   End
   Begin VB.TextBox txtSP 
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
      Left            =   1920
      TabIndex        =   26
      Top             =   4995
      Width           =   1380
   End
   Begin VB.TextBox txtRRP 
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
      Left            =   1920
      TabIndex        =   25
      Top             =   4620
      Width           =   1380
   End
   Begin VB.CommandButton cmdGenerateEAN 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&generate"
      Height          =   315
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   255
      Width           =   750
   End
   Begin VB.CommandButton cmdSetDefault 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Default V.A.T. rate"
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
      Left            =   7935
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3210
      Width           =   1755
   End
   Begin VB.TextBox txtVAT 
      Alignment       =   2  'Center
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
      Height          =   330
      Left            =   6465
      TabIndex        =   21
      Top             =   3240
      Width           =   1380
   End
   Begin VB.TextBox txtEAN 
      Alignment       =   2  'Center
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
      Height          =   330
      Left            =   6675
      TabIndex        =   15
      Top             =   255
      Width           =   2010
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
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
      Left            =   10425
      Picture         =   "frmProductNonBook.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5670
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9480
      Picture         =   "frmProductNonBook.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5685
      Width           =   930
   End
   Begin VB.TextBox txtNote 
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
      Height          =   1500
      Left            =   6690
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   735
      Width           =   2250
   End
   Begin VB.TextBox txtErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   6105
      Width           =   4350
   End
   Begin VB.CommandButton cmdNewCode 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&New code"
      Height          =   420
      Left            =   3645
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   225
      Width           =   1050
   End
   Begin VB.TextBox txtManufacturer 
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
      Left            =   1920
      TabIndex        =   4
      Top             =   2340
      Width           =   3915
   End
   Begin VB.TextBox txtPackaging 
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
      Height          =   585
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1530
      Width           =   3900
   End
   Begin VB.TextBox txtProductName 
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
      Height          =   570
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   765
      Width           =   3900
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
      Left            =   1935
      TabIndex        =   0
      Top             =   255
      Width           =   1680
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboProductType 
      Height          =   390
      Left            =   1920
      OleObjectBlob   =   "frmProductNonBook.frx":0E1E
      TabIndex        =   17
      Top             =   2865
      Width           =   3045
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboCategory 
      Height          =   390
      Left            =   1935
      OleObjectBlob   =   "frmProductNonBook.frx":2044
      TabIndex        =   18
      Top             =   3480
      Width           =   3045
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboCategoryHeading 
      Height          =   390
      Left            =   1950
      OleObjectBlob   =   "frmProductNonBook.frx":326A
      TabIndex        =   33
      Top             =   4080
      Width           =   3045
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Special"
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
      Height          =   285
      Left            =   1065
      TabIndex        =   32
      Top             =   5745
      Width           =   750
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cost"
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
      Height          =   285
      Left            =   1065
      TabIndex        =   31
      Top             =   5370
      Width           =   750
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "S.P."
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
      Height          =   285
      Left            =   1065
      TabIndex        =   30
      Top             =   5010
      Width           =   750
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "R.R.P."
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
      Height          =   285
      Left            =   1065
      TabIndex        =   29
      Top             =   4635
      Width           =   750
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "V.A.T. Rate"
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
      Height          =   285
      Left            =   5235
      TabIndex        =   23
      Top             =   3270
      Width           =   1080
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Category"
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
      TabIndex        =   20
      Top             =   3525
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Product type"
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
      Top             =   2940
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "E.A.N."
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
      Height          =   285
      Left            =   5520
      TabIndex        =   16
      Top             =   285
      Width           =   1080
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Catalogue heading"
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
      Height          =   285
      Left            =   15
      TabIndex        =   14
      Top             =   4140
      Width           =   1755
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Note"
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
      Height          =   285
      Left            =   6135
      TabIndex        =   11
      Top             =   735
      Width           =   420
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Manufacturer"
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
      TabIndex        =   9
      Top             =   2355
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Product packaging"
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
      TabIndex        =   8
      Top             =   1545
      Width           =   1605
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Product name"
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
      Left            =   345
      TabIndex        =   7
      Top             =   795
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Left            =   1170
      TabIndex        =   6
      Top             =   285
      Width           =   660
   End
End
Attribute VB_Name = "frmProductNonBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tlProductTypes As z_TextList
Private WithEvents oProd As a_Product
Attribute oProd.VB_VarHelpID = -1
Dim flgLoading As Boolean
Private tlSections As z_TextList
Dim mCancel As Boolean
Dim XA As XArrayDB
Dim frmPrevious As Form

Sub Component(pProduct As a_Product, Optional pPrevForm As Form)
Dim ar() As String
    Set frmPrevious = pPrevForm
    Set oProd = pProduct
    oProd.BeginEdit
    oProd.NonStock = False
    
    SetupCombos
    
    cboCategory.BeginUpdate
    tlSections.CollectionAsArray ar
    cboCategory.PutItems ar
    cboCategory.EndUpdate
    
    cboProductType.BeginUpdate
    tlProductTypes.CollectionAsArray ar
    cboProductType.PutItems ar
    cboProductType.EndUpdate

End Sub


Private Sub cboCategory_Click()
    If flgLoading Then Exit Sub
    oProd.SetCategoryID tlSections.Key(cboCategory)
End Sub

Private Sub cboCatHead_Click()
    If flgLoading Then Exit Sub
    oProd.SetCatalogueheadingID tlSections.Key(cboCategoryHeading)
End Sub



Private Sub cmdDelete_Click()

    If MsgBox("Confirm deletion of product: " & oProd.Title & vbCrLf & "This is permanent!", vbOKCancel + vbExclamation, "Confirm") = vbOK Then
        oProd.Delete
        oProd.ApplyEdit
        Unload Me
    End If


End Sub



Private Sub cmdSetDefault_Click()
    Me.txtVAT = oPC.Configuration.vatRate
End Sub


Private Sub Form_Initialize()
    Set tlSections = New z_TextList
    tlSections.Load ltDictionary, , dtCategory
    
    Set tlProductTypes = New z_TextList
    tlProductTypes.Load ltDictionary, , dtProductType

End Sub

Private Sub Form_Terminate()
    Set tlSections = Nothing
    Set tlProductTypes = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If oProd.IsEditing Then oProd.CancelEdit
End Sub


Private Sub oProd_Valid(strMsg As String)
    Me.txtErrors = strMsg
    Me.cmdOK.Enabled = (strMsg = "")
End Sub
Private Sub cmdCancel_Click()
    oProd.CancelEdit
    Unload Me
End Sub

Private Sub cmdNewCode_Click()
    Me.txtCode = "#"
    oProd.SetCode "#"
End Sub

Private Sub cmdOK_Click()
Dim lngResult As Long
Dim frmPreview As Form
    oProd.ApplyEdit lngResult
    If lngResult = 99 Then
        MsgBox "Invalid values - check that the code is has not been already used", , "Save failed"
    Else
        If frmPrevious Is Nothing Then
            If oPC.Configuration.AntiquarianYN Then
                Set frmPreview = New frmProductPrevAQ
            Else
                Set frmPreview = New frmProductPrev
            End If
        Else
            Set frmPreview = frmPrevious
        End If
        frmPreview.Component oProd
        frmPreview.RefreshForm
        frmPreview.Show
        
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    left = 10
    top = 10
    Width = 11500
    Height = 6800
    LoadControls
    Me.cmdNewCode.Enabled = oProd.IsNew
    oProd.GetStatus
End Sub
Private Sub LoadControls()
    flgLoading = True
    txtCode = oProd.code
    txtEAN = oProd.EAN
    txtProductName = oProd.Title
    txtPackaging = oProd.SubTitle
    txtManufacturer = oProd.Author
    txtRRP = oProd.RRPF
    txtSP = oProd.SPF
    txtCost = oProd.CostF
    txtSpecialPrice = oProd.SpecialPriceF
    Me.txtNote = oProd.Note
    Me.txtVAT = oProd.VATRateF
    flgLoading = False
End Sub


Private Sub txtCode_Validate(Cancel As Boolean)
    oProd.SetCode txtCode
End Sub

Private Sub txtEAN_LostFocus()
    If flgLoading Then Exit Sub
    txtEAN = oProd.EAN
End Sub
Private Sub txtEAN_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtEAN_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetEAN(txtEAN)
    If Err Then
      Beep
      intPos = txtEAN.SelStart
      txtEAN = oProd.EAN
      txtEAN.SelStart = intPos - 1
    End If
End Sub


Private Sub txtPackaging_LostFocus()
    If flgLoading Then Exit Sub
    txtPackaging = oProd.SubTitle
End Sub
Private Sub txtPackaging_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtPackaging_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetSubTitle(txtPackaging)
    If Err Then
      Beep
      intPos = txtPackaging.SelStart
      txtPackaging = oProd.SubTitle
      txtPackaging.SelStart = intPos - 1
    End If
End Sub


Private Sub txtProductName_LostFocus()
    If flgLoading Then Exit Sub
    txtProductName = oProd.Title
End Sub
Private Sub txtProductName_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtProductName_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetTitle(txtProductName)
    If Err Then
      Beep
      intPos = txtProductName.SelStart
      txtProductName = oProd.Title
      txtProductName.SelStart = intPos - 1
    End If
End Sub
Private Sub txtManufacturer_LostFocus()
    If flgLoading Then Exit Sub
    txtManufacturer = oProd.Author
End Sub
Private Sub txtManufacturer_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtManufacturer_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetAuthor(txtManufacturer)
    If Err Then
      Beep
      intPos = txtManufacturer.SelStart
      txtManufacturer = oProd.Author
      txtManufacturer.SelStart = intPos - 1
    End If
End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtNote_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.setnote(txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oProd.Note
      txtNote.SelStart = intPos - 1
    End If
End Sub

Private Sub txtVAT_GotFocus()
    AutoSelect txtVAT
End Sub
Private Sub txtVAT_LostFocus()
    If flgLoading Then Exit Sub
    txtVAT = oProd.VATRateToUse
End Sub
Private Sub txtVAT_Validate(Cancel As Boolean)
   If flgLoading Then Exit Sub
   Cancel = Not Not oProd.SetVat(txtVAT)
End Sub

Sub SetupCombos()
    
    cboProductType.BeginUpdate
    cboProductType.WidthList = 190
    cboProductType.HeightList = 162
    cboProductType.AllowSizeGrip = True
    cboProductType.AutoDropDown = True
    cboProductType.Columns.Add "Product type"
    cboProductType.Columns(0).Width = 190
    cboProductType.BackColorLock = Me.BackColor
    cboProductType.EndUpdate

    cboCategory.BeginUpdate
    cboCategory.WidthList = 190
    cboCategory.HeightList = 162
    cboCategory.AllowSizeGrip = True
    cboCategory.AutoDropDown = True
    cboCategory.Columns.Add "Category"
    cboCategory.Columns(0).Width = 190
    cboCategory.BackColorLock = Me.BackColor
    cboCategory.EndUpdate
End Sub

