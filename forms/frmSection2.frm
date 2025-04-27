VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSection2 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Allocate product to category"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5085
   StartUpPosition =   1  'CenterOwner
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
      Left            =   4380
      Picture         =   "frmSection2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1440
      Width           =   375
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1440
      Left            =   345
      TabIndex        =   8
      Top             =   2280
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   2540
      View            =   3
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
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Priority"
         Object.Width           =   2540
      EndProperty
   End
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
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3420
      Width           =   330
   End
   Begin VB.CommandButton cmdRemoveSection 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Remove"
      Height          =   315
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1920
      Width           =   750
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Picture         =   "frmSection2.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3780
      Width           =   1000
   End
   Begin VB.ComboBox cboSection 
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
      ItemData        =   "frmSection2.frx":0714
      Left            =   360
      List            =   "frmSection2.frx":0716
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1425
      Width           =   4020
   End
   Begin VB.CommandButton cmdAddSection 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Add"
      Height          =   315
      Left            =   345
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1905
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BIC classification"
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
      Left            =   375
      TabIndex        =   6
      Top             =   90
      Width           =   2490
   End
   Begin VB.Label lblBIC 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   690
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   4395
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "category"
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
      Left            =   465
      TabIndex        =   1
      Top             =   1170
      Width           =   1080
   End
End
Attribute VB_Name = "frmSection2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oProd As a_Product
Dim flgLoading As Boolean

Public Sub component(pProd As a_Product)
    Set oProd = pProd
    oProd.BeginEdit
End Sub

Private Sub cmdClose_Click()
       Unload Me
End Sub

Private Sub cmdRefresh_Click()

    oPC.Configuration.ReloadCategories
    LoadCombo cboSection, oPC.Configuration.Sections_Short
    If cboSection.ListCount > 0 Then
        cboSection.ListIndex = 0
    Else
        cboSection.ListIndex = -1
    End If

End Sub

Private Sub cmdUP_Click()
    If oProd.ProductSections.Key(lvw.SelectedItem) <> oPC.Configuration.WebExportID And _
            InStr(1, lvw.SelectedItem, "Multibuy") = 0 Then
        oProd.ProductSections.Mark oProd.ProductSections.Key(lvw.SelectedItem)
        LoadPSECs
    Else
        MsgBox "You cannot assign a priority category to the multibuy category.", vbInformation, "Can't do this"
    End If
End Sub

Private Sub Form_Load()
    flgLoading = True
    LoadCombo cboSection, oPC.Configuration.Sections_Short
    LoadPSECs
    RestrictCustomerTypes
    lblBIC.Caption = Replace(oProd.BICDescription, "&", "&&")
   ' LoadcboSect
    flgLoading = False
End Sub

Private Sub LoadPSECs()
Dim lstItem As ListItem
Dim i As Long
    lvw.ListItems.Clear
    For i = 1 To oProd.ProductSections.Count
        Set lstItem = lvw.ListItems.Add
        With oProd.ProductSections(i)
            lstItem.text = .Description
            If lstItem.Key = "" Then lstItem.Key = .Key
            lstItem.SubItems(1) = IIf(.Priority = 99, "***", "")

        End With

    Next i
        lvw.Sorted = True
        lvw.SortKey = 1
        lvw.SortOrder = lvwDescending
EXIT_Handler:
    Set lstItem = Nothing
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
    ErrorIn "frmProduct.cmdAddSection_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdRemoveSection_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If lvw.ListItems.Count = 0 Then Exit Sub
    If lvw.ListItems.Count = 0 Then Exit Sub
    If Not oProd.ProductSections.Remove(oProd.ProductSections.Key(lvw.SelectedItem)) Then
        MsgBox "Cannot remove this category assignment, possibly it is the master category. First assign a new master category.", vbInformation + vbOKOnly, "Can't do this"
        Exit Sub
    End If
    If oPC.Configuration.Sections.Key(lvw.SelectedItem) <> 0 Then   'only if not a 'system' category like 'for web export'
        oProd.ProductSections.Remove oProd.ProductSections.Key(lvw.SelectedItem)
        cboSection.AddItem lvw.SelectedItem
        cboSection.ListIndex = 0
        LoadPSECs
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdRemoveSection_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub RestrictCustomerTypes()
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oProd.ApplyEdit
End Sub

'Private Sub LoadcboSect()
'Dim ar() As String
'    cboSect.BeginUpdate
'    cboSect.WidthList = 190
'    cboSect.HeightList = 162
'    cboSect.AllowSizeGrip = True
'    cboSect.AutoDropDown = True
'    cboSect.SelForeColor = vbBlue
'    cboSect.Columns.Add "Section"
'    cboSect.Columns.Add "Seesafe"
'    cboSect.Columns(0).Width = 190
'    cboSect.Columns(1).Width = 0
'    cboSect.BackColorLock = Me.BackColor
'    cboSect.EndUpdate
'
'
'
'    cboSect.BeginUpdate
'    oPC.Configuration.Sections_Short.CollectionAsArray ar
'    cboSect.PutItems ar
'    cboSect.EndUpdate
'
'End Sub
