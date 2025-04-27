VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Begin VB.Form frmMergePTs 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Merge product types"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin EXCOMBOBOXLibCtl.ComboBox cboPTFrom 
      Height          =   315
      Left            =   285
      OleObjectBlob   =   "frmMergePT.frx":0000
      TabIndex        =   9
      Top             =   525
      Width           =   2895
   End
   Begin VB.CommandButton cmdMerge 
      BackColor       =   &H00C4BCA4&
      Caption         =   "MERGE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4350
      Picture         =   "frmMergePT.frx":13AA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1410
      Width           =   1000
   End
   Begin VB.CommandButton cmdCountTo 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Count"
      Height          =   390
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   510
      Width           =   810
   End
   Begin VB.CommandButton cmdCountFrom 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Count"
      Height          =   390
      Left            =   3405
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   540
      Width           =   810
   End
   Begin VB.CommandButton cmdDeleteUnusedPTs 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Delete all unused product types"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6390
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1980
      Width           =   3045
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboPTTo 
      Height          =   315
      Left            =   5670
      OleObjectBlob   =   "frmMergePT.frx":1734
      TabIndex        =   10
      Top             =   555
      Width           =   2895
   End
   Begin VB.Label lblPTTo 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   750
      Left            =   5505
      TabIndex        =   6
      Top             =   960
      Width           =   3090
   End
   Begin VB.Label lblPTFrom 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   750
      Left            =   285
      TabIndex        =   5
      Top             =   1050
      Width           =   3090
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "this product type "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   6105
      TabIndex        =   4
      Top             =   165
      Width           =   2505
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "into . . ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   3780
      TabIndex        =   3
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Merge this product type . . ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   495
      TabIndex        =   2
      Top             =   195
      Width           =   3120
   End
End
Attribute VB_Name = "frmMergePTs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngPTFrom As Long
Dim lngCntPTFrom As Long
Dim lngPTTo As Long
Dim lngCntPTTo As Long
Dim strPTFrom As String
Dim strPTTo As String
Dim tlFrom As z_TextList

Private Sub SetupCBOs()
    cboPTFrom.BeginUpdate
    cboPTFrom.WidthList = 190
    cboPTFrom.HeightList = 162
    cboPTFrom.AllowSizeGrip = True
    cboPTFrom.AutoDropDown = True
    cboPTFrom.SelForeColor = vbRed
    cboPTFrom.Columns.Add "Product type"
    cboPTFrom.Columns.Add "Seesafe"
    cboPTFrom.Columns(0).Width = 190
    cboPTFrom.Columns(1).Width = 0
    cboPTFrom.BackColorLock = Me.BackColor
  '  cboPTFrom.style = DropDownList
    cboPTFrom.EndUpdate
    
    cboPTTo.BeginUpdate
    cboPTTo.WidthList = 190
    cboPTTo.HeightList = 162
    cboPTTo.AllowSizeGrip = True
    cboPTTo.AutoDropDown = True
    cboPTTo.SelForeColor = vbRed
    cboPTTo.Columns.Add "Product type"
    cboPTTo.Columns.Add "Seesafe"
    cboPTTo.Columns(0).Width = 190
    cboPTTo.Columns(1).Width = 0
    cboPTTo.BackColorLock = Me.BackColor
   ' cboPTFrom.style = DropDownList
    cboPTTo.EndUpdate
End Sub


Private Sub cmdDeleteUnusedPTs_Click()
Dim oSM As New z_StockManager
    If MsgBox("You want to delete all unused product types?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    oSM.DeleteUnusedPTs
    ReloadCombos
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdMerge_Click()
Dim oSM As New z_StockManager
    If MsgBox("You want to reallocate all products that are of product type: " & strPTFrom & " to product type: " & strPTTo & "?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    oSM.MergePTs lngPTFrom, lngPTTo
    Screen.MousePointer = vbDefault
    MsgBox "The Merge has completed"
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Initialize()
Dim ar() As String
Dim arFrom() As String

    Set tlFrom = New z_TextList
    tlFrom.Load ltProductType
    tlFrom.CollectionAsArray arFrom
    oPC.Configuration.ProductTypes.CollectionAsArray ar
    SetupCBOs
    cboPTFrom.BeginUpdate
    cboPTFrom.PutItems arFrom
    cboPTFrom.EndUpdate

    cboPTTo.BeginUpdate
    cboPTTo.PutItems ar
    cboPTTo.EndUpdate

End Sub

Private Sub cmdCountFrom_Click()
Dim tmp As Long
Dim oSM As New z_StockManager
    If cboPTFrom.Items.SelectCount = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    tmp = tlFrom.Key(cboPTFrom.Items.CellCaption(cboPTFrom.Items.SelectedItem, 0))
   ' If tmp = lngPTFrom Then Exit Sub
    lngPTFrom = tmp
    lngCntPTFrom = oSM.CountProductsPerPTID(lngPTFrom)
    lblPTFrom = lngCntPTFrom & " products"
    strPTFrom = cboPTFrom.Items.CellCaption(cboPTFrom.Items.SelectedItem, 0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub cboPTTo_SelectionChanged()
    lblPTTo = ""
End Sub
Private Sub cboPTFROM_SelectionChanged()
    lblPTFrom = ""
End Sub

Private Sub cmdCountTo_Click()
Dim tmp As Long
Dim oSM As New z_StockManager

If cboPTTo.Items.SelectCount = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    tmp = oPC.Configuration.ProductTypes.Key(cboPTTo.Items.CellCaption(cboPTTo.Items.SelectedItem, 0))
  '  If tmp = lngPTTo Then Exit Sub
    lngPTTo = tmp
    lngCntPTTo = oSM.CountProductsPerPTID(lngPTTo)
    lblPTTo = lngCntPTTo & " products"
    strPTTo = cboPTTo.Items.CellCaption(cboPTTo.Items.SelectedItem, 0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub ReloadCombos()
Dim ar() As String
Dim arFrom() As String
    
    Set tlFrom = Nothing
    Set tlFrom = New z_TextList
    tlFrom.Load ltProductType
    tlFrom.CollectionAsArray arFrom

    oPC.Configuration.RefreshProductTypes
    oPC.Configuration.ProductTypes.CollectionAsArray ar
    SetupCBOs
    cboPTFrom.Items.RemoveAllItems
    cboPTFrom.BeginUpdate
    cboPTFrom.PutItems arFrom
    cboPTFrom.EndUpdate

    cboPTTo.Items.RemoveAllItems
    cboPTTo.BeginUpdate
    cboPTTo.PutItems ar
    cboPTTo.EndUpdate
    lblPTFrom.Caption = ""
    lblPTTo.Caption = ""
End Sub
