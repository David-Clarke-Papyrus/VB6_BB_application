VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDictionary 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Dictionary maintenance"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   4665
   Begin VB.CommandButton cmdDefault 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Set as default"
      Height          =   375
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2925
      Width           =   1605
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   3510
      Picture         =   "frmDictionary.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1000
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2115
      Left            =   105
      TabIndex        =   1
      Top             =   765
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   3731
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Description"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Short code"
         Object.Width           =   2187
      EndProperty
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Remove"
      Height          =   375
      Left            =   3495
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2925
      Width           =   1005
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00D3D3CB&
      Height          =   1335
      Left            =   90
      TabIndex        =   6
      Top             =   3330
      Width           =   4425
      Begin VB.CheckBox chkActive 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Active"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   195
         TabIndex        =   4
         Top             =   840
         Value           =   1  'Checked
         Width           =   840
      End
      Begin VB.TextBox txtShort 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   420
         Width           =   990
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
         Height          =   390
         Left            =   3195
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   810
         Width           =   1095
      End
      Begin VB.TextBox txtDESC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1185
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   420
         Width           =   3120
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   1350
         TabIndex        =   12
         Top             =   210
         Width           =   1110
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   210
         TabIndex        =   11
         Top             =   195
         Width           =   930
      End
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frmDictionary.frx":038A
      Left            =   90
      List            =   "frmDictionary.frx":038C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   4410
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   135
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oUtil As z_UTIL
Dim tlType As z_TextList
Dim lngKey As Long
Dim lngType As DictionaryType
Dim tlDictTypes As z_TextList
Dim strSystemName As String

Private Sub LoadList()
Dim i As Long
Dim lstItem As ListItem
    Set tlType = Nothing
    Set tlType = New z_TextList
    Select Case tlDictTypes.f3(tlDictTypes.Key(cboType))
    Case "IG"
        tlType.Load ltInterestGroupAll
    Case "CT"
        tlType.Load ltCustomerTypeAll
    Case "ST"
        tlType.Load ltSupplier
    Case "SE"
        tlType.Load ltSectionsAll
    Case "SR"
        tlType.Load ltSRAll
    Case "DS"
        tlType.Load ltDispatchModes
    Case "TB"
        tlType.Load ltTextBite
    Case "PS"
        tlType.Load ltProductStatus
    Case "OA"
        tlType.Load ltCOActionCode
    End Select
    lvw.ListItems.Clear
    For i = 1 To tlType.Count
        Set lstItem = lvw.ListItems.Add
        With lstItem
            .text = tlType.ItemByOrdinalIndex(i)
            .SubItems(1) = tlType.f3ByOrdinalIndex(i)
        End With
    Next

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDefault_Click()
    If UCase(cboType) = UCase("Category") Then
    
    ElseIf UCase(cboType) = UCase("Customer type") Then
    
    End If
End Sub

Private Sub lvw_AfterLabelEdit(Cancel As Integer, NewString As String)
    Cancel = True

End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub Lvw_DblClick()
    If tlType.f4ByOrdinalIndex(lvw.SelectedItem.Index) > "" Then
        MsgBox "You can't edit s system-defined item.", vbInformation, "Can't do this"
        Exit Sub
    End If
    If lvw.ListItems.Count = 0 Then Exit Sub
    If lvw.SelectedItem.Index < 1 Then Exit Sub
    txtDESC = tlType.ItemByOrdinalIndex(lvw.SelectedItem.Index)
    txtShort = tlType.f3ByOrdinalIndex(lvw.SelectedItem.Index)
    Me.chkActive = IIf(tlType.ActiveByOrdinal(lvw.SelectedItem.Index), 1, 0)
    lngKey = tlType.Key(lvw.SelectedItem.text)
    cmdAdd.Caption = "Update"
End Sub

Private Sub cboType_Click()
 '   Me.lbDict.Clear
    lvw.ListItems.Clear
    strSystemName = tlDictTypes.f3(tlDictTypes.Key(cboType))
    LoadList ' cboType    '.ItemData(cboType.ListIndex) + 1
    cmdDefault.Enabled = (UCase(cboType) = UCase("Category") Or UCase(cboType) = UCase("Customer type"))
    If (UCase(cboType) = "TEXT BITES") Then
        txtDESC.Height = 1200
        cmdAdd.TOP = 2115
        fr1.Height = 2625
        cmdclose.TOP = 6090
        Me.Height = 7485
    Else
        txtDESC.Height = 285
        cmdAdd.TOP = 810
        fr1.Height = 1425
        cmdclose.TOP = 4860
        Me.Height = 6075
    End If
End Sub

Private Sub cmdAdd_Click()
Dim oUtil As z_UTIL
    If txtDESC = "" Or txtShort = "" Then Exit Sub
    Set oUtil = New z_UTIL
    If cmdAdd.Caption = "Update" Then
        oUtil.UpdateDictRow txtDESC, txtShort, strSystemName, lngKey, Me.chkActive = 1      'cboType.ItemData(cboType.ListIndex), lngKey
    Else
        oUtil.AddDictRow txtDESC, txtShort, strSystemName  'cboType.ItemData(cboType.ListIndex)
    End If
    LoadList
    txtDESC = ""
    txtShort = ""
    Me.cmdAdd.Caption = "Add"
    Me.cmdAdd.Enabled = False
    oPC.Configuration.RefreshSections
End Sub

Private Sub cmdRemove_Click()
    If MsgBox("You want to remove " & lvw.SelectedItem.text, vbOKCancel + vbQuestion, "Confirm") = vbCancel Then
        Exit Sub
    End If
    Set oUtil = New z_UTIL
    oUtil.DeleteDictRow tlType.Key(lvw.SelectedItem.text)
    LoadList
    txtDESC = ""
    txtShort = ""
    Me.cmdAdd.Caption = "Add"
    Me.cmdAdd.Enabled = False
    oPC.Configuration.RefreshSections
End Sub

Private Sub Form_Initialize()
    Set tlDictTypes = New z_TextList
    tlDictTypes.Load ltDictTypesFiltered
End Sub

Private Sub Form_Load()
    If Me.WindowState <> 2 Then
        TOP = 100
        Left = 100
        Width = 4850
        Height = 6500
    End If
    LoadCombo Me.cboType, tlDictTypes
End Sub

Private Sub Form_Terminate()
    Set tlDictTypes = Nothing
End Sub

'Private Sub lbDict_DblClick()
'    Me.txtDESC = Me.lbDict.List(lbDict.ListIndex)
'    Me.cmdAdd.Caption = "Update"
'    lngKey = tlType.Key(lbDict.List(lbDict.ListIndex))
'End Sub



Private Sub txtDESC_Change()
    cmdAdd.Enabled = (Len(txtDESC) > 0)
End Sub
