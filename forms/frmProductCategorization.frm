VERSION 5.00
Begin VB.Form frmProductCategorizations 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Product categorizations"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   5625
   Begin VB.TextBox txtNewCategorizationValue 
      Height          =   315
      Left            =   2595
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3045
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdOK_2 
      BackColor       =   &H00C4BCA4&
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4575
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel_2 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdAddCategorizationValue 
      BackColor       =   &H00C4BCA4&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2655
      Width           =   270
   End
   Begin VB.CommandButton cmdRemoveCategorizationValue 
      BackColor       =   &H00C4BCA4&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2655
      Width           =   270
   End
   Begin VB.CommandButton cmdRemoveCategorization 
      BackColor       =   &H00C4BCA4&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   615
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   270
   End
   Begin VB.CommandButton cmdCancel_1 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Cancel"
      Height          =   315
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdOK_1 
      BackColor       =   &H00C4BCA4&
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1665
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtNewCategorization 
      Height          =   315
      Left            =   135
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3045
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.CommandButton cmdAddCategorization 
      BackColor       =   &H00C4BCA4&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   270
   End
   Begin VB.ListBox lstProductCategorizationValues 
      Height          =   2205
      Left            =   2595
      TabIndex        =   2
      Top             =   390
      Width           =   2775
   End
   Begin VB.ListBox lstProductCategorizations 
      Height          =   2205
      Left            =   180
      TabIndex        =   1
      Top             =   405
      Width           =   2265
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
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
      Left            =   2025
      Picture         =   "frmProductCategorization.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4125
      Width           =   1000
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "dbl-click to edit"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   240
      Left            =   3660
      TabIndex        =   16
      Top             =   2655
      Width           =   1290
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "dbl-click to edit"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   240
      Left            =   1140
      TabIndex        =   15
      Top             =   2640
      Width           =   1290
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Categorization values"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   2700
      TabIndex        =   14
      Top             =   150
      Width           =   2205
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Categorizations"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   225
      TabIndex        =   13
      Top             =   150
      Width           =   2205
   End
End
Attribute VB_Name = "frmProductCategorizations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oPT As a_PT
Attribute oPT.VB_VarHelpID = -1
Dim tlProductTypes As z_TextList
Dim tlProductCategorizations As z_TextList
Dim tlProductCategorizationValues As z_TextList
Dim bEditing As Boolean
Dim flgLoading As Boolean
Dim tlTmp As z_TextList


Private Sub cmdAddCategorization_Click()
    If flgLoading Then Exit Sub
    bEditing = False
    SwitchCategorizationsValuesFields False
    SwitchCategorizationsFields True
End Sub

Private Sub cmdAddCategorizationValue_Click()
    If flgLoading Then Exit Sub
    bEditing = False
    SwitchCategorizationsFields False
    SwitchCategorizationsValuesFields True
End Sub


Private Sub cmdRemoveCategorization_Click()
Dim oSQL As New z_SQL
    
    If MsgBox("You are removing the categorization: '" & lstProductCategorizations & "?. Confirm.", vbOKCancel + vbQuestion, "Confirm") = vbCancel Then Exit Sub
    oSQL.RunProc "RemoveProductCategorization", Array(tlProductCategorizations.KeyByOrdinalIndex(lstProductCategorizations.ListIndex + 1)), "Deactivating"
    txtNewCategorization = ""
    cmdOK_1.Enabled = False
    SwitchCategorizationsFields False
    flgLoading = True
    LoadtlProductCategorization
    flgLoading = False
    Exit Sub

End Sub

Private Sub cmdRemoveCategorizationValue_Click()
Dim oSQL As New z_SQL
    
    If MsgBox("You are removing the categorization value: '" & lstProductCategorizationValues & "?. Confirm.", vbOKCancel + vbQuestion, "Confirm") = vbCancel Then Exit Sub
    oSQL.RunProc "RemoveProductCategorization", Array(tlTmp.KeyByOrdinalIndex(lstProductCategorizationValues.ListIndex + 1)), "Deactivating"
    txtNewCategorizationValue = ""
    cmdOK_2.Enabled = False
    SwitchCategorizationsValuesFields False
    flgLoading = True
    LoadListbox lstProductCategorizationValues, GetTextList(lstProductCategorizations.text)
    flgLoading = False
    Exit Sub
End Sub

Private Sub lstProductCategorizations_DblClick()
    If flgLoading Then Exit Sub
    bEditing = True
    SwitchCategorizationsValuesFields False
    SwitchCategorizationsFields True
    txtNewCategorization = lstProductCategorizations
End Sub
Private Sub lstProductCategorizationValues_DblClick()
    If flgLoading Then Exit Sub
    bEditing = True
    SwitchCategorizationsFields False
    SwitchCategorizationsValuesFields True
    txtNewCategorizationValue = lstProductCategorizationValues
End Sub

Private Sub cmdCancel_1_Click()
    txtNewCategorization = ""
    SwitchCategorizationsFields False
    cmdOK_1.Enabled = False
End Sub
Private Sub cmdCancel_2_Click()
    txtNewCategorization = ""
    SwitchCategorizationsFields False
    cmdOK_2.Enabled = False
End Sub
Private Sub SwitchCategorizationsFields(bON As Boolean)
    Me.txtNewCategorization.Visible = bON
    Me.cmdCancel_1.Visible = bON
    Me.cmdOK_1.Visible = bON
End Sub
Private Sub SwitchCategorizationsValuesFields(bON As Boolean)
    Me.txtNewCategorizationValue.Visible = bON
    Me.cmdCancel_2.Visible = bON
    Me.cmdOK_2.Visible = bON
End Sub

Private Sub cmdOK_1_Click()
    On Error GoTo errHandler
Dim lngPCID As Long
Dim oSQL As New z_SQL
Dim txtRememberSelection As String
    
    If bEditing Then
        If MsgBox("You are altering the categorization: '" & lstProductCategorizations & "' to '" & txtNewCategorization & "'?. Confirm.", vbOKCancel + vbQuestion, "Confirm") = vbCancel Then Exit Sub
    End If
    txtRememberSelection = lstProductCategorizations.text
    If bEditing Then
        oSQL.RunProc "UpdateProductCategorization", Array(txtNewCategorization, tlProductCategorizations.Key(lstProductCategorizations.text)), ""
    Else
        oSQL.RunProc "InsertNewProductCategorization", Array(txtNewCategorization), ""
    End If
    txtNewCategorization = ""
    cmdOK_1.Enabled = False
    SwitchCategorizationsFields False
    flgLoading = True
    LoadtlProductCategorization
    lstProductCategorizations.text = txtRememberSelection
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductCategorizations.cmdOK_1_Click"
End Sub

Private Sub cmdOK_2_Click()
    On Error GoTo errHandler
Dim lngPCID As Long
Dim oSQL As New z_SQL
    If bEditing Then
        If MsgBox("You are altering the categorization value: '" & lstProductCategorizationValues & "' to '" & txtNewCategorizationValue & "'?. Confirm.", vbOKCancel + vbQuestion, "Confirm") = vbCancel Then Exit Sub
    End If
    If lstProductCategorizations = "" Then Exit Sub
    If bEditing Then
       oSQL.RunProc "UpdateProductCategorization", Array(txtNewCategorizationValue, tlTmp.Key(lstProductCategorizationValues)), ""
    Else
        oSQL.RunProc "InsertNewProductCategorizationValue", Array(txtNewCategorizationValue, lstProductCategorizations), ""
    End If
    txtNewCategorizationValue = ""
    cmdOK_2.Enabled = False
    SwitchCategorizationsValuesFields False
    flgLoading = True
    LoadListbox lstProductCategorizationValues, GetTextList(lstProductCategorizations.text)
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductCategorizations.cmdOK_2_Click"
End Sub




Private Sub txtNewCategorization_Change()
    cmdOK_1.Enabled = Len(txtNewCategorization) > 1
End Sub
Private Sub txtNewCategorizationValue_Change()
    cmdOK_2.Enabled = Len(txtNewCategorizationValue) > 1
End Sub

Private Function GetTextList(txt As String) As z_TextList
        Set tlTmp = New z_TextList
        tlTmp.Load ltProductCategorizationValues, CStr(tlProductCategorizations.Key(txt))
        Set GetTextList = tlTmp
End Function



Private Sub cmdClose_Click()
    Unload Me
End Sub




Private Sub Form_Unload(Cancel As Integer)
    oPC.Configuration.RefreshProductTypes
End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lstProductCategorizations_Click()
    flgLoading = True
    LoadListbox lstProductCategorizationValues, GetTextList(lstProductCategorizations.text)
    flgLoading = False
End Sub


Private Sub oPT_DBActionSTatus(Status As Integer)
    Select Case Status
    Case 22
        MsgBox "Addition of product type failed, it would have created a duplicate value"
    End Select
    
End Sub

Private Sub oPT_Inuse()
    MsgBox "This product type is being used. You should merge it with another product type " & vbCrLf _
    & "rather than deleting it." & vbCrLf _
    & "See under Tools/Utilities/Merge two product types"
End Sub

Private Sub Form_Load()
    
    flgLoading = True
    If Me.WindowState <> 2 Then
        TOP = 500
        Left = 250
        Width = 5745
        Height = 5220
    End If
    
    LoadtlProductCategorization
    flgLoading = False
End Sub
Private Sub LoadtlProductCategorization()
    Set tlProductCategorizations = New z_TextList
    tlProductCategorizations.Load ltProductCategorizations
    LoadProductCategorizationListBox
End Sub
Private Sub LoadtlProductCategorizationValues()
    Set tlProductCategorizationValues = New z_TextList
    tlProductCategorizationValues.Load ltProductCategorizationValues
    LoadProductCategorizationValuesListBox
End Sub

Private Sub LoadProductCategorizationListBox()
    LoadListbox Me.lstProductCategorizations, tlProductCategorizations
End Sub
Private Sub LoadProductCategorizationValuesListBox()
    LoadListbox Me.lstProductCategorizationValues, tlProductCategorizationValues
End Sub

Private Sub LoadList()
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

End Sub



