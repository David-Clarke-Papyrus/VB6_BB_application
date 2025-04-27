VERSION 5.00
Begin VB.Form frmStore 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Store"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   Icon            =   "frmStore.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   7125
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkExternal 
      BackColor       =   &H00D3D3CB&
      Caption         =   "External store"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   1035
      TabIndex        =   15
      Top             =   1095
      Width           =   1485
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
      Left            =   6000
      Picture         =   "frmStore.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4935
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
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
      Height          =   615
      Left            =   4980
      Picture         =   "frmStore.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4935
      Width           =   1000
   End
   Begin VB.CheckBox chkActive 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Active"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   2745
      TabIndex        =   11
      Top             =   1095
      Width           =   870
   End
   Begin VB.TextBox txtSystem 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4935
      TabIndex        =   2
      Top             =   1080
      Width           =   1395
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   5475
      TabIndex        =   1
      Top             =   630
      Width           =   840
   End
   Begin VB.TextBox txtDelAddress 
      Appearance      =   0  'Flat
      Height          =   1515
      Left            =   1605
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3270
      Width           =   5415
   End
   Begin VB.TextBox txtBillAddress 
      Appearance      =   0  'Flat
      Height          =   1515
      Left            =   1605
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1635
      Width           =   5415
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2115
      TabIndex        =   0
      Top             =   165
      Width           =   4200
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "These addresses are used on purchase orders"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   60
      TabIndex        =   14
      Top             =   4890
      Width           =   4470
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "System"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   3660
      TabIndex        =   10
      Top             =   1125
      Width           =   1110
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Store code (exactly 3 uppercase chars)"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   2055
      TabIndex        =   9
      Top             =   690
      Width           =   3285
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery address"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   3255
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Billing address"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   270
      TabIndex        =   7
      Top             =   1620
      Width           =   1245
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   720
      Left            =   150
      TabIndex        =   6
      Top             =   4260
      Width           =   1650
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Store name"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   885
      TabIndex        =   5
      Top             =   210
      Width           =   1110
   End
End
Attribute VB_Name = "frmStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents oStore As a_Store
Attribute oStore.VB_VarHelpID = -1
Dim flgLoading As Boolean
Private Sub oStore_Valid(pMsg As String)
    EnableOK pMsg = ""
    lblErrors = pMsg
End Sub
Private Sub EnableOK(pOK As Boolean)
    Me.cmdOK.Enabled = pOK
End Sub
'Private Sub oStore_Valid(pErrors As String, status As Boolean)
'    EnableOK status
'    lblErrors = pErrors
'End Sub

Public Sub Component(poStore As a_Store)
    Set oStore = poStore
    oStore.GetStatus
End Sub
Private Sub LoadControls()
    flgLoading = True
    txtName = oStore.Description
    txtBillAddress = oStore.BillAddress
    txtDelAddress = oStore.DelAddress
    txtCode = oStore.code
    txtSystem = oStore.SystemName
    Me.chkActive = IIf(oStore.IsActive, 1, 0)
  '  Me.chkExternal = IIf(oStore.IsExternal, 1, 0)
    flgLoading = False
End Sub
Private Sub cmdCancel_Click()
    oStore.CancelEdit
    oStore.BeginEdit
    Unload Me
End Sub


Private Sub cmdOK_Click()
Dim lngResult As Long
    If oStore.StoreIndexClashes = True Then
        MsgBox "This store code has already been used for another store. This record cannot be saved.", vbOKOnly, "Can't do this"
        Exit Sub
    End If

    oStore.ApplyEdit
    oStore.BeginEdit
    Unload Me
End Sub

Private Sub Form_Load()
    LoadControls
End Sub

Private Sub chkActive_Click()
    oStore.SetActive chkActive = 1
End Sub
'Private Sub chkExternal_Click()
'    oStore.SetExternal chkExternal = 1
'End Sub


Private Sub txtName_Change()
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oStore.SetDescription txtName
    If Err Then
      Beep
      intPos = txtName.SelStart
      txtName = oStore.Description
      txtName.SelStart = intPos - 1
    End If
    
End Sub

Private Sub txtName_GotFocus()
    AutoSelect Controls("txtName")
End Sub

Private Sub txtName_LostFocus()
   txtName.Text = oStore.Description
End Sub

Private Sub txtSystem_Change()
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oStore.SetSystemName txtSystem
    If Err Then
      Beep
      intPos = txtSystem.SelStart
      txtSystem = oStore.SystemName
      txtSystem.SelStart = intPos - 1
    End If
    
End Sub

Private Sub txtSystem_GotFocus()
    AutoSelect Controls("txtSystem")
End Sub

Private Sub txtSystem_LostFocus()
   txtSystem.Text = oStore.SystemName
End Sub



Private Sub txtBillAddress_Change()
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oStore.SetBillAddress txtBillAddress
    If Err Then
      Beep
      intPos = txtBillAddress.SelStart
      txtBillAddress = oStore.BillAddress
      txtBillAddress.SelStart = intPos - 1
    End If
    
End Sub
Private Sub txtDelAddress_Change()
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oStore.SetDelAddress txtDelAddress
    If Err Then
      Beep
      intPos = txtDelAddress.SelStart
      txtDelAddress = oStore.DelAddress
      txtDelAddress.SelStart = intPos - 1
    End If
    
End Sub
Private Sub txtCode_Change()
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oStore.SetCode txtCode
    If Err Then
      Beep
      intPos = txtCode.SelStart
      txtCode = oStore.code
      txtCode.SelStart = intPos - 1
    End If
    
End Sub
Private Sub txtCode_LostFocus()
   txtCode.Text = oStore.code
End Sub

Private Sub txtBillAddress_GotFocus()
    AutoSelect Controls("txtBillAddress")
End Sub

Private Sub txtBillAddress_LostFocus()
   txtBillAddress.Text = oStore.BillAddress
End Sub

Private Sub txtDelAddress_GotFocus()
    AutoSelect Controls("txtDelAddress")
End Sub

Private Sub txtDelAddress_LostFocus()
   txtDelAddress.Text = oStore.DelAddress
End Sub

Private Sub txtSysten_Change()

End Sub
