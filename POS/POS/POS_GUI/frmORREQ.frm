VERSION 5.00
Begin VB.Form frmORREQ 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Order request"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   8805
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDep3 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00714942&
      Height          =   345
      Left            =   3030
      TabIndex        =   22
      Text            =   "txtItem1"
      Top             =   2715
      Width           =   855
   End
   Begin VB.TextBox txtDep2 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00714942&
      Height          =   345
      Left            =   3030
      TabIndex        =   21
      Text            =   "txtItem1"
      Top             =   2310
      Width           =   855
   End
   Begin VB.TextBox txtDep1 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00714942&
      Height          =   345
      Left            =   3030
      TabIndex        =   20
      Text            =   "txtItem1"
      Top             =   1890
      Width           =   855
   End
   Begin VB.TextBox txtItem3 
      ForeColor       =   &H00714942&
      Height          =   345
      Left            =   705
      TabIndex        =   17
      Text            =   "txtItem3"
      Top             =   2700
      Width           =   1830
   End
   Begin VB.CommandButton cmdSelectItem3 
      Height          =   315
      Left            =   2610
      Picture         =   "frmORREQ.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2730
      Width           =   345
   End
   Begin VB.TextBox txtItem2 
      ForeColor       =   &H00714942&
      Height          =   345
      Left            =   705
      TabIndex        =   13
      Text            =   "txtItem2"
      Top             =   2295
      Width           =   1830
   End
   Begin VB.CommandButton cmdSelectItem2 
      Height          =   315
      Left            =   2610
      Picture         =   "frmORREQ.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2325
      Width           =   345
   End
   Begin VB.CommandButton cmdSelectItem1 
      Height          =   315
      Left            =   2610
      Picture         =   "frmORREQ.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      Width           =   345
   End
   Begin VB.TextBox txtItem1 
      ForeColor       =   &H00714942&
      Height          =   345
      Left            =   705
      TabIndex        =   8
      Text            =   "txtItem1"
      Top             =   1890
      Width           =   1830
   End
   Begin VB.TextBox txtItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   750
      IMEMode         =   3  'DISABLE
      Left            =   690
      MaxLength       =   350
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3975
      Width           =   3180
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00DACDCD&
      Cancel          =   -1  'True
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
      Height          =   465
      Left            =   1185
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5295
      Width           =   1260
   End
   Begin VB.TextBox txtDeposit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   4950
      MaxLength       =   20
      TabIndex        =   2
      Top             =   3975
      Width           =   1755
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00DACDCD&
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
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
      Left            =   2595
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5310
      Width           =   1260
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   990
      IMEMode         =   3  'DISABLE
      Left            =   135
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   585
      Width           =   7500
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Product code"
      ForeColor       =   &H00714942&
      Height          =   255
      Left            =   1035
      TabIndex        =   24
      Top             =   1650
      Width           =   1245
   End
   Begin VB.Label lblDep 
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit (e.g. 45.99)"
      ForeColor       =   &H00714942&
      Height          =   210
      Left            =   2955
      TabIndex        =   23
      Top             =   1650
      Width           =   1620
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H00714942&
      Height          =   255
      Left            =   135
      TabIndex        =   19
      Top             =   2760
      Width           =   510
   End
   Begin VB.Label lblItem3 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00714942&
      Height          =   240
      Left            =   3990
      TabIndex        =   18
      Top             =   2760
      Width           =   3645
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      ForeColor       =   &H00714942&
      Height          =   255
      Left            =   135
      TabIndex        =   15
      Top             =   2355
      Width           =   510
   End
   Begin VB.Label lblItem2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00714942&
      Height          =   240
      Left            =   3990
      TabIndex        =   14
      Top             =   2355
      Width           =   3645
   End
   Begin VB.Label lblItem1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00714942&
      Height          =   240
      Left            =   3990
      TabIndex        =   11
      Top             =   1950
      Width           =   3645
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00714942&
      Height          =   255
      Left            =   135
      TabIndex        =   10
      Top             =   1950
      Width           =   510
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   330
      Left            =   75
      TabIndex        =   7
      Top             =   4035
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   330
      Left            =   3960
      TabIndex        =   6
      Top             =   4005
      Width           =   1005
   End
   Begin VB.Label lblMsg 
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Details of customer"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   330
      Left            =   135
      TabIndex        =   5
      Top             =   240
      Width           =   3195
   End
End
Attribute VB_Name = "frmORREQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim strCustomer As String
Dim strItem As String
Dim strDeposit As String
Private xMLDoc As ujXML

Private Sub cmdCancel_Click()
    bCancelled = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If CLng(strDeposit) < 1000 Then
        If MsgBox("You are accepting an unusually low deposit. Do you wish to continue?", vbQuestion + vbYesNo, "Please check deposit value") = vbNo Then
            Exit Sub
        End If
    End If
    Me.Hide
End Sub



Private Sub cmdSelectItem1_Click()
Dim f As New frmQuickProductFind

    f.component "<Enter search string>"
    f.Show vbModal
    txtItem1 = f.EAN
    Me.lblItem1 = f.Description
    Unload f
    
End Sub
Private Sub cmdSelectItem2_Click()
Dim f As New frmQuickProductFind

    f.component "<Enter search string>"
    f.Show vbModal
    txtItem2 = f.EAN
    Me.lblItem2 = f.Description
    Unload f
    
End Sub
Private Sub cmdSelectItem3_Click()
Dim f As New frmQuickProductFind

    f.component "<Enter search string>"
    f.Show vbModal
    txtItem3 = f.EAN
    Me.lblItem3 = f.Description
    Unload f
    
End Sub

Private Sub Form_Load()
Dim arType() As String
Dim i As Integer

    txtDeposit = CStr(oPC.DefaultDeposit)
End Sub

Private Sub txtDeposit_Validate(Cancel As Boolean)
    Cancel = (Not (IsNumeric(strDeposit)))
    CheckOKStatus
End Sub
Private Sub txtDeposit_Change()
    strDeposit = txtDeposit
End Sub
Public Property Get DepositF() As String
    DepositF = strDeposit
End Property

Public Property Get Deposit() As Long
    Deposit = CLng(strDeposit)
End Property

Private Sub txtDeposit_GotFocus()
    AutoSelect Controls("txtDeposit")
End Sub




Private Sub txtCustomer_Change()
   strCustomer = txtCustomer
End Sub
Public Property Get Customer() As String
   Customer = Replace(strCustomer, vbTab, "")
End Property
Private Sub txtCustomer_Validate(Cancel As Boolean)
  '  Cancel = (Not (Len(txtCustomer) > 10))
    CheckOKStatus
End Sub


Private Sub txtItem_Change()
   strItem = txtItem
End Sub
Public Property Get Item() As String
   Item = Replace(strItem, vbTab, "")
End Property
Private Sub txtItem_Validate(Cancel As Boolean)
  '  Cancel = (Not (Len(txtItem) > 6))
    CheckOKStatus
End Sub
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property

Private Sub CheckOKStatus()
    If Len(Me.txtItem) > 6 And Len(txtCustomer) > 10 Then
        Me.cmdOK.Enabled = True
    End If
End Sub

Public Function GetDetailsXML() As String
    Set xMLDoc = New ujXML
    
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "OR_1"
            .chCreate "MessageType"
                .elText = "ORDER REQUEST"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "CustomerDetails"
                   .elText = Me.txtCustomer
            .elCreateSibling "Notes"
                .elText = Me.txtItem
            .elCreateSibling "Deposit"
                .elText = txtDeposit
            .elCreateSibling "ItemList"
            If txtItem1 > "" Then
                .chCreate "Item"
                .chCreate "EAN", True
                    .elText = txtItem1
                .elCreateSibling "DEP", True
                    .elText = txtDep1
                .navUP
                .navUP
            End If
            If txtItem2 > "" Then
                .chCreate "Item"
                .chCreate "EAN", True
                    .elText = txtItem2
                .elCreateSibling "DEP", True
                    .elText = txtDep2
                .navUP
                .navUP
            End If
            If txtItem3 > "" Then
                .chCreate "Item"
                .chCreate "EAN", True
                    .elText = txtItem3
                .elCreateSibling "DEP", True
                    .elText = txtDep3
                .navUP
            End If
            .navUP
    End With
    GetDetailsXML = xMLDoc.docXML
End Function
