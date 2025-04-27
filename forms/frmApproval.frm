VERSION 5.00
Begin VB.Form frmApproval 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Return approval form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   2310
      Picture         =   "frmApproval.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2295
      Width           =   1000
   End
   Begin VB.TextBox txtApprovalDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2805
      TabIndex        =   1
      Top             =   1005
      Width           =   1530
   End
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
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
      Left            =   3360
      Picture         =   "frmApproval.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2295
      Width           =   1000
   End
   Begin VB.TextBox txtApprovalRef 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2805
      TabIndex        =   0
      Top             =   495
      Width           =   1530
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      Caption         =   "Approval expiry date"
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
      Height          =   255
      Left            =   15
      TabIndex        =   4
      Top             =   1020
      Width           =   2580
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      Caption         =   "Supplier's approval number"
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
      Height          =   255
      Left            =   30
      TabIndex        =   3
      Top             =   525
      Width           =   2580
   End
End
Attribute VB_Name = "frmApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mApprovalRef As String
Dim mApprovalDate As Date
Dim bCancel As Boolean

Public Property Get IsCancelled() As Boolean
    On Error GoTo errHandler
    IsCancelled = bCancel
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmApproval.IsCancelled"
End Property

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancel = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmApproval.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo errHandler
    bCancel = False
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmApproval.cmdSelect_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtApprovalDate_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If IsDate(txtApprovalDate) Then
        mApprovalDate = CDate(txtApprovalDate)
        If mApprovalDate < Date Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmApproval.txtApprovalDate_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtApprovalRef_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    mApprovalRef = FNS(txtApprovalRef)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmApproval.txtApprovalRef_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Public Property Get ApprovalRef() As String
    ApprovalRef = mApprovalRef
End Property
Public Property Get ApprovalDate() As Date
    ApprovalDate = mApprovalDate
End Property

