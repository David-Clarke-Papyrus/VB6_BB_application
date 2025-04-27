VERSION 5.00
Begin VB.Form frmReturnRejection 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Return rejection form"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtProduct 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   690
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   135
      Width           =   4275
   End
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
      Picture         =   "frmReturnRejection.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1000
   End
   Begin VB.TextBox txtReason 
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
      Height          =   750
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1785
      Width           =   4230
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
      Picture         =   "frmReturnRejection.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1000
   End
   Begin VB.TextBox txtQty 
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
      Left            =   3525
      TabIndex        =   0
      Top             =   1005
      Width           =   825
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      Caption         =   "Note"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1515
      Width           =   1800
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      Caption         =   "Total qty rejected for this row"
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
      Left            =   465
      TabIndex        =   3
      Top             =   1020
      Width           =   2985
   End
End
Attribute VB_Name = "frmReturnRejection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mQty As Long
Dim mReason As String
Dim mPRLID As Long
Dim mReturned As Long
Dim bCancel As Boolean

Public Sub component(pTitle As String, pRLID As Long, pQtyRejected As Long, pReturned As Long, pNote As String)
    On Error GoTo errHandler
    txtProduct = pTitle
    mPRLID = pRLID
    mReturned = pReturned
    mReason = pNote
    mQty = pQtyRejected
    txtReason = mReason
    txtQty = CStr(mQty)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnRejection.component(pTitle,pRLID,pQtyRejected,pReturned,pNote)", Array(pTitle, _
         pRLID, pQtyRejected, pReturned, pNote)
End Sub
Public Property Get IsCancelled() As Boolean
    On Error GoTo errHandler
    IsCancelled = bCancel
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnRejection.IsCancelled"
End Property

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancel = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnRejection.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo errHandler
Dim oSM As z_StockManager
    Set oSM = New z_StockManager
    oSM.RejectReturn mQty, mPRLID, mReason
    bCancel = False
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnRejection.cmdSelect_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim lngTmp As Long

    If Not ConvertToLng(txtQty, lngTmp) Then
        Cancel = True
    Else
        mQty = CLng(txtQty)
        If mQty > mReturned Then
            MsgBox "You cannot reject more than you returned.", vbInformation, "Can't do this"
            Cancel = True
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnRejection.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtReason_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    mReason = Trim(txtReason)
    If Len(mReason) < 2 Then
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnRejection.txtReason_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
