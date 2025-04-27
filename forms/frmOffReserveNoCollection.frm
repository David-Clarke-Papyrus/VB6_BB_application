VERSION 5.00
Begin VB.Form frmOffReserveNoCollection 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Return to stock from reserve"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDetails 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   405
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1245
      Width           =   3135
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
      Left            =   1590
      TabIndex        =   0
      Top             =   180
      Width           =   615
   End
   Begin VB.CommandButton cmdNo 
      BackColor       =   &H00D7D1BF&
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1935
      Width           =   1170
   End
   Begin VB.CommandButton cmdYes 
      BackColor       =   &H00D7D1BF&
      Caption         =   "Yes"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   660
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1170
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copies have NOT been collected and must be taken off reserve."
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
      Height          =   735
      Left            =   375
      TabIndex        =   3
      Top             =   645
      Width           =   3360
   End
End
Attribute VB_Name = "frmOffReserveNoCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strChoice As String
Dim strMsg As String
Dim lngQty As Long


Private Sub cmdNo_Click()
    On Error GoTo errHandler
    strChoice = "No"
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOffReserveNoCollection.cmdNo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdYes_Click()
    On Error GoTo errHandler
    strChoice = "Yes"
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOffReserveNoCollection.cmdYes_Click", , EA_NORERAISE
    HandleError
End Sub

Public Property Get GetChoice() As String
    GetChoice = strChoice
End Property
Public Property Get GetQty() As Long
    GetQty = lngQty
End Property


Public Sub component(pDetails As String, pQty As Long)
    On Error GoTo errHandler
    txtDetails = pDetails
    lngQty = pQty
    txtQty = lngQty
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOffReserveNoCollection.component(pDetails,pQty)", Array(pDetails, pQty)
End Sub


Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If Not ConvertToLng(txtQty, lngQty) Then
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOffReserveNoCollection.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
