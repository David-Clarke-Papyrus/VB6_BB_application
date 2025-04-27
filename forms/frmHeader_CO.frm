VERSION 5.00
Begin VB.Form frmHeader_CO 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Sales order details"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4515
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6195
      Width           =   135
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
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
      Height          =   615
      Left            =   3225
      Picture         =   "frmHeader_CO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2010
      Width           =   1000
   End
   Begin VB.TextBox txtMemo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   225
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1110
      Width           =   4005
   End
   Begin VB.TextBox txtRef 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   210
      MaxLength       =   20
      TabIndex        =   0
      Top             =   435
      Width           =   4035
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "(Click ESC to cancel)"
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
      Height          =   240
      Left            =   195
      TabIndex        =   6
      Top             =   1965
      Width           =   1800
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Memo"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   270
      TabIndex        =   4
      Top             =   885
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer order reference"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   270
      TabIndex        =   3
      Top             =   195
      Width           =   2850
   End
End
Attribute VB_Name = "frmHeader_CO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strRef As String
Dim strMemo As String
Dim flgLoading As Boolean
Dim oCO As a_CO
Dim bCancel As Boolean

'Public Sub component(pOrderRef As String, pMemo As String)
'    strRef = pOrderRef
'    strMemo = pMemo
'End Sub
'
Public Property Get Cancelled() As Boolean
    Cancelled = bCancel
End Property

Public Sub component(pCO As a_CO)
    On Error GoTo errHandler
    Set oCO = pCO
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_CO.component(pCO)", pCO
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancel = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_CO.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_CO.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    txtMemo = oCO.Memo
    strMemo = oCO.Memo
    txtRef = oCO.OrderRef
    strRef = oCO.OrderRef
    CheckCloseButton
    bCancel = False
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_CO.Form_Load", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtRef_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtRef
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_CO.txtRef_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtRef_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
If KeyCode = 27 Then
    Set oCO = Nothing
    bCancel = True
    Me.Hide
End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_CO.txtRef_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
If KeyAscii = 27 Then
    Set oCO = Nothing
    bCancel = True
    Me.Hide
End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_CO.txtRef_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub txtRef_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim oSQL As New z_SQL
Dim strResult As String

    If flgLoading Then Exit Sub
    If Len(txtRef) <= 2 And txtRef <> "" Then
        Cancel = True
    Else
        Cancel = False
        oCO.OrderRef = txtRef
        strRef = txtRef
    End If
    If oPC.GetProperty("CheckRefsOnCO") = "TRUE" Then
        oSQL.FindCORefMatch oCO.OrderRef, strResult
        If strResult > "" Then
            If MsgBox("This reference number is found on document(s): " & strResult & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
                Cancel = True
            End If
        End If
    End If
    CheckCloseButton
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_CO.txtRef_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub CheckCloseButton()
    On Error GoTo errHandler
    cmdclose.Enabled = (Len(txtRef) > 2) Or txtRef = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_CO.CheckCloseButton"
End Sub
Public Property Get Ref() As String
    Ref = strRef
End Property

Public Property Get Memo() As String
    Memo = strMemo
End Property
Private Sub txtMemo_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Len(txtMemo) <= 2 And txtMemo <> "" Then
        Cancel = True
    Else
        Cancel = False
        strMemo = txtMemo
        oCO.SetMemo strMemo
    End If
    CheckCloseButton
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader_CO.txtMemo_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
