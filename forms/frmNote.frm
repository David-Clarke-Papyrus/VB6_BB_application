VERSION 5.00
Begin VB.Form frmNote 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Document memo"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   Icon            =   "frmNote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4590
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
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
      Left            =   270
      Picture         =   "frmNote.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1485
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
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
      Left            =   3015
      Picture         =   "frmNote.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1000
   End
   Begin VB.TextBox txtMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   450
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Document memo"
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
      Left            =   255
      TabIndex        =   1
      Top             =   150
      Width           =   1935
   End
End
Attribute VB_Name = "frmNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flgLoading As Boolean

Public Sub component(pMemo As String)
    On Error GoTo errHandler
    txtMemo = pMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmNote.component(pMemo)", pMemo
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmNote.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Public Property Get Memo() As String
    On Error GoTo errHandler
    
    'Memo = stripCRLF(Trim(txtMemo))
    Memo = (Trim(txtMemo))
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmNote.Memo"
End Property



Private Sub Command1_Click()
    On Error GoTo errHandler
Dim f As New frmFindTextBite
    f.Show vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmNote.Command1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtMemo_Change()
    On Error GoTo errHandler
    txtMemo = HandleTextWithBites(txtMemo)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmNote.txtMemo_Change", , EA_NORERAISE
    HandleError
End Sub

'Private Sub txtMemo_LostFocus()
'Dim strarg As String
'Dim iStart As Integer
'Dim iEnd As Integer
'Dim oU As New z_UTIL
'Dim strResult As String
'Dim f As frmFindTextBite
'    strResult = ""
'    iStart = InStr(1, txtMemo, "$$$") + 3
'    iEnd = InStr(iStart, txtMemo, " ")
'    If iEnd = 0 Then iEnd = Len(txtMemo)
'    If iStart > 0 Then
'        strarg = Mid(txtMemo, iStart, iEnd)
'        strResult = oU.GetTextBite(strarg)
'    Else
'        iStart = InStr(1, txtMemo, "$$")
'        If iStart > 0 Then
'            f.Show
'        End If
'    End If
'    txtMemo = Replace(txtMemo, "$$$" & strarg, strResult)
'
'
'End Sub
Private Sub txtMemo_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If InStr(1, txtMemo, Chr(13)) > 0 Then
        If MsgBox("There are multiple lines in the text you are saving.", vbExclamation + vbOKCancel, "Warning") = vbCancel Then
            Cancel = True
            Exit Sub
        End If
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmNote.txtMemo_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
