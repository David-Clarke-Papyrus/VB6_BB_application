VERSION 5.00
Begin VB.Form frmOrderStatusReport 
   Caption         =   "Order status report"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDiarize 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   285
      TabIndex        =   1
      Top             =   1020
      Width           =   840
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   510
      Left            =   2490
      Picture         =   "frmOrderStatusReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1065
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   510
      Left            =   3495
      Picture         =   "frmOrderStatusReport.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1065
      Width           =   1000
   End
   Begin VB.TextBox txtNote 
      Height          =   285
      Left            =   285
      TabIndex        =   0
      Top             =   330
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Diarize . . . e.g. 2w,4w,2m etc"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   300
      TabIndex        =   5
      Top             =   810
      Width           =   2160
   End
   Begin VB.Label Label1 
      Caption         =   "Note"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   330
      TabIndex        =   2
      Top             =   105
      Width           =   2160
   End
End
Attribute VB_Name = "frmOrderStatusReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strNote As String
Dim strDiarize As String

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    strNote = ""
    strDiarize = ""
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderStatusReport.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderStatusReport.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub txtDiarize_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim s As String
Dim bOK As Boolean

    If txtDiarize = "" Then
        strDiarize = ""
        Exit Sub
    End If
    txtDiarize = UCase(txtDiarize)
    s = TranslateDiaryPeriods(txtDiarize, bOK)
    If bOK Then
        strDiarize = txtDiarize
    Else
        MsgBox "Invalid format", vbInformation + vbOKOnly, "Warning"
        Cancel = True
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderStatusReport.txtDiarize_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim s As String
Dim bOK As Boolean

'    If txtNote = "" Then
'        strNote = ""
'        Exit Sub
'    End If
'
'    s = TranslateDiaryPeriods(txtNote, bOK)
'    If bOK Then
        strNote = txtNote
'    Else
'        MsgBox "Invalid format", vbInformation + vbOKOnly, "Warning"
'        Cancel = True
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderStatusReport.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Public Property Get Note() As String
    Note = strNote
End Property
Public Property Get Diarize() As String
    Diarize = strDiarize
End Property

