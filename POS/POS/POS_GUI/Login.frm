VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1875
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   125
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080C0FF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2145
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080C0FF&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   690
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1140
   End
   Begin VB.ComboBox cboName 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   360
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   210
      Width           =   2400
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   615
      Width           =   2400
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Index           =   1
      Left            =   135
      TabIndex        =   2
      Top             =   630
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Canceled As Boolean
Public SalesPerson As String

Dim gPassword As String
Dim gUserName As String
Dim bZOp As Boolean

Public Sub Component(Optional IsZOperator As Boolean) '(sUserName As String)
    On Error GoTo errHandler
'    Dim i As Integer
    bZOp = IsZOperator
    If bZOp Then
        Me.Caption = "Login for Z Action cleared Operator."
    End If
    LoadCombo oGD.SalesPersonList, Me.cboName
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLogin.Component(IsZOperator)", IsZOperator
End Sub

Private Sub cboName_Click()
    On Error GoTo errHandler
    If cboName = "ADM" Then
        Me.txtPassword = "admin"
    Else
        Me.txtPassword = ""
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLogin.cboName_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
Dim msg As String
    If bZOp Then
        msg = "Z Total Action can not be accessed without a valid Username / Password!"
    Else
        msg = "Application will be locked without valid Usaername / Password!"
    End If
    If MsgBox(msg & vbLf & _
              "Cancel anyway?", vbYesNo + vbExclamation, "WARNING!") = vbNo Then
      Exit Sub
    End If
    Canceled = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLogin.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long
Dim msg As String
Dim iOff As Integer
    
    If Len(Me.txtPassword) = 0 Or Len(Me.cboName.Text) = 0 Then GoTo NoJoy
    If Not oGD.CheckPassword(Me.cboName.Text, Me.txtPassword) Then
        'Handle general login
        msg = "UserName or Password is not valid!" & vbLf & _
              "Try again?"
        GoTo NoJoy
    ElseIf bZOp Then
        'Handle ZAction login
        If Not oGD.IsValidZOpPass(Me.txtPassword, ZACTION_LEVEL) Then
            msg = "Access to Z Action not cleared!" & vbLf & _
                  "Try again?"
            GoTo NoJoy
        End If
    End If
    iOff = InStr(Me.cboName.Text, "(")
    SalesPerson = Mid(Me.cboName.Text, iOff + 1, 4)
    If InStr(SalesPerson, ")") > 0 Then
      SalesPerson = Left(SalesPerson, Len(SalesPerson) - 1)
    End If
    Screen.MousePointer = vbDefault
    

 
EXIT_Handler:
    Me.Hide
    Exit Sub
NoJoy:
    If MsgBox(msg, vbYesNo + vbCritical, "Password not valid") = vbYes Then
        Exit Sub
    Else
        Canceled = True
        GoTo EXIT_Handler
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLogin.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
  'Place under splash form
  Left = (Screen.Width - Width) / 2
  Top = (Screen.Height - Height) / 2 + 2000
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLogin.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Public Property Get Password() As String
  Password = gPassword
End Property

Public Property Get Username() As String
  Username = gUserName
End Property

Private Sub txtPassword_GotFocus()
    On Error GoTo errHandler
    With txtPassword
    If Len(.Text) > 0 Then
      .SelStart = 0
      .SelLength = Len(.Text)
    End If
  End With

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLogin.txtPassword_GotFocus", , EA_NORERAISE
    HandleError
End Sub
