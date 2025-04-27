VERSION 5.00
Begin VB.Form frmRediarize 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Rediarize"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3930
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkReprinting 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Reprinting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   1380
      TabIndex        =   4
      Top             =   195
      Width           =   1710
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
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
      Left            =   1995
      Picture         =   "frmRediarize.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1740
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   390
      Left            =   1485
      TabIndex        =   1
      Text            =   "1M"
      Top             =   720
      Width           =   915
   End
   Begin VB.CommandButton cmdRediarize 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Rediarize"
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
      Left            =   960
      Picture         =   "frmRediarize.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1740
      Width           =   1000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "e.g. 2W = 2 weeks; 3M = 3 months"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   750
      TabIndex        =   2
      Top             =   1275
      Width           =   2385
   End
End
Attribute VB_Name = "frmRediarize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iNumber As Integer
Dim strUnit As String
Dim bCancel As Boolean


Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRediarize.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancel = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRediarize.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRediarize_Click()
    On Error GoTo errHandler
    If Text1 = "" Then
        MsgBox "You have not entered a rediarized period. Either enter one or Cancel."
        Exit Sub
    End If
    bCancel = False
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRediarize.cmdRediarize_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Text1_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim strNumber As String
Dim strUnit As String
    Text1 = Trim(Text1)
    If Text1 = "" Then
        Exit Sub
    End If
    strNumber = Left(Text1, Len(Text1) - 1)
    strUnit = UCase(Right(Text1, 1))
    If Not IsNumeric(strNumber) Then
        Cancel = True
    Else
        iNumber = CInt(strNumber)
        If iNumber < 1 Then
            Cancel = True
        End If
    End If
    If Not (strUnit = "W" Or strUnit = "M") Then
        Cancel = True
    End If
        
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRediarize.Text1_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = bCancel
End Property

Public Property Get RediarizedPeriod() As String
    On Error GoTo errHandler
    RediarizedPeriod = Trim(Text1)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRediarize.RediarizedPeriod"
End Property
Public Property Get Reason() As String
    On Error GoTo errHandler
    If Me.chkReprinting = 1 Then
        Reason = "R"
    Else
        Reason = ""
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmRediarize.Reason"
End Property

