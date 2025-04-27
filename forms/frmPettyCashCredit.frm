VERSION 5.00
Begin VB.Form frmPettyCashCredit 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Petty cash return"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   6705
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lTypes 
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
      Height          =   1740
      ItemData        =   "frmPettyCashCredit.frx":0000
      Left            =   240
      List            =   "frmPettyCashCredit.frx":0002
      TabIndex        =   0
      Top             =   150
      Width           =   6180
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
      Left            =   2010
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3825
      Width           =   1260
   End
   Begin VB.TextBox txtAmount 
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
      Left            =   2010
      MaxLength       =   20
      TabIndex        =   2
      Top             =   3165
      Width           =   2700
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
      Left            =   3405
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1260
   End
   Begin VB.TextBox txtReason 
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
      Left            =   1455
      MaxLength       =   100
      TabIndex        =   1
      Top             =   2370
      Width           =   3750
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   330
      Left            =   2835
      TabIndex        =   5
      Top             =   2865
      Width           =   1005
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason for cash return"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   330
      Left            =   1710
      TabIndex        =   4
      Top             =   2025
      Width           =   3195
   End
End
Attribute VB_Name = "frmPettyCashCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim strReason As String
Dim strAmount As String
Dim arPC() As String
Private Sub cmdCancel_Click()
    bCancelled = True
    Me.Hide
End Sub
Private Sub cmdOK_Click()
    If lTypes.SelCount > 0 And strAmount > "" Then
        Me.Hide
    Else
        MsgBox "Form incomplete.", vbInformation, "Can't do this"
    End If
End Sub
Public Sub Component(pPCExchanges() As String)
Dim i As Integer
Dim arPC() As String

    For i = 1 To UBound(pPCExchanges)
        arPC = Split(pPCExchanges(i), "|")
        Me.lTypes.AddItem arPC(2) & " : " & arPC(1) & " : " & arPC(4) & " : " & arPC(3)
    Next
    Me.lTypes.AddItem "<Deposit to petty cash - no corresponding withdrawal>"
End Sub


Private Sub lTypes_Validate(Cancel As Boolean)
    cmdOK.Enabled = FormComplete
End Sub

Private Sub txtAmount_Validate(Cancel As Boolean)
    cmdOK.Enabled = FormComplete
End Sub
Private Sub txtReason_Change()
    strReason = txtReason
End Sub
Private Sub txtAmount_Change()
    strAmount = txtAmount
End Sub
Public Property Get Reason() As String
    Reason = lTypes & ":" & Replace(strReason, vbTab, "")
End Property
Public Property Get Amount() As String
    Amount = strAmount
End Property
Private Sub txtAmount_GotFocus()
    AutoSelect Controls("txtAmount")
End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property
Private Sub txtReason_Validate(Cancel As Boolean)
    cmdOK.Enabled = FormComplete
End Sub
Function FormComplete() As Boolean
    FormComplete = (Len(txtReason) > 3) And (txtReason > "") And (IsNumeric(txtAmount)) And (txtAmount > "") And lTypes.SelCount > 0
End Function

