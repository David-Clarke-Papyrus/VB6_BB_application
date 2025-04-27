VERSION 5.00
Begin VB.Form frmSelectInvoiceDispatchSheet 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2670
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   1725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Height          =   300
      Left            =   570
      Picture         =   "frmSelectInvoiceDispatchSheet.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2190
      Width           =   540
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      ForeColor       =   &H8000000D&
      Height          =   1245
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   1500
      Begin VB.OptionButton optSelected 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Selected only"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1320
      End
      Begin VB.OptionButton optThisMonth 
         BackColor       =   &H00D3D3CB&
         Caption         =   "This month"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   870
         Width           =   1125
      End
      Begin VB.OptionButton optThisWeek 
         BackColor       =   &H00D3D3CB&
         Caption         =   "This week"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   630
         Width           =   1125
      End
      Begin VB.OptionButton optThisDay 
         BackColor       =   &H00D3D3CB&
         Caption         =   "This day"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   375
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmSelectInvoiceDispatchSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sRangeSelected As String

Public Sub component(TOP As Long, Left As Long)
    On Error GoTo errHandler
    Me.TOP = TOP
    Me.Left = Left
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSelectInvoiceDispatchSheet.component(top,Left)", Array(TOP, Left)
End Sub
Private Sub cmdGo_Click()
    On Error GoTo errHandler
        Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSelectInvoiceDispatchSheet.cmdGo_Click", , EA_NORERAISE
    HandleError
End Sub

Public Property Get RangeSelected() As String
    On Error GoTo errHandler
    If optSelected = True Then
        sRangeSelected = "Selected"
    ElseIf optThisDay = True Then
        sRangeSelected = "ThisDay"
    ElseIf optThisWeek = True Then
        sRangeSelected = "ThisWeek"
    Else
        sRangeSelected = "ThisMonth"
    End If
    RangeSelected = sRangeSelected
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSelectInvoiceDispatchSheet.RangeSelected"
End Property

