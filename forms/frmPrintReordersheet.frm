VERSION 5.00
Begin VB.Form frmPrintReordersheet 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Print reorder sheet"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3930
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSummary 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Summary"
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
      Height          =   240
      Left            =   510
      TabIndex        =   3
      Top             =   810
      Value           =   1  'Checked
      Width           =   2910
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
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
      Left            =   1320
      Picture         =   "frmPrintReordersheet.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3090
      Width           =   1000
   End
   Begin VB.Frame frmSort 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Sort by "
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
      Height          =   1425
      Left            =   510
      TabIndex        =   1
      Top             =   1350
      Width           =   2760
      Begin VB.PictureBox Picture 
         BackColor       =   &H00D3D3CB&
         Height          =   990
         Left            =   165
         ScaleHeight     =   930
         ScaleWidth      =   2445
         TabIndex        =   4
         Top             =   270
         Width           =   2505
         Begin VB.OptionButton optSuppliername 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Supplier name"
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
            Height          =   195
            Left            =   135
            TabIndex        =   6
            Top             =   120
            Value           =   -1  'True
            Width           =   2145
         End
         Begin VB.OptionButton optDescription 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Product description"
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
            Height          =   195
            Left            =   135
            TabIndex        =   5
            Top             =   600
            Width           =   2145
         End
      End
   End
   Begin VB.CheckBox chkFilter 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Only ordered products"
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
      Height          =   240
      Left            =   510
      TabIndex        =   0
      Top             =   345
      Value           =   1  'Checked
      Width           =   2910
   End
End
Attribute VB_Name = "frmPrintReordersheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOrderedOnly As Boolean
Dim strSequence As String
Public Sub component(sCaption As String, sButtonCaption As String)
    If sCaption > "" Then
        Me.Caption = sCaption
    End If
    If sButtonCaption > "" Then
        Me.cmdPrint.Caption = sButtonCaption
    End If

End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    If Me.optDescription Then
        strSequence = "DESCRIP"
    Else
        strSequence = "LASTSUPPLIERNAME"
    End If
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintReordersheet.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Public Property Get Sequence() As String
    On Error GoTo errHandler
    Sequence = strSequence
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPrintReordersheet.Sequence"
End Property
Public Property Get OrderedOnly() As Boolean
    OrderedOnly = (chkFilter = 1)
End Property
